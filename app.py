#!/usr/bin/env python3
"""
台股均量分析工具 v3 - 雲端部署版
使用背景執行緒 + 輪詢，避免雲端環境 SSE timeout 問題
"""

import requests
import time
import re
import json
import threading
import uuid
from datetime import datetime
from flask import Flask, jsonify, send_from_directory, request
import yfinance as yf
import pandas as pd
import os

app = Flask(__name__, static_folder='static')

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
    'Accept': 'application/json, text/plain, */*',
    'Accept-Language': 'zh-TW,zh;q=0.9',
}

# ── 任務狀態儲存（記憶體）──
tasks = {}  # task_id -> task_state


# ──────────────────────────────────────────────
# TWSE + TPEX 資料取得
# ──────────────────────────────────────────────
def get_twse_stocks():
    stocks = []
    try:
        url = "https://www.twse.com.tw/rwd/zh/afterTrading/STOCK_DAY_ALL"
        resp = requests.get(url, params={"response": "json"}, headers=HEADERS, timeout=20)
        resp.raise_for_status()
        data = resp.json()
        for row in data.get('data', []):
            try:
                code = row[0].strip()
                name = row[1].strip()
                if not re.match(r'^\d{4}$', code):
                    continue
                close  = float(row[7].replace(',', '').strip())
                change = float(row[8].replace(',', '').strip())
                prev   = close - change
                if prev <= 0:
                    continue
                change_pct = round((change / prev) * 100, 2)
                if change_pct <= 0:
                    continue
                vol = int(row[2].replace(',', ''))
                stocks.append({
                    'code': code, 'symbol': f"{code}.TW", 'name': name,
                    'change_pct': change_pct, 'price': close,
                    'volume': vol, 'market': '上市',
                })
            except Exception:
                continue
        print(f"[TWSE] 上漲個股：{len(stocks)} 支")
    except Exception as e:
        print(f"[TWSE] 失敗：{e}")
    return stocks


def get_tpex_stocks():
    stocks = []
    try:
        url = "https://www.tpex.org.tw/openapi/v1/tpex_mainboard_quotes"
        resp = requests.get(url, headers=HEADERS, timeout=20)
        resp.raise_for_status()
        data = resp.json()
        for item in data:
            try:
                code = str(item.get('SecuritiesCompanyCode', '')).strip()
                name = str(item.get('CompanyName', '')).strip()
                if not re.match(r'^\d{4,5}$', code):
                    continue
                close  = float(str(item.get('Close',  '0')).replace(',', '') or 0)
                change = float(str(item.get('Change', '0')).replace(',', '') or 0)
                if close <= 0:
                    continue
                prev = close - change
                if prev <= 0:
                    continue
                change_pct = round((change / prev) * 100, 2)
                if change_pct <= 0:
                    continue
                vol_str = str(item.get('TradeVolume', '0')).replace(',', '')
                vol = int(float(vol_str)) if vol_str else 0
                stocks.append({
                    'code': code, 'symbol': f"{code}.TWO", 'name': name,
                    'change_pct': change_pct, 'price': close,
                    'volume': vol, 'market': '上櫃',
                })
            except Exception:
                continue
        print(f"[TPEX] 上漲個股：{len(stocks)} 支")
    except Exception as e:
        print(f"[TPEX] 失敗：{e}")
    return stocks


def get_ranking_stocks(top_n=100):
    """上市前100 + 上櫃前100，各自取漲幅前 top_n，合計最多 200 支"""
    twse = get_twse_stocks()
    tpex = get_tpex_stocks()
    twse.sort(key=lambda x: -x['change_pct'])
    tpex.sort(key=lambda x: -x['change_pct'])
    result = twse[:top_n] + tpex[:top_n]
    print(f"[合計] 上市前{len(twse[:top_n])}名 + 上櫃前{len(tpex[:top_n])}名 = {len(result)} 支")
    return result


def get_volume_data(symbol, days=60):
    try:
        ticker = yf.Ticker(symbol)
        hist = ticker.history(period=f"{days}d")
        if hist.empty:
            alt = symbol.replace('.TW', '.TWO') if symbol.endswith('.TW') else symbol.replace('.TWO', '.TW')
            hist = yf.Ticker(alt).history(period=f"{days}d")
            if not hist.empty:
                return hist['Volume'], alt
        return (hist['Volume'], symbol) if not hist.empty else (None, symbol)
    except Exception as e:
        return None, symbol


def analyze_volume_condition(volume_series, min_days=2, max_days=5):
    if volume_series is None or len(volume_series) < 25:
        return None
    ma5  = volume_series.rolling(window=5).mean()
    ma20 = volume_series.rolling(window=20).mean()
    combined = pd.DataFrame({'ma5': ma5, 'ma20': ma20}).dropna()
    if len(combined) == 0:
        return None
    condition = combined['ma5'] > combined['ma20']
    consecutive = 0
    for val in reversed(condition.values):
        if val:
            consecutive += 1
        else:
            break
    if not (min_days <= consecutive <= max_days):
        return None
    latest_ma5  = round(combined['ma5'].iloc[-1])
    latest_ma20 = round(combined['ma20'].iloc[-1])
    ratio = round(latest_ma5 / latest_ma20, 3) if latest_ma20 > 0 else 0
    return {
        'consecutive_days': consecutive,
        'ma5_volume':  int(latest_ma5),
        'ma20_volume': int(latest_ma20),
        'ratio': ratio,
    }


# ──────────────────────────────────────────────
# 背景執行緒執行分析任務
# ──────────────────────────────────────────────
def run_analysis_task(task_id):
    task = tasks[task_id]
    task['status'] = 'fetching'
    task['msg'] = '正在從 TWSE / TPEX 取得漲幅排行榜...'
    start_time = time.time()

    try:
        stocks = get_ranking_stocks(top_n=100)
        if not stocks:
            task['status'] = 'error'
            task['msg'] = '無法取得排行資料，TWSE/TPEX API 可能暫時離線'
            return

        total = len(stocks)
        task['total'] = total
        task['status'] = 'analyzing'
        task['msg'] = f'取得 {total} 支個股（上市前100 + 上櫃前100），開始分析均量...'

        qualified = []
        for i, stock in enumerate(stocks):
            task['current'] = i + 1
            task['current_code'] = stock['code']
            task['current_name'] = stock['name']

            volume_data, actual_symbol = get_volume_data(stock['symbol'])
            result = analyze_volume_condition(volume_data)
            if result:
                qualified.append({**stock, 'symbol': actual_symbol, **result})
            time.sleep(0.15)

        qualified.sort(key=lambda x: (-x['consecutive_days'], -x['ratio']))
        elapsed = round(time.time() - start_time, 1)

        task['status']    = 'done'
        task['stocks']    = qualified
        task['total_found'] = len(qualified)
        task['scanned']   = total
        task['elapsed']   = elapsed
        task['timestamp'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        task['msg']       = f'完成！找到 {len(qualified)} 支符合條件個股'

    except Exception as e:
        task['status'] = 'error'
        task['msg'] = f'分析過程發生錯誤：{str(e)}'


# ──────────────────────────────────────────────
# Flask Routes
# ──────────────────────────────────────────────
@app.route('/')
def index():
    return send_from_directory('static', 'index.html')


@app.route('/api/analyze/start', methods=['POST'])
def analyze_start():
    """建立分析任務，回傳 task_id"""
    task_id = str(uuid.uuid4())[:8]
    tasks[task_id] = {
        'status': 'pending',
        'msg': '準備開始...',
        'current': 0,
        'total': 0,
        'current_code': '',
        'current_name': '',
        'stocks': [],
        'total_found': 0,
        'scanned': 0,
        'elapsed': 0,
        'timestamp': '',
    }
    t = threading.Thread(target=run_analysis_task, args=(task_id,), daemon=True)
    t.start()
    return jsonify({'task_id': task_id})


@app.route('/api/analyze/status/<task_id>')
def analyze_status(task_id):
    """輪詢任務狀態（每 1.5 秒呼叫一次）"""
    task = tasks.get(task_id)
    if not task:
        return jsonify({'error': '找不到任務'}), 404
    return jsonify(task)


@app.route('/api/stock/<code>')
def stock_detail(code):
    volume_data, actual_symbol = get_volume_data(f"{code}.TW", days=60)
    if volume_data is None:
        volume_data, actual_symbol = get_volume_data(f"{code}.TWO", days=60)
    if volume_data is None:
        return jsonify({'error': f'無法取得 {code} 的資料'}), 404

    ma5  = volume_data.rolling(5).mean()
    ma20 = volume_data.rolling(20).mean()
    dates = [d.strftime('%Y-%m-%d') for d in volume_data.index]

    return jsonify({
        'code': code, 'symbol': actual_symbol, 'dates': dates,
        'volume': [int(v) for v in volume_data.values],
        'ma5':  [round(v) if not pd.isna(v) else None for v in ma5.values],
        'ma20': [round(v) if not pd.isna(v) else None for v in ma20.values],
    })


if __name__ == '__main__':
    os.makedirs('static', exist_ok=True)
    print("=" * 50)
    print("🚀 台股均量追蹤系統 v3 啟動（雲端部署版）")
    print("📡 資料來源：TWSE + TPEX 公開 API")
    print("🌐 請開啟瀏覽器前往：http://localhost:5001")
    print("=" * 50)
    app.run(host='0.0.0.0', port=5001, debug=False, threaded=True)
