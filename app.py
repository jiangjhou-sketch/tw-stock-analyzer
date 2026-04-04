#!/usr/bin/env python3
"""
台股均量分析工具 v2
資料來源：台灣證交所 (TWSE) + 櫃買中心 (TPEX) 公開 API
分析漲幅排行中，5日均量 > 20日均量 且連續 3~5 日的個股
"""

import requests
import time
import re
import json
from datetime import datetime
from flask import Flask, jsonify, send_from_directory, Response
import yfinance as yf
import pandas as pd
import os

app = Flask(__name__, static_folder='static')

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
    'Accept': 'application/json, text/plain, */*',
    'Accept-Language': 'zh-TW,zh;q=0.9',
}


# ─────────────────────────────────────────────
# 資料來源 1：台灣證交所 (TWSE) - 上市股票
# ─────────────────────────────────────────────
def get_twse_stocks():
    """從 TWSE 取得上市股票當日行情，回傳上漲個股"""
    stocks = []
    try:
        url = "https://www.twse.com.tw/rwd/zh/afterTrading/STOCK_DAY_ALL"
        resp = requests.get(url, params={"response": "json"}, headers=HEADERS, timeout=20, verify=False)
        resp.raise_for_status()
        data = resp.json()
        rows = data.get('data', [])
        print(f"[TWSE] 取得 {len(rows)} 筆原始資料")

        for row in rows:
            try:
                code = row[0].strip()
                name = row[1].strip()
                # 只取 4 碼一般股（排除 ETF、特別股等）
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


# ─────────────────────────────────────────────
# 資料來源 2：櫃買中心 (TPEX) - 上櫃股票
# ─────────────────────────────────────────────
def get_tpex_stocks():
    """從 TPEX 取得上櫃股票當日行情，回傳上漲個股"""
    stocks = []
    try:
        url = "https://www.tpex.org.tw/openapi/v1/tpex_mainboard_quotes"
        resp = requests.get(url, headers=HEADERS, timeout=20, verify=False)
        resp.raise_for_status()
        data = resp.json()
        print(f"[TPEX] 取得 {len(data)} 筆原始資料")

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
    """合併上市+上櫃，取漲幅前 N 名"""
    all_stocks = get_twse_stocks() + get_tpex_stocks()
    all_stocks.sort(key=lambda x: -x['change_pct'])
    result = all_stocks[:top_n]
    print(f"[合計] 漲幅前 {len(result)} 名")
    return result


# ─────────────────────────────────────────────
# 取得歷史成交量 (yfinance)
# ─────────────────────────────────────────────
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
        print(f"  [yfinance] {symbol} 失敗：{e}")
        return None, symbol


# ─────────────────────────────────────────────
# 均量條件分析
# ─────────────────────────────────────────────
def analyze_volume_condition(volume_series, min_days=3, max_days=5):
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


# ─────────────────────────────────────────────
# Flask Routes
# ─────────────────────────────────────────────
@app.route('/')
def index():
    return send_from_directory('static', 'index.html')


@app.route('/api/analyze')
def analyze():
    """SSE 串流 API，逐步回報進度"""
    def generate():
        start_time = time.time()

        yield f"data: {json.dumps({'type':'status','msg':'正在從 TWSE / TPEX 取得漲幅排行榜...'}, ensure_ascii=False)}\n\n"

        stocks = get_ranking_stocks(top_n=100)

        if not stocks:
            yield f"data: {json.dumps({'type':'error','msg':'無法取得排行資料，TWSE/TPEX API 可能暫時離線，請稍後重試'}, ensure_ascii=False)}\n\n"
            return

        total = len(stocks)
        yield f"data: {json.dumps({'type':'status','msg':f'取得 {total} 支個股，開始逐一分析均量條件...'}, ensure_ascii=False)}\n\n"

        qualified = []
        for i, stock in enumerate(stocks):
            yield f"data: {json.dumps({'type':'progress','current':i+1,'total':total,'code':stock['code'],'name':stock['name']}, ensure_ascii=False)}\n\n"

            volume_data, actual_symbol = get_volume_data(stock['symbol'])
            result = analyze_volume_condition(volume_data)
            if result:
                qualified.append({**stock, 'symbol': actual_symbol, **result})

            time.sleep(0.2)

        qualified.sort(key=lambda x: (-x['consecutive_days'], -x['ratio']))
        elapsed = round(time.time() - start_time, 1)

        yield f"data: {json.dumps({'type':'result','stocks':qualified,'total':len(qualified),'scanned':total,'elapsed':elapsed,'timestamp':datetime.now().strftime('%Y-%m-%d %H:%M:%S')}, ensure_ascii=False)}\n\n"

    return Response(generate(), mimetype='text/event-stream',
                    headers={'Cache-Control': 'no-cache', 'X-Accel-Buffering': 'no'})


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
    print("🚀 台股均量追蹤系統 v2 啟動")
    print("📡 資料來源：TWSE + TPEX 公開 API（不需認證）")
    print("🌐 請開啟瀏覽器前往：http://localhost:5001")
    print("=" * 50)
    app.run(host='0.0.0.0', port=5001, debug=False, threaded=True)
