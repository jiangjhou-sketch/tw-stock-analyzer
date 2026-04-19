#!/usr/bin/env python3
"""
台股均量追蹤系統 v4 - 技術分析版
資料來源：TWSE + TPEX 公開 API
指標：均量條件 + KD / MACD / 布林通道
"""

import requests, time, re, json, threading, uuid, os, io
from datetime import datetime
from flask import Flask, jsonify, send_from_directory, request, send_file
import yfinance as yf
import pandas as pd
import numpy as np

app = Flask(__name__, static_folder='static')

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
    'Accept': 'application/json, text/plain, */*',
    'Accept-Language': 'zh-TW,zh;q=0.9',
}

tasks = {}   # task_id -> task_state


# ══════════════════════════════════════════════
# 錯誤處理
# ══════════════════════════════════════════════
@app.after_request
def add_json_header(response):
    if response.status_code == 404 and not response.content_type.startswith('application/json'):
        response.data = json.dumps({'error': '找不到路由', 'status': 404})
        response.content_type = 'application/json'
    return response

@app.errorhandler(404)
def not_found(e):
    return jsonify({'error': '找不到此路由', 'status': 404}), 404

@app.errorhandler(500)
def server_error(e):
    return jsonify({'error': f'伺服器內部錯誤：{str(e)}', 'status': 500}), 500


# ══════════════════════════════════════════════
# TWSE + TPEX 行情資料
# ══════════════════════════════════════════════

# ── 常見上市股代碼清單（備援用，共約 300 支流動性較好的個股）──
_TWSE_FALLBACK_CODES = [
    # 半導體/電子
    '2330','2454','2382','2308','2303','2317','3711','2379','3034','2357',
    '2395','2337','2344','2376','3008','2385','6770','3481','2351','2334',
    '6446','3702','2449','2301','2474','2388','3533','2353','2356','3035',
    '2409','2323','2347','3036','2399','2360','2377','2340','2363','2049',
    # 金融
    '2881','2882','2886','2884','2891','2885','2892','2880','2883','2887',
    '2888','2889','5876','5880','2823','2836','5820','2841','2838','2834',
    # 傳產/食品
    '1301','1303','1216','1326','1402','2002','2105','1102','1101','2207',
    '2204','1304','1308','1309','1313','2006','2015','2049','1319','1321',
    # 通訊/網路
    '2412','3045','4904','4938','3231','6669','3714','4977','3006','6409',
    # 光電/面板
    '2475','3481','2353','6415','3293','2426','2069','8046','3324','6128',
    # 生技/醫療
    '4744','4726','6488','4720','4537','1795','4552','4530','6548','4168',
    '6547','4153','4743','4174','4175','6472','4142','4176','6279','4106',
    # 其他電子
    '2408','2498','2492','2486','2484','2478','2467','2466','2465','2464',
    '2461','2458','2450','2448','2436','2433','2432','2430','2428','2423',
    '2420','2419','2417','2414','2413','2410','2406','2405','2404','2402',
    '2401','3057','3088','3094','3105','3130','3149','3189','3205','3209',
    '3211','3227','3234','3242','3259','3276','3287','3289','3305','3311',
    '3315','3317','3321','3325','3330','3338','3374','3380','3413','3416',
    '3419','3427','3432','3443','3450','3455','3462','3463','3464','3466',
    '3468','3494','3501','3504','3508','3513','3515','3521','3526','3527',
    # 建材/運輸
    '2603','2609','2615','2618','2637','2601','5701','2606','2605','2604',
    '2610','2611','2616','5608','2630','6472','2634','2636','2637','9910',
    # 百貨/零售
    '2912','2903','2915','2905','2906','9945','2910','2908','2504','2511',
]

def _parse_twse_row(row):
    """解析 TWSE STOCK_DAY_ALL 的一行資料，回傳 stock dict 或 None"""
    try:
        code = row[0].strip()
        if not re.match(r'^\d{4}$', code):
            return None
        name   = row[1].strip()
        close  = float(row[7].replace(',', '').strip())
        change = float(row[8].replace(',', '').strip())
        prev   = close - change
        if prev <= 0: return None
        chg_pct = round((change / prev) * 100, 2)
        if chg_pct <= 0: return None
        return {'code': code, 'symbol': f"{code}.TW", 'name': name,
                'change_pct': chg_pct, 'price': close,
                'volume': int(row[2].replace(',', '')), 'market': '上市'}
    except Exception:
        return None

def _twse_via_api():
    """嘗試 TWSE 官方 API（多個端點）"""
    endpoints = [
        ("https://www.twse.com.tw/rwd/zh/afterTrading/STOCK_DAY_ALL", {"response":"json"}),
        ("https://www.twse.com.tw/exchangeReport/STOCK_DAY_ALL",       {"response":"json"}),
        ("https://openapi.twse.com.tw/v1/exchangeReport/STOCK_DAY_ALL", {}),
    ]
    headers_list = [
        # 方案1：模擬 XHR
        {**HEADERS, 'Referer':'https://www.twse.com.tw/zh/trading/historical/stock-day-all.html',
         'X-Requested-With':'XMLHttpRequest',
         'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36'},
        # 方案2：模擬一般瀏覽器
        {'User-Agent':'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Safari/605.1.15',
         'Referer':'https://www.twse.com.tw/','Accept':'*/*'},
    ]
    for url, params in endpoints:
        for hdrs in headers_list:
            try:
                session = requests.Session()
                # 先訪問首頁以取得 cookie
                session.get("https://www.twse.com.tw/zh/", headers=hdrs, timeout=8)
                resp = session.get(url, params=params, headers=hdrs, timeout=20)
                if resp.status_code != 200:
                    continue
                ct = resp.headers.get('content-type', '')
                if 'json' not in ct and not resp.text.strip().startswith('{'):
                    continue
                data = resp.json()
                # openapi 回傳 list，舊 API 回傳 {"data":[...]}
                rows = data if isinstance(data, list) else data.get('data', [])
                if not rows:
                    continue
                # 若 list of dict（openapi 格式），轉換欄位
                if isinstance(rows[0], dict):
                    stocks = []
                    for item in rows:
                        try:
                            code = str(item.get('Code','') or item.get('stock_code','')).strip()
                            if not re.match(r'^\d{4}$', code): continue
                            close  = float(str(item.get('ClosingPrice','0')).replace(',','') or 0)
                            change = float(str(item.get('Change','0')).replace(',','') or 0)
                            prev   = close - change
                            if prev <= 0: continue
                            chg = round((change/prev)*100, 2)
                            if chg <= 0: continue
                            vol = int(str(item.get('TradeVolume','0')).replace(',','') or 0)
                            stocks.append({'code':code,'symbol':f"{code}.TW",
                                           'name':str(item.get('Name',code)),
                                           'change_pct':chg,'price':close,
                                           'volume':vol,'market':'上市'})
                        except Exception: continue
                    if stocks:
                        print(f"[TWSE API openapi] {url} 成功，取得 {len(stocks)} 支")
                        return stocks
                else:
                    stocks = [s for row in rows for s in [_parse_twse_row(row)] if s]
                    if stocks:
                        print(f"[TWSE API] {url} 成功，取得 {len(stocks)} 支")
                        return stocks
            except Exception as e:
                continue
    return []

def _twse_via_yfinance(top_n=200):
    """
    TWSE API 全部失敗時的備援：
    用 yfinance 批次下載已知上市股，計算當日漲跌幅後篩出上漲個股
    """
    print("[TWSE fallback] TWSE API 失敗，改用 yfinance 批次下載...")
    codes = list(dict.fromkeys(_TWSE_FALLBACK_CODES))  # 去重保序
    symbols = [f"{c}.TW" for c in codes]

    # 分批下載（每批 30 支，避免超時）
    all_info = {}
    batch_size = 30
    for i in range(0, len(symbols), batch_size):
        batch = symbols[i:i+batch_size]
        try:
            df = yf.download(batch, period='5d', auto_adjust=True,
                             progress=False, timeout=30, group_by='ticker')
            if df.empty:
                continue
            # group_by='ticker' 時結構是 MultiIndex columns
            for sym in batch:
                try:
                    if len(batch) == 1:
                        cl = df['Close'].dropna()
                        op = df['Open'].dropna()
                        vo = df['Volume'].dropna()
                    else:
                        cl = df[sym]['Close'].dropna() if sym in df else pd.Series(dtype=float)
                        op = df[sym]['Open'].dropna()  if sym in df else pd.Series(dtype=float)
                        vo = df[sym]['Volume'].dropna() if sym in df else pd.Series(dtype=float)
                    if len(cl) < 1: continue
                    close_v = float(cl.iloc[-1])
                    open_v  = float(op.iloc[-1]) if len(op)>0 else close_v
                    vol_v   = int(vo.iloc[-1])   if len(vo)>0 else 0
                    if open_v <= 0: continue
                    chg_pct = round((close_v - open_v) / open_v * 100, 2)
                    if chg_pct <= 0: continue
                    code = sym.replace('.TW','')
                    all_info[sym] = {
                        'code': code, 'symbol': sym,
                        'name': code,   # yfinance 批次無法取名稱，先用代碼
                        'change_pct': chg_pct, 'price': close_v,
                        'volume': vol_v, 'market': '上市',
                    }
                except Exception:
                    continue
        except Exception as e:
            print(f"  [yf batch {i//batch_size+1}] 失敗: {e}")
        time.sleep(0.3)

    stocks = list(all_info.values())
    print(f"[TWSE fallback] yfinance 取得上漲 {len(stocks)} 支")
    return stocks

def get_twse_stocks():
    """取得上市股漲幅排行（多層備援）"""
    # 方案1：直連 TWSE API
    stocks = _twse_via_api()
    if stocks:
        return stocks
    # 方案2：yfinance 備援
    stocks = _twse_via_yfinance()
    if stocks:
        return stocks
    print("[TWSE] 所有方案均失敗，回傳空清單")
    return []


def get_tpex_stocks():
    stocks = []
    try:
        url = "https://www.tpex.org.tw/openapi/v1/tpex_mainboard_quotes"
        resp = requests.get(url, headers=HEADERS, timeout=20)
        resp.raise_for_status()
        for item in resp.json():
            try:
                code = str(item.get('SecuritiesCompanyCode', '')).strip()
                if not re.match(r'^\d{4,5}$', code): continue
                name   = str(item.get('CompanyName', '')).strip()
                close  = float(str(item.get('Close',  '0')).replace(',', '') or 0)
                change = float(str(item.get('Change', '0')).replace(',', '') or 0)
                if close <= 0: continue
                prev = close - change
                if prev <= 0: continue
                chg_pct = round((change / prev) * 100, 2)
                if chg_pct <= 0: continue
                vol_s = str(item.get('TradeVolume', '0')).replace(',', '')
                stocks.append({'code': code, 'symbol': f"{code}.TWO", 'name': name,
                                'change_pct': chg_pct, 'price': close,
                                'volume': int(float(vol_s)) if vol_s else 0, 'market': '上櫃'})
            except Exception:
                continue
        print(f"[TPEX] 上漲：{len(stocks)} 支")
    except Exception as e:
        print(f"[TPEX] 失敗：{e}")
    return stocks


def get_ranking_stocks(top_n=100, mode='both'):
    if mode == 'twse':
        s = get_twse_stocks(); s.sort(key=lambda x: -x['change_pct']); return s[:top_n]
    if mode == 'tpex':
        s = get_tpex_stocks(); s.sort(key=lambda x: -x['change_pct']); return s[:top_n]
    if mode == 'combined':
        s = get_twse_stocks() + get_tpex_stocks(); s.sort(key=lambda x: -x['change_pct']); return s[:top_n]
    # both
    tw = get_twse_stocks(); tw.sort(key=lambda x: -x['change_pct'])
    tp = get_tpex_stocks(); tp.sort(key=lambda x: -x['change_pct'])
    return tw[:top_n] + tp[:top_n]


# ══════════════════════════════════════════════
# 取得 OHLCV（Open / High / Low / Close / Volume）
# ══════════════════════════════════════════════
def get_ohlcv(symbol, days=90):
    """取得足夠天數的 OHLCV 資料（需要至少 60 根計算 MACD）"""
    try:
        ticker = yf.Ticker(symbol)
        hist   = ticker.history(period=f"{days}d")
        if hist.empty:
            alt  = symbol.replace('.TW', '.TWO') if symbol.endswith('.TW') else symbol.replace('.TWO', '.TW')
            hist = yf.Ticker(alt).history(period=f"{days}d")
            if not hist.empty:
                return hist, alt
        return (hist, symbol) if not hist.empty else (None, symbol)
    except Exception:
        return None, symbol


# ══════════════════════════════════════════════
# 技術指標計算（純 pandas / numpy，不需額外套件）
# ══════════════════════════════════════════════
def calc_kd(high, low, close, n=9, m1=3, m2=3):
    """
    KD 隨機指標（台式 9 日 KD）
    回傳 K、D 最新值及交叉訊號
    """
    low_n  = low.rolling(n).min()
    high_n = high.rolling(n).max()
    denom  = high_n - low_n
    rsv    = np.where(denom == 0, 50.0,
                      (close - low_n) / denom * 100)
    rsv_s  = pd.Series(rsv, index=close.index)

    K = pd.Series(50.0, index=close.index, dtype=float)
    D = pd.Series(50.0, index=close.index, dtype=float)
    for i in range(1, len(rsv_s)):
        K.iloc[i] = (m1 - 1) / m1 * K.iloc[i-1] + 1 / m1 * rsv_s.iloc[i]
        D.iloc[i] = (m2 - 1) / m2 * D.iloc[i-1] + 1 / m2 * K.iloc[i]

    k_cur, k_prev = round(K.iloc[-1], 2), round(K.iloc[-2], 2)
    d_cur, d_prev = round(D.iloc[-1], 2), round(D.iloc[-2], 2)

    # 黃金交叉：K 由下往上穿越 D；死亡交叉：K 由上往下穿越 D
    golden = (k_prev < d_prev) and (k_cur > d_cur)
    death  = (k_prev > d_prev) and (k_cur < d_cur)
    oversold   = k_cur < 20          # 超賣區
    overbought = k_cur > 80          # 超買區
    k_above_d  = k_cur > d_cur

    return {
        'k': k_cur, 'd': d_cur,
        'kd_golden': golden,
        'kd_death':  death,
        'kd_oversold':   oversold,
        'kd_overbought': overbought,
        'kd_k_above_d':  k_above_d,
        'kd_signal': ('黃金交叉' if golden else
                      '死亡交叉' if death  else
                      '超賣區'   if oversold else
                      '多頭排列' if k_above_d else '空頭排列'),
    }


def calc_macd(close, fast=12, slow=26, signal=9):
    """
    MACD（指數移動平均差）
    回傳 MACD 線、訊號線、柱狀圖最新值及交叉訊號
    """
    ema_fast   = close.ewm(span=fast,   adjust=False).mean()
    ema_slow   = close.ewm(span=slow,   adjust=False).mean()
    macd_line  = ema_fast - ema_slow
    sig_line   = macd_line.ewm(span=signal, adjust=False).mean()
    histogram  = macd_line - sig_line

    m_cur  = round(macd_line.iloc[-1], 4)
    m_prev = round(macd_line.iloc[-2], 4)
    s_cur  = round(sig_line.iloc[-1],  4)
    s_prev = round(sig_line.iloc[-2],  4)
    h_cur  = round(histogram.iloc[-1], 4)
    h_prev = round(histogram.iloc[-2], 4)

    golden   = (m_prev < s_prev) and (m_cur > s_cur)
    death    = (m_prev > s_prev) and (m_cur < s_cur)
    above_zero = m_cur > 0
    hist_expand = h_cur > 0 and h_cur > h_prev  # 柱狀圖正向擴張

    return {
        'macd': m_cur, 'macd_signal': s_cur, 'macd_hist': h_cur,
        'macd_golden':      golden,
        'macd_death':       death,
        'macd_above_zero':  above_zero,
        'macd_hist_expand': hist_expand,
        'macd_signal_str': ('黃金交叉' if golden  else
                             '死亡交叉' if death   else
                             '零軸上方' if above_zero else '零軸下方'),
    }


def calc_bband(close, n=20, k=2):
    """
    布林通道（Bollinger Bands）
    回傳上軌、中軌、下軌及目前位置訊號
    """
    mid    = close.rolling(n).mean()
    std    = close.rolling(n).std()
    upper  = mid + k * std
    lower  = mid - k * std
    bw     = ((upper - lower) / mid * 100)  # 帶寬 %

    c_cur   = close.iloc[-1]
    u_cur   = round(upper.iloc[-1], 2)
    m_cur   = round(mid.iloc[-1],   2)
    l_cur   = round(lower.iloc[-1], 2)
    bw_cur  = round(bw.iloc[-1],    2)

    # 位置：0=下軌 ~ 1=上軌
    pos = round((c_cur - l_cur) / (u_cur - l_cur), 3) if (u_cur - l_cur) > 0 else 0.5

    near_upper = c_cur >= u_cur * 0.99    # 突破或接近上軌
    near_lower = c_cur <= l_cur * 1.01    # 觸及或接近下軌
    above_mid  = c_cur > m_cur

    return {
        'bb_upper': u_cur, 'bb_mid': m_cur, 'bb_lower': l_cur,
        'bb_width': bw_cur, 'bb_pos': pos,
        'bb_near_upper': near_upper,
        'bb_near_lower': near_lower,
        'bb_above_mid':  above_mid,
        'bb_signal': ('突破上軌' if near_upper else
                      '接近下軌' if near_lower else
                      '中軌以上' if above_mid  else '中軌以下'),
    }


def calc_all_ta(hist):
    """整合計算所有技術指標，回傳 dict；失敗傳回空 dict"""
    try:
        if hist is None or len(hist) < 30:
            return {}
        high  = hist['High']
        low   = hist['Low']
        close = hist['Close']
        ta = {}
        ta.update(calc_kd(high, low, close))
        ta.update(calc_macd(close))
        ta.update(calc_bband(close))
        return ta
    except Exception as e:
        print(f"  [TA] 計算失敗：{e}")
        return {}


# ══════════════════════════════════════════════
# 均量條件
# ══════════════════════════════════════════════
def analyze_volume_condition(volume_series, min_days=2, max_days=5):
    if volume_series is None or len(volume_series) < 25:
        return None
    ma5  = volume_series.rolling(window=5).mean()
    ma20 = volume_series.rolling(window=20).mean()
    df   = pd.DataFrame({'ma5': ma5, 'ma20': ma20}).dropna()
    if len(df) == 0:
        return None
    cond = df['ma5'] > df['ma20']
    consecutive = 0
    for v in reversed(cond.values):
        if v: consecutive += 1
        else: break
    if not (min_days <= consecutive <= max_days):
        return None
    return {
        'consecutive_days': consecutive,
        'ma5_volume':  int(round(df['ma5'].iloc[-1])),
        'ma20_volume': int(round(df['ma20'].iloc[-1])),
        'ratio': round(df['ma5'].iloc[-1] / df['ma20'].iloc[-1], 3) if df['ma20'].iloc[-1] > 0 else 0,
    }


# ══════════════════════════════════════════════
# 背景任務
# ══════════════════════════════════════════════
def run_analysis_task(task_id):
    task = tasks[task_id]
    task['status'] = 'fetching'
    start_time = time.time()

    try:
        mode = task.get('mode', 'both')
        mode_labels = {'twse':'上市前100', 'tpex':'上櫃前100',
                       'combined':'合併前100', 'both':'上市前100+上櫃前100'}
        task['msg'] = f"正在取得【{mode_labels.get(mode,mode)}】排行榜..."

        # 暫時 hook print 以截取備援訊息
        import sys as _sys
        class _MsgCapture:
            def __init__(self, orig, task):
                self.orig = orig; self.task = task
            def write(self, s):
                self.orig.write(s)
                s = s.strip()
                if s and not s.startswith('[') or 'fallback' in s or '備援' in s or '失敗' in s:
                    if len(s) > 5:
                        self.task['msg'] = s[:80]
            def flush(self): self.orig.flush()
        _orig_stdout = _sys.stdout
        _sys.stdout = _MsgCapture(_orig_stdout, task)
        try:
            stocks = get_ranking_stocks(top_n=100, mode=mode)
        finally:
            _sys.stdout = _orig_stdout
        if not stocks:
            task['status'] = 'error'
            task['msg'] = '無法取得排行資料，TWSE/TPEX API 可能暫時離線'
            return

        total = len(stocks)
        task['total'] = total
        task['status'] = 'analyzing'
        task['msg'] = f'取得 {total} 支個股，計算均量＋技術指標...'

        qualified = []
        for i, stock in enumerate(stocks):
            task['current']      = i + 1
            task['current_code'] = stock['code']
            task['current_name'] = stock['name']

            hist, actual_symbol = get_ohlcv(stock['symbol'], days=90)
            vol_result = analyze_volume_condition(
                hist['Volume'] if hist is not None else None)

            if vol_result:
                ta = calc_all_ta(hist)
                entry = {**stock, 'symbol': actual_symbol, **vol_result, **ta}
                qualified.append(entry)
                task['total_found'] = len(qualified)

            time.sleep(0.15)

        qualified.sort(key=lambda x: (-x['consecutive_days'], -x['ratio']))
        elapsed = round(time.time() - start_time, 1)

        task.update({
            'status': 'done', 'stocks': qualified,
            'total_found': len(qualified), 'scanned': total,
            'elapsed': elapsed,
            'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'msg': f'完成！找到 {len(qualified)} 支符合均量條件個股',
        })

    except Exception as e:
        task['status'] = 'error'
        task['msg'] = f'分析過程發生錯誤：{str(e)}'


# ══════════════════════════════════════════════
# Flask Routes
# ══════════════════════════════════════════════
@app.route('/')
def index():
    static_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'static')
    return send_from_directory(static_dir, 'index.html')


@app.route('/api/analyze/start', methods=['POST', 'GET'])
def analyze_start():
    body = request.get_json(silent=True) or {}
    mode = body.get('mode', 'both')
    if mode not in ('twse', 'tpex', 'combined', 'both'):
        mode = 'both'
    task_id = str(uuid.uuid4())[:8]
    tasks[task_id] = {
        'status': 'pending', 'msg': '準備開始...', 'mode': mode,
        'current': 0, 'total': 0, 'current_code': '', 'current_name': '',
        'stocks': [], 'total_found': 0, 'scanned': 0, 'elapsed': 0, 'timestamp': '',
    }
    threading.Thread(target=run_analysis_task, args=(task_id,), daemon=True).start()
    return jsonify({'task_id': task_id})


@app.route('/api/analyze/status/<task_id>')
def analyze_status(task_id):
    task = tasks.get(task_id)
    if not task:
        return jsonify({'error': '找不到任務'}), 404
    return jsonify(task)


@app.route('/api/stock/<code>')
def stock_detail(code):
    """取得個股完整歷史（OHLCV + 技術指標序列，供圖表繪製）"""
    hist, actual_symbol = get_ohlcv(f"{code}.TW", days=90)
    if hist is None:
        hist, actual_symbol = get_ohlcv(f"{code}.TWO", days=90)
    if hist is None:
        return jsonify({'error': f'無法取得 {code} 的資料'}), 404

    close  = hist['Close']
    high   = hist['High']
    low    = hist['Low']
    volume = hist['Volume']
    dates  = [d.strftime('%Y-%m-%d') for d in hist.index]

    def s(series):
        return [round(v, 4) if not pd.isna(v) else None for v in series]

    # Volume MAs
    ma5v  = volume.rolling(5).mean()
    ma20v = volume.rolling(20).mean()

    # Price MAs
    ma5p  = close.rolling(5).mean()
    ma20p = close.rolling(20).mean()

    # MACD series
    ema12     = close.ewm(span=12, adjust=False).mean()
    ema26     = close.ewm(span=26, adjust=False).mean()
    macd_line = ema12 - ema26
    sig_line  = macd_line.ewm(span=9, adjust=False).mean()
    histogram = macd_line - sig_line

    # KD series
    low9  = low.rolling(9).min()
    high9 = high.rolling(9).max()
    denom = high9 - low9
    rsv   = np.where(denom == 0, 50.0, (close - low9) / denom * 100)
    rsv_s = pd.Series(rsv, index=close.index)
    K     = pd.Series(50.0, index=close.index, dtype=float)
    D     = pd.Series(50.0, index=close.index, dtype=float)
    for i in range(1, len(rsv_s)):
        K.iloc[i] = 2/3 * K.iloc[i-1] + 1/3 * rsv_s.iloc[i]
        D.iloc[i] = 2/3 * D.iloc[i-1] + 1/3 * K.iloc[i]

    # Bollinger Bands
    bb_mid   = close.rolling(20).mean()
    bb_std   = close.rolling(20).std()
    bb_upper = bb_mid + 2 * bb_std
    bb_lower = bb_mid - 2 * bb_std

    def iv(series):  # int list for volume
        return [int(v) if not pd.isna(v) else None for v in series]

    return jsonify({
        'code': code, 'symbol': actual_symbol, 'dates': dates,
        'open':  s(hist['Open']),
        'high':  s(high),
        'low':   s(low),
        'close': s(close),
        'volume': iv(volume),
        'ma5v': iv(ma5v), 'ma20v': iv(ma20v),
        'ma5p': s(ma5p),  'ma20p': s(ma20p),
        'macd': s(macd_line), 'macd_signal': s(sig_line), 'macd_hist': s(histogram),
        'kd_k': s(K), 'kd_d': s(D),
        'bb_upper': s(bb_upper), 'bb_mid': s(bb_mid), 'bb_lower': s(bb_lower),
    })


# ══════════════════════════════════════════════
# 匯出 PDF
# ══════════════════════════════════════════════
@app.route('/api/export/pdf', methods=['POST'])
def export_pdf():
    from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle,
                                     Paragraph, Spacer, HRFlowable)
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.lib import colors
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont

    body       = request.get_json(force=True)
    stocks     = body.get('stocks', [])
    ts         = body.get('timestamp', datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
    scanned    = body.get('scanned', 0)
    mode_label = body.get('mode_label', '')

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=landscape(A4),
                             leftMargin=12*mm, rightMargin=12*mm,
                             topMargin=14*mm, bottomMargin=14*mm)

    CN_FONT = 'Helvetica'
    try:
        import glob
        cands = (glob.glob('/usr/share/fonts/truetype/noto/*CJK*Regular*.ttf') +
                 glob.glob('/usr/share/fonts/truetype/wqy/*.ttf'))
        if cands:
            pdfmetrics.registerFont(TTFont('CJK', cands[0]))
            CN_FONT = 'CJK'
    except Exception:
        pass

    title_st = ParagraphStyle('T', fontName=CN_FONT, fontSize=15,
                               textColor=colors.HexColor('#00d4aa'), spaceAfter=3)
    sub_st   = ParagraphStyle('S', fontName=CN_FONT, fontSize=8,
                               textColor=colors.HexColor('#94a3b8'), spaceAfter=8)
    foot_st  = ParagraphStyle('F', fontName=CN_FONT, fontSize=7,
                               textColor=colors.HexColor('#64748b'))

    story = [
        Paragraph('台股均量追蹤系統 掃描報告', title_st),
        Paragraph(f'掃描時間：{ts}　範圍：{mode_label}　已掃描：{scanned} 支　符合均量條件：{len(stocks)} 支', sub_st),
        Paragraph('篩選條件：5日均量 > 20日均量 連續2～5日　技術指標欄位為最新數值', sub_st),
        HRFlowable(width='100%', thickness=0.8, color=colors.HexColor('#1e2d47'), spaceAfter=8),
    ]

    if not stocks:
        story.append(Paragraph('本次掃描無符合條件的個股。', sub_st))
    else:
        def fv(n):
            if not n: return '-'
            if n >= 1e8: return f"{n/1e8:.1f}億"
            if n >= 1e4: return f"{n/1e3:.0f}K"
            return str(n)
        def boolmark(v): return '✓' if v else ''

        hdrs = ['代碼','名稱','市場','漲幅','連續','5MA量','比率',
                'K值','D值','KD訊號','MACD','MACD訊號','柱狀','MACD訊號',
                '布林位','布林訊號']
        rows = [hdrs]
        for s in stocks:
            chg = f"+{s['change_pct']:.2f}" if s.get('change_pct',0)>=0 else f"{s.get('change_pct',0):.2f}"
            rows.append([
                s.get('code',''), s.get('name',''), s.get('market',''),
                chg+'%', str(s.get('consecutive_days','')),
                fv(s.get('ma5_volume')), f"{s.get('ratio',0):.2f}x",
                f"{s.get('k',0):.1f}", f"{s.get('d',0):.1f}",
                s.get('kd_signal','-'),
                f"{s.get('macd',0):.3f}", f"{s.get('macd_signal',0):.3f}",
                f"{s.get('macd_hist',0):.3f}", s.get('macd_signal_str','-'),
                f"{s.get('bb_pos',0):.0%}", s.get('bb_signal','-'),
            ])

        cw = [14,26,11,14,12,16,14, 12,12,18, 16,16,16,18, 14,18]
        cw = [x*mm for x in cw]

        t = Table(rows, colWidths=cw, repeatRows=1)
        cmd = [
            ('FONTNAME',  (0,0),(-1,-1), CN_FONT),
            ('FONTSIZE',  (0,0),(-1,-1), 7),
            ('ALIGN',     (0,0),(-1,-1), 'CENTER'),
            ('VALIGN',    (0,0),(-1,-1), 'MIDDLE'),
            ('BACKGROUND',(0,0),(-1,0),  colors.HexColor('#0f3460')),
            ('TEXTCOLOR', (0,0),(-1,0),  colors.HexColor('#00d4aa')),
            ('FONTSIZE',  (0,0),(-1,0),  7.5),
            ('BOTTOMPADDING',(0,0),(-1,0), 5),
            ('TOPPADDING',   (0,0),(-1,0), 5),
            ('ROWBACKGROUNDS',(0,1),(-1,-1),
             [colors.HexColor('#111827'), colors.HexColor('#0f172a')]),
            ('TEXTCOLOR', (0,1),(-1,-1), colors.HexColor('#cbd5e1')),
            ('GRID',      (0,0),(-1,-1), 0.25, colors.HexColor('#1e2d47')),
            ('BOTTOMPADDING',(0,1),(-1,-1), 4),
            ('TOPPADDING',   (0,1),(-1,-1), 4),
            ('TEXTCOLOR', (3,1),(3,-1), colors.HexColor('#00d98b')),
            ('TEXTCOLOR', (6,1),(6,-1), colors.HexColor('#0084ff')),
        ]
        t.setStyle(TableStyle(cmd))
        story.append(t)

    story += [Spacer(1,6*mm),
              HRFlowable(width='100%',thickness=0.4,color=colors.HexColor('#1e2d47'),spaceAfter=3),
              Paragraph('資料來源：TWSE / TPEX　　本報告僅供參考，不構成投資建議', foot_st)]

    doc.build(story)
    buf.seek(0)
    fname = f"stock_report_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf"
    return send_file(buf, mimetype='application/pdf', as_attachment=True, download_name=fname)


# ══════════════════════════════════════════════
# 匯出 DOCX
# ══════════════════════════════════════════════
@app.route('/api/export/docx', methods=['POST'])
def export_docx():
    from docx import Document
    from docx.shared import Pt, RGBColor, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    body       = request.get_json(force=True)
    stocks     = body.get('stocks', [])
    ts         = body.get('timestamp', datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
    scanned    = body.get('scanned', 0)
    mode_label = body.get('mode_label', '')

    doc = Document()
    sec = doc.sections[0]
    sec.page_width = Cm(42); sec.page_height = Cm(29.7)
    sec.left_margin = sec.right_margin = Cm(1.2)
    sec.top_margin  = sec.bottom_margin = Cm(1.5)

    def set_bg(cell, hex_color):
        tc = cell._tc; pr = tc.get_or_add_tcPr()
        shd = OxmlElement('w:shd')
        shd.set(qn('w:fill'), hex_color); shd.set(qn('w:val'), 'clear')
        pr.append(shd)

    def set_border(cell, color='1E2D47'):
        tc = cell._tc; pr = tc.get_or_add_tcPr()
        borders = OxmlElement('w:tcBorders')
        for s in ['top','left','bottom','right']:
            el = OxmlElement(f'w:{s}')
            el.set(qn('w:val'),'single'); el.set(qn('w:sz'),'4'); el.set(qn('w:color'), color)
            borders.append(el)
        pr.append(borders)

    h = doc.add_heading('台股均量追蹤系統 — 掃描報告', 0)
    h.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r = h.runs[0]; r.font.color.rgb = RGBColor(0,0xD4,0xAA); r.font.size = Pt(16)

    p = doc.add_paragraph()
    for txt in [f'掃描時間：{ts}', f'　範圍：{mode_label}', f'　已掃描：{scanned} 支', f'　符合均量條件：{len(stocks)} 支']:
        run = p.add_run(txt); run.font.size = Pt(9); run.font.color.rgb = RGBColor(0x94,0xA3,0xB8)
    p2 = doc.add_paragraph()
    r2 = p2.add_run('篩選條件：5日均量 > 20日均量 連續2～5日　技術指標為最新值')
    r2.font.size = Pt(8.5); r2.font.color.rgb = RGBColor(0x64,0x74,0x8B)
    doc.add_paragraph()

    if not stocks:
        doc.add_paragraph('本次掃描無符合條件的個股。')
    else:
        def fv(n):
            if not n: return '-'
            if n >= 1e8: return f"{n/1e8:.1f}億"
            if n >= 1e4: return f"{n/1e3:.0f}K"
            return str(n)

        hdrs  = ['代碼','名稱','市場','漲幅%','連續日','5MA量','比率',
                 'K值','D值','KD訊號','MACD','MACD訊號','柱狀圖','MACD狀態',
                 '布林位置','布林訊號']
        widths = [1.6,3.0,1.2,1.6,1.4,2.2,1.6, 1.4,1.4,2.2, 2.0,2.0,2.0,2.2, 2.0,2.2]

        tbl = doc.add_table(rows=1, cols=len(hdrs))
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl.style = 'Table Grid'
        for i, w in enumerate(widths):
            for cell in tbl.columns[i].cells:
                cell.width = Cm(w)

        hrow = tbl.rows[0]
        for cell, txt in zip(hrow.cells, hdrs):
            set_bg(cell, '0F3460'); set_border(cell)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(txt); run.bold = True
            run.font.size = Pt(8); run.font.color.rgb = RGBColor(0,0xD4,0xAA)

        for idx, s in enumerate(stocks):
            row = tbl.add_row(); row.height = Cm(0.7)
            bg = '111827' if idx % 2 == 0 else '0F172A'
            chg = f"+{s['change_pct']:.2f}" if s.get('change_pct',0)>=0 else f"{s.get('change_pct',0):.2f}"
            vals = [
                s.get('code',''), s.get('name',''), s.get('market',''),
                chg+'%', str(s.get('consecutive_days','')),
                fv(s.get('ma5_volume')), f"{s.get('ratio',0):.2f}x",
                f"{s.get('k',0):.1f}", f"{s.get('d',0):.1f}",
                s.get('kd_signal','-'),
                f"{s.get('macd',0):.3f}", f"{s.get('macd_signal',0):.3f}",
                f"{s.get('macd_hist',0):.3f}", s.get('macd_signal_str','-'),
                f"{s.get('bb_pos',0):.0%}", s.get('bb_signal','-'),
            ]
            vcols = [
                RGBColor(0xFF,0xFF,0xFF), RGBColor(0xCB,0xD5,0xE1), RGBColor(0xFF,0xB5,0x47),
                (RGBColor(0,0xD9,0x8B) if s.get('change_pct',0)>=0 else RGBColor(0xFF,0x4D,0x6D)),
                RGBColor(0,0xD4,0xAA), RGBColor(0,0xD4,0xAA), RGBColor(0,0x84,0xFF),
                RGBColor(0xFF,0xD7,0x00), RGBColor(0xFF,0xD7,0x00), RGBColor(0xCB,0xD5,0xE1),
                RGBColor(0,0x84,0xFF),   RGBColor(0x94,0xA3,0xB8), RGBColor(0x94,0xA3,0xB8), RGBColor(0xCB,0xD5,0xE1),
                RGBColor(0x7C,0x3A,0xED), RGBColor(0xCB,0xD5,0xE1),
            ]
            for cell, val, vc in zip(row.cells, vals, vcols):
                set_bg(cell, bg); set_border(cell)
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                p = cell.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = p.add_run(val); run.font.size = Pt(7.5); run.font.color.rgb = vc

    doc.add_paragraph()
    fp = doc.add_paragraph('資料來源：TWSE / TPEX　　本報告僅供參考，不構成投資建議')
    fp.runs[0].font.size = Pt(7.5); fp.runs[0].font.color.rgb = RGBColor(0x64,0x74,0x8B)

    buf = io.BytesIO(); doc.save(buf); buf.seek(0)
    fname = f"stock_report_{datetime.now().strftime('%Y%m%d_%H%M')}.docx"
    return send_file(buf,
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        as_attachment=True, download_name=fname)


if __name__ == '__main__':
    os.makedirs('static', exist_ok=True)
    print("=" * 55)
    print("🚀 台股均量追蹤系統 v4 啟動（技術分析版）")
    print("📡 資料來源：TWSE + TPEX 公開 API")
    print("📊 技術指標：KD / MACD / 布林通道")
    print("🌐 請開啟瀏覽器前往：http://localhost:5001")
    print("=" * 55)
    app.run(host='0.0.0.0', port=5001, debug=False, threaded=True)
