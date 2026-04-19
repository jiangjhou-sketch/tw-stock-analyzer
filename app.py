#!/usr/bin/env python3
"""
台股均量追蹤系統 v4
資料來源：TWSE (三層備援) + TPEX OpenAPI
技術指標：KD / MACD / 布林通道
"""

import requests, time, re, json, threading, uuid, os, io
from datetime import datetime
from flask import Flask, jsonify, send_from_directory, request, send_file
import yfinance as yf
import pandas as pd
import numpy as np

app = Flask(__name__, static_folder='static')

BASE_HEADERS = {
    'User-Agent': (
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
        'AppleWebKit/537.36 (KHTML, like Gecko) '
        'Chrome/124.0.0.0 Safari/537.36'
    ),
    'Accept': 'application/json, text/javascript, */*; q=0.01',
    'Accept-Language': 'zh-TW,zh;q=0.9,en-US;q=0.8',
}

tasks = {}  # task_id -> state dict


# ══════════════════════════════════════════════════════
# Flask 錯誤處理
# ══════════════════════════════════════════════════════
@app.errorhandler(404)
def not_found(e):
    return jsonify({'error': '找不到此路由', 'status': 404}), 404

@app.errorhandler(500)
def server_error(e):
    return jsonify({'error': f'伺服器錯誤：{e}', 'status': 500}), 500

@app.after_request
def ensure_json_on_error(resp):
    if resp.status_code in (404, 500) and 'application/json' not in resp.content_type:
        resp.data = json.dumps({'error': f'HTTP {resp.status_code}'})
        resp.content_type = 'application/json'
    return resp


# ══════════════════════════════════════════════════════
# TWSE 上市股資料（三層備援）
# ══════════════════════════════════════════════════════

# 約 300 支流動性較好的上市股代碼（yfinance 備援用）
_FALLBACK_TWSE = [
    '2330','2454','2382','2308','2303','2317','3711','2379','3034','2357',
    '2395','2337','2344','2376','3008','2385','6770','3481','2351','2334',
    '6446','3702','2449','2301','2474','2388','3533','2353','2356','3035',
    '2409','2323','2347','3036','2399','2360','2377','2340','2363','2049',
    '2881','2882','2886','2884','2891','2885','2892','2880','2883','2887',
    '2888','2889','5876','5880','2823','2836','5820','2841','2838','2834',
    '1301','1303','1216','1326','1402','2002','2105','1102','1101','2207',
    '2204','1304','1308','1309','1313','2006','2015','1319','1321','1402',
    '2412','3045','4904','4938','3231','6669','3714','4977','3006','6409',
    '2475','6415','3293','2426','2069','8046','3324','6128','2408','2498',
    '2492','2486','2484','2478','2467','2466','2465','2464','2461','2458',
    '2450','2448','2436','2433','2432','2430','2428','2423','2420','2419',
    '2417','2414','2413','2410','2406','2405','2404','2402','2401','2498',
    '3057','3088','3094','3105','3130','3149','3189','3205','3209','3211',
    '3227','3234','3242','3259','3276','3287','3289','3305','3311','3315',
    '3317','3321','3325','3330','3338','3374','3380','3413','3416','3419',
    '3427','3432','3443','3450','3455','3462','3463','3464','3466','3468',
    '3494','3501','3504','3508','3513','3515','3521','3526','3527','4744',
    '4726','6488','4720','4537','1795','4552','4530','6548','4168','6547',
    '4153','4743','4174','4175','6472','4142','4176','6279','4106','2603',
    '2609','2615','2618','2637','2601','5701','2606','2605','2604','2610',
    '2611','2616','5608','2630','2634','2636','2912','2903','2915','2905',
    '2906','9945','2910','2908','2504','2511','6271','6669','8詣45','5234',
    '2441','2443','2444','2449','6005','6147','6148','6150','6152','6153',
    '6154','6155','6158','6160','6161','6162','6163','6164','6165','6166',
    '6167','6168','6169','6170','6171','6172','6173','6174','6175','6176',
    '6177','6178','6179','6180','6181','6182','6183','6184','6185','6186',
    '6187','6188','6189','6190','6191','6192','6193','6194','6195','6196',
    '6197','6198','6199','6200','2912','3037','2441','2449','6271','5234',
]

def _parse_twse_row(row):
    """解析 TWSE 舊式 data row → stock dict 或 None"""
    try:
        code = str(row[0]).strip()
        if not re.match(r'^\d{4}$', code):
            return None
        name   = str(row[1]).strip()
        close  = float(str(row[7]).replace(',', '').strip())
        change = float(str(row[8]).replace(',', '').strip())
        prev   = close - change
        if prev <= 0:
            return None
        chg = round((change / prev) * 100, 2)
        if chg <= 0:
            return None
        vol = int(str(row[2]).replace(',', ''))
        return {
            'code': code, 'symbol': f'{code}.TW', 'name': name,
            'change_pct': chg, 'price': close, 'volume': vol, 'market': '上市',
        }
    except Exception:
        return None


def _parse_twse_openapi_item(item):
    """解析 TWSE OpenAPI dict 格式 → stock dict 或 None"""
    try:
        code = str(item.get('Code', item.get('stock_code', ''))).strip()
        if not re.match(r'^\d{4}$', code):
            return None
        name   = str(item.get('Name', code))
        close  = float(str(item.get('ClosingPrice', '0')).replace(',', '') or 0)
        change = float(str(item.get('Change', '0')).replace(',', '') or 0)
        prev   = close - change
        if prev <= 0:
            return None
        chg = round((change / prev) * 100, 2)
        if chg <= 0:
            return None
        vol = int(float(str(item.get('TradeVolume', '0')).replace(',', '') or 0))
        return {
            'code': code, 'symbol': f'{code}.TW', 'name': name,
            'change_pct': chg, 'price': close, 'volume': vol, 'market': '上市',
        }
    except Exception:
        return None


def _twse_direct_api():
    """嘗試直連 TWSE API（多端點 × 多 header 組合）"""
    endpoints = [
        ('https://www.twse.com.tw/rwd/zh/afterTrading/STOCK_DAY_ALL', {'response': 'json'}),
        ('https://www.twse.com.tw/exchangeReport/STOCK_DAY_ALL',       {'response': 'json'}),
        ('https://openapi.twse.com.tw/v1/exchangeReport/STOCK_DAY_ALL', {}),
    ]
    header_variants = [
        {**BASE_HEADERS,
         'Referer': 'https://www.twse.com.tw/zh/trading/historical/stock-day-all.html',
         'X-Requested-With': 'XMLHttpRequest'},
        {'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) Safari/605.1.15',
         'Referer': 'https://www.twse.com.tw/',
         'Accept': '*/*'},
    ]
    for url, params in endpoints:
        for hdrs in header_variants:
            try:
                sess = requests.Session()
                sess.get('https://www.twse.com.tw/zh/', headers=hdrs, timeout=8)
                resp = sess.get(url, params=params, headers=hdrs, timeout=20)
                if resp.status_code != 200:
                    continue
                data = resp.json()
                rows = data if isinstance(data, list) else data.get('data', [])
                if not rows:
                    continue
                if isinstance(rows[0], dict):
                    stocks = [s for item in rows for s in [_parse_twse_openapi_item(item)] if s]
                else:
                    stocks = [s for row in rows for s in [_parse_twse_row(row)] if s]
                if stocks:
                    print(f'[TWSE direct] {url} OK → {len(stocks)} 支')
                    return stocks
            except Exception:
                continue
    return []


def _twse_yfinance_fallback():
    """TWSE API 失敗時：用 yfinance 批次取已知上市股漲跌幅"""
    print('[TWSE fallback] 改用 yfinance 批次下載...')
    codes   = list(dict.fromkeys(c for c in _FALLBACK_TWSE if re.match(r'^\d{4}$', c)))
    symbols = [f'{c}.TW' for c in codes]
    result  = {}

    for i in range(0, len(symbols), 30):
        batch = symbols[i:i+30]
        try:
            df = yf.download(batch, period='5d', auto_adjust=True,
                             progress=False, timeout=30, group_by='ticker')
            if df.empty:
                continue
            for sym in batch:
                try:
                    if len(batch) == 1:
                        cl = df['Close'].dropna()
                        op = df['Open'].dropna()
                        vo = df['Volume'].dropna()
                    else:
                        if sym not in df.columns.get_level_values(0):
                            continue
                        cl = df[sym]['Close'].dropna()
                        op = df[sym]['Open'].dropna()
                        vo = df[sym]['Volume'].dropna()
                    if cl.empty:
                        continue
                    c_val = float(cl.iloc[-1])
                    o_val = float(op.iloc[-1]) if not op.empty else c_val
                    v_val = int(vo.iloc[-1])   if not vo.empty else 0
                    if o_val <= 0:
                        continue
                    chg = round((c_val - o_val) / o_val * 100, 2)
                    if chg <= 0:
                        continue
                    code = sym.replace('.TW', '')
                    result[sym] = {
                        'code': code, 'symbol': sym, 'name': code,
                        'change_pct': chg, 'price': c_val,
                        'volume': v_val, 'market': '上市',
                    }
                except Exception:
                    continue
        except Exception as e:
            print(f'  [yf batch] 失敗: {e}')
        time.sleep(0.3)

    stocks = list(result.values())
    print(f'[TWSE fallback] 取得 {len(stocks)} 支上漲個股')
    return stocks


def get_twse_stocks():
    stocks = _twse_direct_api()
    if not stocks:
        stocks = _twse_yfinance_fallback()
    if not stocks:
        print('[TWSE] 所有方案均失敗')
    return stocks


# ══════════════════════════════════════════════════════
# TPEX 上櫃股資料
# ══════════════════════════════════════════════════════
def get_tpex_stocks():
    stocks = []
    try:
        url  = 'https://www.tpex.org.tw/openapi/v1/tpex_mainboard_quotes'
        resp = requests.get(url, headers=BASE_HEADERS, timeout=20)
        resp.raise_for_status()
        for item in resp.json():
            try:
                code = str(item.get('SecuritiesCompanyCode', '')).strip()
                if not re.match(r'^\d{4,5}$', code):
                    continue
                name   = str(item.get('CompanyName', '')).strip()
                close  = float(str(item.get('Close',  '0')).replace(',', '') or 0)
                change = float(str(item.get('Change', '0')).replace(',', '') or 0)
                if close <= 0:
                    continue
                prev = close - change
                if prev <= 0:
                    continue
                chg = round((change / prev) * 100, 2)
                if chg <= 0:
                    continue
                vol = int(float(str(item.get('TradeVolume', '0')).replace(',', '') or 0))
                stocks.append({
                    'code': code, 'symbol': f'{code}.TWO', 'name': name,
                    'change_pct': chg, 'price': close, 'volume': vol, 'market': '上櫃',
                })
            except Exception:
                continue
        print(f'[TPEX] 上漲：{len(stocks)} 支')
    except Exception as e:
        print(f'[TPEX] 失敗：{e}')
    return stocks


# ══════════════════════════════════════════════════════
# 合併排行
# ══════════════════════════════════════════════════════
def get_ranking_stocks(top_n=100, mode='both'):
    if mode == 'twse':
        s = get_twse_stocks()
        s.sort(key=lambda x: -x['change_pct'])
        return s[:top_n]
    if mode == 'tpex':
        s = get_tpex_stocks()
        s.sort(key=lambda x: -x['change_pct'])
        return s[:top_n]
    if mode == 'combined':
        s = get_twse_stocks() + get_tpex_stocks()
        s.sort(key=lambda x: -x['change_pct'])
        return s[:top_n]
    # both
    tw = get_twse_stocks(); tw.sort(key=lambda x: -x['change_pct'])
    tp = get_tpex_stocks(); tp.sort(key=lambda x: -x['change_pct'])
    return tw[:top_n] + tp[:top_n]


# ══════════════════════════════════════════════════════
# 取得 OHLCV
# ══════════════════════════════════════════════════════
def get_ohlcv(symbol, days=90):
    try:
        hist = yf.Ticker(symbol).history(period=f'{days}d')
        if hist.empty:
            alt  = symbol.replace('.TW', '.TWO') if symbol.endswith('.TW') else symbol.replace('.TWO', '.TW')
            hist = yf.Ticker(alt).history(period=f'{days}d')
            if not hist.empty:
                return hist, alt
        return (hist, symbol) if not hist.empty else (None, symbol)
    except Exception:
        return None, symbol


# ══════════════════════════════════════════════════════
# 技術指標
# ══════════════════════════════════════════════════════
def calc_kd(high, low, close, n=9, m=3):
    low_n  = low.rolling(n).min()
    high_n = high.rolling(n).max()
    denom  = high_n - low_n
    rsv    = pd.Series(
        np.where(denom == 0, 50.0, (close - low_n) / denom * 100),
        index=close.index
    )
    K = pd.Series(50.0, index=close.index, dtype=float)
    D = pd.Series(50.0, index=close.index, dtype=float)
    for i in range(1, len(rsv)):
        K.iloc[i] = (m - 1) / m * K.iloc[i - 1] + 1 / m * rsv.iloc[i]
        D.iloc[i] = (m - 1) / m * D.iloc[i - 1] + 1 / m * K.iloc[i]

    k_cur, k_prv = round(K.iloc[-1], 2), round(K.iloc[-2], 2)
    d_cur, d_prv = round(D.iloc[-1], 2), round(D.iloc[-2], 2)
    golden = (k_prv < d_prv) and (k_cur > d_cur)
    death  = (k_prv > d_prv) and (k_cur < d_cur)
    return dict(
        k=k_cur, d=d_cur,
        kd_golden=golden, kd_death=death,
        kd_oversold=(k_cur < 20), kd_overbought=(k_cur > 80),
        kd_k_above_d=(k_cur > d_cur),
        kd_signal=('黃金交叉' if golden else '死亡交叉' if death else
                   '超賣區' if k_cur < 20 else '多頭排列' if k_cur > d_cur else '空頭排列'),
        kd_k_series=K, kd_d_series=D,
    )


def calc_macd(close, fast=12, slow=26, sig=9):
    ema_f = close.ewm(span=fast, adjust=False).mean()
    ema_s = close.ewm(span=slow, adjust=False).mean()
    ml    = ema_f - ema_s
    sl    = ml.ewm(span=sig, adjust=False).mean()
    hist  = ml - sl
    m_cur, m_prv = round(ml.iloc[-1], 4), round(ml.iloc[-2], 4)
    s_cur, s_prv = round(sl.iloc[-1], 4), round(sl.iloc[-2], 4)
    h_cur, h_prv = round(hist.iloc[-1], 4), round(hist.iloc[-2], 4)
    golden = (m_prv < s_prv) and (m_cur > s_cur)
    death  = (m_prv > s_prv) and (m_cur < s_cur)
    return dict(
        macd=m_cur, macd_sig_val=s_cur, macd_hist=h_cur,
        macd_golden=golden, macd_death=death,
        macd_above_zero=(m_cur > 0),
        macd_hist_expand=(h_cur > 0 and h_cur > h_prv),
        macd_signal_str=('黃金交叉' if golden else '死亡交叉' if death else
                         '零軸上方' if m_cur > 0 else '零軸下方'),
        macd_line=ml, macd_signal_line=sl, macd_histogram=hist,
    )


def calc_bband(close, n=20, k=2):
    mid   = close.rolling(n).mean()
    std   = close.rolling(n).std()
    upper = mid + k * std
    lower = mid - k * std
    c = close.iloc[-1]
    u = round(upper.iloc[-1], 2)
    m = round(mid.iloc[-1],   2)
    l = round(lower.iloc[-1], 2)
    pos = round((c - l) / (u - l), 3) if (u - l) > 0 else 0.5
    return dict(
        bb_upper=u, bb_mid=m, bb_lower=l,
        bb_width=round((u - l) / m * 100, 2) if m else 0,
        bb_pos=pos,
        bb_near_upper=(c >= u * 0.99),
        bb_near_lower=(c <= l * 1.01),
        bb_above_mid=(c > m),
        bb_signal=('突破上軌' if c >= u * 0.99 else '接近下軌' if c <= l * 1.01 else
                   '中軌以上' if c > m else '中軌以下'),
        bb_upper_series=upper, bb_mid_series=mid, bb_lower_series=lower,
    )


def calc_all_ta(hist):
    """計算所有技術指標，回傳 dict（排除 series 欄位）"""
    try:
        if hist is None or len(hist) < 30:
            return {}
        kd   = calc_kd(hist['High'], hist['Low'], hist['Close'])
        macd = calc_macd(hist['Close'])
        bb   = calc_bband(hist['Close'])
        # 只保留純量值（去掉 series）
        skip = {'kd_k_series', 'kd_d_series',
                'macd_line', 'macd_signal_line', 'macd_histogram',
                'bb_upper_series', 'bb_mid_series', 'bb_lower_series'}
        out = {}
        for d in (kd, macd, bb):
            for k, v in d.items():
                if k not in skip:
                    out[k] = v
        return out
    except Exception as e:
        print(f'  [TA] 計算失敗：{e}')
        return {}


# ══════════════════════════════════════════════════════
# 均量條件
# ══════════════════════════════════════════════════════
def analyze_volume_condition(vol_series, min_days=2, max_days=5):
    if vol_series is None or len(vol_series) < 25:
        return None
    ma5  = vol_series.rolling(5).mean()
    ma20 = vol_series.rolling(20).mean()
    df   = pd.DataFrame({'ma5': ma5, 'ma20': ma20}).dropna()
    if df.empty:
        return None
    cond        = df['ma5'] > df['ma20']
    consecutive = sum(1 for _ in iter(lambda: next(
        (False for v in reversed(cond.values) if not v), True), False))
    # 重新算（上面寫法太花俏，用簡單迴圈）
    consecutive = 0
    for v in reversed(cond.values):
        if v:
            consecutive += 1
        else:
            break
    if not (min_days <= consecutive <= max_days):
        return None
    ma5_v  = int(round(df['ma5'].iloc[-1]))
    ma20_v = int(round(df['ma20'].iloc[-1]))
    return {
        'consecutive_days': consecutive,
        'ma5_volume':  ma5_v,
        'ma20_volume': ma20_v,
        'ratio': round(ma5_v / ma20_v, 3) if ma20_v else 0,
    }


# ══════════════════════════════════════════════════════
# 背景分析任務
# ══════════════════════════════════════════════════════
def run_analysis_task(task_id):
    task = tasks[task_id]
    task['status'] = 'fetching'
    start_time     = time.time()

    try:
        mode = task.get('mode', 'both')
        mode_labels = {
            'twse': '上市前100', 'tpex': '上櫃前100',
            'combined': '合併前100', 'both': '上市前100+上櫃前100',
        }
        task['msg'] = f"正在取得【{mode_labels.get(mode, mode)}】排行榜..."

        stocks = get_ranking_stocks(top_n=100, mode=mode)

        if not stocks:
            task['status'] = 'error'
            task['msg']    = '無法取得排行資料（TWSE/TPEX API 離線或盤中無資料）'
            return

        total          = len(stocks)
        task['total']  = total
        task['status'] = 'analyzing'
        task['msg']    = f'取得 {total} 支個股，計算均量＋技術指標...'

        qualified = []
        for i, stock in enumerate(stocks):
            task['current']      = i + 1
            task['current_code'] = stock['code']
            task['current_name'] = stock['name']

            hist, actual_sym = get_ohlcv(stock['symbol'], days=90)
            vol_ok = analyze_volume_condition(
                hist['Volume'] if hist is not None else None
            )
            if vol_ok:
                ta    = calc_all_ta(hist)
                entry = {**stock, 'symbol': actual_sym, **vol_ok, **ta}
                qualified.append(entry)
                task['total_found'] = len(qualified)

            time.sleep(0.15)

        qualified.sort(key=lambda x: (-x['consecutive_days'], -x['ratio']))
        elapsed = round(time.time() - start_time, 1)

        task.update({
            'status':      'done',
            'stocks':      qualified,
            'total_found': len(qualified),
            'scanned':     total,
            'elapsed':     elapsed,
            'timestamp':   datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'msg':         f'完成！找到 {len(qualified)} 支符合均量條件個股',
        })

    except Exception as e:
        import traceback
        task['status'] = 'error'
        task['msg']    = f'分析錯誤：{str(e)}'
        print(traceback.format_exc())


# ══════════════════════════════════════════════════════
# Flask Routes
# ══════════════════════════════════════════════════════
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
    hist, sym = get_ohlcv(f'{code}.TW', 90)
    if hist is None:
        hist, sym = get_ohlcv(f'{code}.TWO', 90)
    if hist is None:
        return jsonify({'error': f'無法取得 {code} 資料'}), 404

    close, high, low = hist['Close'], hist['High'], hist['Low']
    vol              = hist['Volume']
    dates            = [d.strftime('%Y-%m-%d') for d in hist.index]

    def s(series):
        return [round(float(v), 4) if pd.notna(v) else None for v in series]
    def iv(series):
        return [int(v) if pd.notna(v) else None for v in series]

    # MACD series
    macd_d   = calc_macd(close)
    ml_s     = s(macd_d['macd_line'])
    sig_s    = s(macd_d['macd_signal_line'])
    hist_s   = s(macd_d['macd_histogram'])

    # KD series
    kd_d     = calc_kd(high, low, close)
    kd_k_s   = s(kd_d['kd_k_series'])
    kd_d_s   = s(kd_d['kd_d_series'])

    # BB series
    bb_d     = calc_bband(close)

    return jsonify({
        'code': code, 'symbol': sym, 'dates': dates,
        'open':   s(hist['Open']),
        'high':   s(high), 'low': s(low), 'close': s(close),
        'volume': iv(vol),
        'ma5v':   iv(vol.rolling(5).mean()),
        'ma20v':  iv(vol.rolling(20).mean()),
        'ma5p':   s(close.rolling(5).mean()),
        'ma20p':  s(close.rolling(20).mean()),
        'macd':        ml_s,
        'macd_signal': sig_s,
        'macd_hist':   hist_s,
        'kd_k': kd_k_s, 'kd_d': kd_d_s,
        'bb_upper': s(bb_d['bb_upper_series']),
        'bb_mid':   s(bb_d['bb_mid_series']),
        'bb_lower': s(bb_d['bb_lower_series']),
    })


# ══════════════════════════════════════════════════════
# 匯出 PDF
# ══════════════════════════════════════════════════════
@app.route('/api/export/pdf', methods=['POST'])
def export_pdf():
    from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle,
                                     Paragraph, Spacer, HRFlowable)
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import ParagraphStyle
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

    FONT = 'Helvetica'
    try:
        import glob
        cands = (glob.glob('/usr/share/fonts/truetype/noto/*CJK*Regular*.ttf') +
                 glob.glob('/usr/share/fonts/truetype/wqy/*.ttf'))
        if cands:
            pdfmetrics.registerFont(TTFont('CJK', cands[0]))
            FONT = 'CJK'
    except Exception:
        pass

    def PS(name, **kw):
        return ParagraphStyle(name, fontName=FONT, **kw)

    story = [
        Paragraph('台股均量追蹤系統 掃描報告',
                  PS('T', fontSize=15, textColor=colors.HexColor('#00d4aa'), spaceAfter=4)),
        Paragraph(f'掃描時間：{ts}　範圍：{mode_label}　已掃描：{scanned} 支　符合：{len(stocks)} 支',
                  PS('S', fontSize=8, textColor=colors.HexColor('#94a3b8'), spaceAfter=6)),
        HRFlowable(width='100%', thickness=0.8, color=colors.HexColor('#1e2d47'), spaceAfter=8),
    ]

    if not stocks:
        story.append(Paragraph('本次掃描無符合條件個股', PS('E', fontSize=9, textColor=colors.grey)))
    else:
        def fv(n):
            if not n: return '-'
            if n >= 1e8: return f'{n/1e8:.1f}億'
            if n >= 1e4: return f'{n/1e3:.0f}K'
            return str(n)

        hdrs = ['代碼','名稱','市場','漲幅','連續','5MA量','比率',
                'K','D','KD訊號','MACD','柱狀圖','MACD狀態','布林位','布林訊號']
        rows = [hdrs]
        for s in stocks:
            chg = f"+{s['change_pct']:.2f}%" if s.get('change_pct',0) >= 0 else f"{s.get('change_pct',0):.2f}%"
            rows.append([
                s.get('code',''),   s.get('name',''),  s.get('market',''),
                chg,                str(s.get('consecutive_days','')),
                fv(s.get('ma5_volume')), f"{s.get('ratio',0):.2f}x",
                f"{s.get('k',0):.1f}" if s.get('k') is not None else '-',
                f"{s.get('d',0):.1f}" if s.get('d') is not None else '-',
                s.get('kd_signal','-'),
                f"{s.get('macd',0):.3f}" if s.get('macd') is not None else '-',
                f"{s.get('macd_hist',0):.3f}" if s.get('macd_hist') is not None else '-',
                s.get('macd_signal_str','-'),
                f"{s.get('bb_pos',0):.0%}" if s.get('bb_pos') is not None else '-',
                s.get('bb_signal','-'),
            ])

        cw = [c*mm for c in [14,28,12,14,13,16,14,11,11,18,16,16,18,14,18]]
        t  = Table(rows, colWidths=cw, repeatRows=1)
        t.setStyle(TableStyle([
            ('FONTNAME',  (0,0),(-1,-1), FONT), ('FONTSIZE', (0,0),(-1,-1), 7),
            ('ALIGN',     (0,0),(-1,-1), 'CENTER'), ('VALIGN', (0,0),(-1,-1), 'MIDDLE'),
            ('BACKGROUND',(0,0),(-1,0), colors.HexColor('#0f3460')),
            ('TEXTCOLOR', (0,0),(-1,0), colors.HexColor('#00d4aa')),
            ('ROWBACKGROUNDS',(0,1),(-1,-1),
             [colors.HexColor('#111827'), colors.HexColor('#0f172a')]),
            ('TEXTCOLOR', (0,1),(-1,-1), colors.HexColor('#cbd5e1')),
            ('GRID',      (0,0),(-1,-1), 0.25, colors.HexColor('#1e2d47')),
            ('TOPPADDING',(0,0),(-1,-1), 4), ('BOTTOMPADDING',(0,0),(-1,-1), 4),
            ('TEXTCOLOR', (3,1),(3,-1), colors.HexColor('#00d98b')),
        ]))
        story.append(t)

    story += [
        Spacer(1, 6*mm),
        HRFlowable(width='100%', thickness=0.4, color=colors.HexColor('#1e2d47'), spaceAfter=3),
        Paragraph('資料來源：TWSE / TPEX　本報告僅供參考，不構成投資建議',
                  PS('F', fontSize=7, textColor=colors.HexColor('#64748b'))),
    ]
    doc.build(story)
    buf.seek(0)
    return send_file(buf, mimetype='application/pdf', as_attachment=True,
                     download_name=f"stock_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf")


# ══════════════════════════════════════════════════════
# 匯出 DOCX
# ══════════════════════════════════════════════════════
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
        pr = cell._tc.get_or_add_tcPr()
        shd = OxmlElement('w:shd')
        shd.set(qn('w:fill'), hex_color); shd.set(qn('w:val'), 'clear')
        pr.append(shd)

    def set_border(cell):
        pr = cell._tc.get_or_add_tcPr()
        b  = OxmlElement('w:tcBorders')
        for side in ['top','left','bottom','right']:
            el = OxmlElement(f'w:{side}')
            el.set(qn('w:val'),'single'); el.set(qn('w:sz'),'4'); el.set(qn('w:color'),'1E2D47')
            b.append(el)
        pr.append(b)

    h = doc.add_heading('台股均量追蹤系統 — 掃描報告', 0)
    h.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r = h.runs[0]; r.font.color.rgb = RGBColor(0, 0xD4, 0xAA); r.font.size = Pt(16)

    p = doc.add_paragraph()
    for t in [f'掃描時間：{ts}', f'　範圍：{mode_label}',
               f'　已掃描：{scanned} 支', f'　符合：{len(stocks)} 支']:
        rr = p.add_run(t); rr.font.size = Pt(9); rr.font.color.rgb = RGBColor(0x94, 0xA3, 0xB8)

    doc.add_paragraph()

    if not stocks:
        doc.add_paragraph('本次掃描無符合條件個股')
    else:
        def fv(n):
            if not n: return '-'
            if n >= 1e8: return f'{n/1e8:.1f}億'
            if n >= 1e4: return f'{n/1e3:.0f}K'
            return str(n)

        hdrs   = ['代碼','名稱','市場','漲幅%','連續','5MA量','比率',
                  'K','D','KD訊號','MACD','柱狀','MACD狀態','布林位','布林訊號']
        widths = [1.6,3.2,1.2,1.6,1.4,2.2,1.6,1.3,1.3,2.2,2.0,2.0,2.2,1.8,2.2]

        tbl = doc.add_table(rows=1, cols=len(hdrs))
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl.style = 'Table Grid'
        for i, w in enumerate(widths):
            for cell in tbl.columns[i].cells:
                cell.width = Cm(w)

        for cell, txt in zip(tbl.rows[0].cells, hdrs):
            set_bg(cell, '0F3460'); set_border(cell)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(txt); run.bold = True
            run.font.size = Pt(8); run.font.color.rgb = RGBColor(0, 0xD4, 0xAA)

        for idx, s in enumerate(stocks):
            row = tbl.add_row(); row.height = Cm(0.72)
            bg  = '111827' if idx % 2 == 0 else '0F172A'
            chg = f"+{s['change_pct']:.2f}%" if s.get('change_pct',0) >= 0 else f"{s.get('change_pct',0):.2f}%"
            vals = [
                s.get('code',''), s.get('name',''), s.get('market',''),
                chg, str(s.get('consecutive_days','')),
                fv(s.get('ma5_volume')), f"{s.get('ratio',0):.2f}x",
                f"{s.get('k',0):.1f}" if s.get('k') is not None else '-',
                f"{s.get('d',0):.1f}" if s.get('d') is not None else '-',
                s.get('kd_signal','-'),
                f"{s.get('macd',0):.3f}" if s.get('macd') is not None else '-',
                f"{s.get('macd_hist',0):.3f}" if s.get('macd_hist') is not None else '-',
                s.get('macd_signal_str','-'),
                f"{s.get('bb_pos',0):.0%}" if s.get('bb_pos') is not None else '-',
                s.get('bb_signal','-'),
            ]
            vcols = [
                RGBColor(0xFF,0xFF,0xFF), RGBColor(0xCB,0xD5,0xE1), RGBColor(0xFF,0xB5,0x47),
                (RGBColor(0,0xD9,0x8B) if s.get('change_pct',0)>=0 else RGBColor(0xFF,0x4D,0x6D)),
                RGBColor(0,0xD4,0xAA), RGBColor(0,0xD4,0xAA), RGBColor(0,0x84,0xFF),
                RGBColor(0xFF,0xD7,0),  RGBColor(0xFF,0xD7,0),  RGBColor(0xCB,0xD5,0xE1),
                RGBColor(0,0x84,0xFF),  RGBColor(0x94,0xA3,0xB8), RGBColor(0xCB,0xD5,0xE1),
                RGBColor(0x7C,0x3A,0xED), RGBColor(0xCB,0xD5,0xE1),
            ]
            for cell, val, vc in zip(row.cells, vals, vcols):
                set_bg(cell, bg); set_border(cell)
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                p = cell.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = p.add_run(val); run.font.size = Pt(7.5); run.font.color.rgb = vc

    doc.add_paragraph()
    fp = doc.add_paragraph('資料來源：TWSE / TPEX　本報告僅供參考，不構成投資建議')
    fp.runs[0].font.size = Pt(7.5); fp.runs[0].font.color.rgb = RGBColor(0x64,0x74,0x8B)

    buf = io.BytesIO(); doc.save(buf); buf.seek(0)
    return send_file(
        buf,
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        as_attachment=True,
        download_name=f"stock_{datetime.now().strftime('%Y%m%d_%H%M')}.docx",
    )


if __name__ == '__main__':
    os.makedirs('static', exist_ok=True)
    print('=' * 55)
    print('🚀 台股均量追蹤系統 v4')
    print('📡 TWSE (三層備援) + TPEX OpenAPI')
    print('📊 KD / MACD / 布林通道')
    print('🌐 http://localhost:5001')
    print('=' * 55)
    app.run(host='0.0.0.0', port=5001, debug=False, threaded=True)
