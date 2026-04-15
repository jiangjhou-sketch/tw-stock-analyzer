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

@app.after_request
def add_json_header(response):
    """確保 API 路由回傳正確的 Content-Type"""
    if response.status_code == 404 and not response.content_type.startswith('application/json'):
        import json as _json
        response.data = _json.dumps({'error': f'找不到路由', 'status': 404})
        response.content_type = 'application/json'
    return response

@app.errorhandler(404)
def not_found(e):
    from flask import jsonify as _jsonify
    return _jsonify({'error': '找不到此路由', 'status': 404}), 404

@app.errorhandler(500)
def server_error(e):
    from flask import jsonify as _jsonify
    return _jsonify({'error': f'伺服器內部錯誤：{str(e)}', 'status': 500}), 500

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
    # 兼容本機和雲端路徑
    static_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'static')
    return send_from_directory(static_dir, 'index.html')


@app.route('/api/analyze/start', methods=['POST', 'GET'])
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


# ──────────────────────────────────────────────
# 匯出 PDF
# ──────────────────────────────────────────────
@app.route('/api/export/pdf', methods=['POST'])
def export_pdf():
    from flask import send_file
    from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle,
                                     Paragraph, Spacer, HRFlowable)
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.lib import colors
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    import io

    body = request.get_json(force=True)
    stocks  = body.get('stocks', [])
    ts      = body.get('timestamp', datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
    scanned = body.get('scanned', 0)

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=15*mm, rightMargin=15*mm,
        topMargin=18*mm, bottomMargin=18*mm,
    )

    # ── 嘗試載入中文字體（系統有就用，沒有就退回英文）──
    CN_FONT = 'Helvetica'
    try:
        import glob
        candidates = (
            glob.glob('/usr/share/fonts/truetype/noto/*CJK*Regular*.ttf') +
            glob.glob('/usr/share/fonts/truetype/wqy/*.ttf') +
            glob.glob('/usr/share/fonts/**/*TC*Regular*.ttf', recursive=True) +
            glob.glob('/usr/share/fonts/**/*SC*Regular*.ttf', recursive=True)
        )
        if candidates:
            pdfmetrics.registerFont(TTFont('CJK', candidates[0]))
            CN_FONT = 'CJK'
    except Exception:
        pass

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('title', fontName=CN_FONT, fontSize=16,
                                  textColor=colors.HexColor('#00d4aa'),
                                  spaceAfter=4)
    sub_style   = ParagraphStyle('sub', fontName=CN_FONT, fontSize=9,
                                  textColor=colors.HexColor('#94a3b8'),
                                  spaceAfter=10)
    footer_style = ParagraphStyle('footer', fontName=CN_FONT, fontSize=8,
                                   textColor=colors.HexColor('#64748b'))

    story = []

    # 標題
    story.append(Paragraph('台股均量追蹤系統  掃描報告', title_style))
    story.append(Paragraph(
        f'掃描時間：{ts}　　已掃描個股：{scanned} 支　　符合條件：{len(stocks)} 支',
        sub_style))
    story.append(Paragraph('篩選條件：5日均量 > 20日均量，連續 2～5 個交易日', sub_style))
    story.append(HRFlowable(width='100%', thickness=1,
                             color=colors.HexColor('#1e2d47'), spaceAfter=10))

    if not stocks:
        story.append(Paragraph('本次掃描無符合條件的個股。', sub_style))
    else:
        # 表頭
        headers = ['股票代碼', '名稱', '市場', '漲幅(%)', '連續(日)',
                   '5日均量', '20日均量', '均量比率']
        rows = [headers]
        for s in stocks:
            def fv(n):
                if not n: return '-'
                if n >= 1e8: return f"{n/1e8:.1f}億"
                if n >= 1e4: return f"{n/1e3:.0f}K"
                return str(n)
            chg = f"+{s['change_pct']:.2f}" if s['change_pct'] >= 0 else f"{s['change_pct']:.2f}"
            rows.append([
                s.get('code',''),
                s.get('name',''),
                s.get('market',''),
                chg,
                str(s.get('consecutive_days','')),
                fv(s.get('ma5_volume')),
                fv(s.get('ma20_volume')),
                f"{s.get('ratio',0):.3f}x",
            ])

        col_widths = [22*mm, 38*mm, 16*mm, 20*mm, 20*mm, 26*mm, 26*mm, 22*mm]

        t = Table(rows, colWidths=col_widths, repeatRows=1)

        # 連續天數顏色對應
        day_colors = {2: '#334155', 3: '#422006', 4: '#172554', 5: '#022c22'}

        cmd = [
            # 全體
            ('FONTNAME',  (0,0), (-1,-1), CN_FONT),
            ('FONTSIZE',  (0,0), (-1,-1), 8),
            ('ALIGN',     (0,0), (-1,-1), 'CENTER'),
            ('VALIGN',    (0,0), (-1,-1), 'MIDDLE'),
            ('ROWBACKGROUNDS', (0,1), (-1,-1),
             [colors.HexColor('#111827'), colors.HexColor('#0f172a')]),
            ('TEXTCOLOR',  (0,1), (-1,-1), colors.HexColor('#cbd5e1')),
            # 表頭
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#0f3460')),
            ('TEXTCOLOR',  (0,0), (-1,0), colors.HexColor('#00d4aa')),
            ('FONTSIZE',   (0,0), (-1,0), 8.5),
            ('BOTTOMPADDING', (0,0), (-1,0), 7),
            ('TOPPADDING',    (0,0), (-1,0), 7),
            # 格線
            ('GRID',      (0,0), (-1,-1), 0.3, colors.HexColor('#1e2d47')),
            ('BOTTOMPADDING', (0,1), (-1,-1), 5),
            ('TOPPADDING',    (0,1), (-1,-1), 5),
            # 漲幅欄位顏色
            ('TEXTCOLOR', (3,1), (3,-1), colors.HexColor('#00d98b')),
            # 均量比率顏色
            ('TEXTCOLOR', (7,1), (7,-1), colors.HexColor('#0084ff')),
        ]
        # 按連續天數為整列上色
        for row_i, s in enumerate(stocks, start=1):
            d = s.get('consecutive_days', 3)
            bg = day_colors.get(d, '#0f172a')
            cmd.append(('BACKGROUND', (4, row_i), (4, row_i),
                         colors.HexColor('#00d4aa' if d == 5 else
                                         '#0084ff' if d == 4 else
                                         '#ffb547' if d == 3 else '#334155')))
            cmd.append(('TEXTCOLOR', (4, row_i), (4, row_i), colors.black))

        t.setStyle(TableStyle(cmd))
        story.append(t)

    story.append(Spacer(1, 8*mm))
    story.append(HRFlowable(width='100%', thickness=0.5,
                             color=colors.HexColor('#1e2d47'), spaceAfter=4))
    story.append(Paragraph(
        '資料來源：台灣證券交易所 (TWSE) / 櫃買中心 (TPEX)　　本報告僅供參考，不構成投資建議',
        footer_style))

    doc.build(story)
    buf.seek(0)

    fname = f"stock_report_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf"
    return send_file(buf, mimetype='application/pdf',
                     as_attachment=True, download_name=fname)


# ──────────────────────────────────────────────
# 匯出 DOCX
# ──────────────────────────────────────────────
@app.route('/api/export/docx', methods=['POST'])
def export_docx():
    from flask import send_file
    from docx import Document
    from docx.shared import Pt, RGBColor, Cm, Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    import io

    body = request.get_json(force=True)
    stocks  = body.get('stocks', [])
    ts      = body.get('timestamp', datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
    scanned = body.get('scanned', 0)

    doc = Document()

    # ── 頁面設定 ──
    section = doc.sections[0]
    section.page_width  = Cm(29.7)
    section.page_height = Cm(21.0)
    section.left_margin = section.right_margin = Cm(1.5)
    section.top_margin  = section.bottom_margin = Cm(1.8)

    def set_cell_bg(cell, hex_color):
        tc   = cell._tc
        tcPr = tc.get_or_add_tcPr()
        shd  = OxmlElement('w:shd')
        shd.set(qn('w:fill'), hex_color)
        shd.set(qn('w:val'), 'clear')
        tcPr.append(shd)

    def set_cell_border(cell, color='1E2D47'):
        tc   = cell._tc
        tcPr = tc.get_or_add_tcPr()
        tcBorders = OxmlElement('w:tcBorders')
        for side in ['top','left','bottom','right']:
            el = OxmlElement(f'w:{side}')
            el.set(qn('w:val'), 'single')
            el.set(qn('w:sz'), '4')
            el.set(qn('w:color'), color)
            tcBorders.append(el)
        tcPr.append(tcBorders)

    # ── 標題 ──
    h = doc.add_heading('台股均量追蹤系統 — 掃描報告', 0)
    h.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = h.runs[0]
    run.font.color.rgb = RGBColor(0x00, 0xD4, 0xAA)
    run.font.size = Pt(18)

    # 副標題
    p = doc.add_paragraph()
    p.add_run(f'掃描時間：{ts}').font.size = Pt(9)
    p.add_run(f'　　已掃描：{scanned} 支　　符合條件：{len(stocks)} 支').font.size = Pt(9)
    p.add_run('\n篩選條件：5日均量 > 20日均量，連續 2～5 個交易日').font.size = Pt(9)
    for run in p.runs:
        run.font.color.rgb = RGBColor(0x94, 0xA3, 0xB8)

    doc.add_paragraph()

    if not stocks:
        doc.add_paragraph('本次掃描無符合條件的個股。')
    else:
        # ── 表格 ──
        headers = ['代碼', '名稱', '市場', '漲幅(%)', '連續(日)',
                   '5日均量', '20日均量', '均量比率']
        col_widths_cm = [2.0, 3.8, 1.6, 2.0, 2.0, 2.8, 2.8, 2.4]

        table = doc.add_table(rows=1, cols=len(headers))
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.style = 'Table Grid'

        # 設定欄寬
        for i, w in enumerate(col_widths_cm):
            for cell in table.columns[i].cells:
                cell.width = Cm(w)

        # 表頭
        hdr_row = table.rows[0]
        for i, (cell, txt) in enumerate(zip(hdr_row.cells, headers)):
            set_cell_bg(cell, '0F3460')
            set_cell_border(cell)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(txt)
            run.bold = True
            run.font.size = Pt(8.5)
            run.font.color.rgb = RGBColor(0x00, 0xD4, 0xAA)

        # 資料列
        day_colors_hex = {5:'022C22', 4:'172554', 3:'422006', 2:'1E293B'}

        def fv(n):
            if not n: return '-'
            if n >= 1e8: return f"{n/1e8:.1f}億"
            if n >= 1e4: return f"{n/1e3:.0f}K"
            return str(n)

        for idx, s in enumerate(stocks):
            row = table.add_row()
            row.height = Cm(0.75)
            d = s.get('consecutive_days', 2)
            row_bg = day_colors_hex.get(d, '0F172A')
            alt_bg = '111827' if idx % 2 == 0 else '0F172A'

            chg = f"+{s['change_pct']:.2f}" if s['change_pct'] >= 0 else f"{s['change_pct']:.2f}"
            values = [
                s.get('code',''),
                s.get('name',''),
                s.get('market',''),
                chg,
                str(d),
                fv(s.get('ma5_volume')),
                fv(s.get('ma20_volume')),
                f"{s.get('ratio',0):.3f}x",
            ]
            # 特殊顏色設定
            val_colors = [
                RGBColor(0xFF,0xFF,0xFF),   # 代碼 - 白
                RGBColor(0xCB,0xD5,0xE1),   # 名稱 - 淡灰
                RGBColor(0xFF,0xB5,0x47),   # 市場 - 橙
                (RGBColor(0x00,0xD9,0x8B) if s['change_pct'] >= 0
                 else RGBColor(0xFF,0x4D,0x6D)),  # 漲幅
                RGBColor(0x00,0xD4,0xAA),   # 連續天數 - 綠
                RGBColor(0x00,0xD4,0xAA),   # 5日均量 - 綠
                RGBColor(0x94,0xA3,0xB8),   # 20日均量 - 灰
                RGBColor(0x00,0x84,0xFF),   # 比率 - 藍
            ]

            for i, (cell, val, vc) in enumerate(zip(row.cells, values, val_colors)):
                set_cell_bg(cell, alt_bg)
                set_cell_border(cell)
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                p = cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = p.add_run(val)
                run.font.size = Pt(8)
                run.font.color.rgb = vc
                if i == 0:
                    run.bold = True

    # 頁尾
    doc.add_paragraph()
    footer_p = doc.add_paragraph(
        '資料來源：台灣證券交易所 (TWSE) / 櫃買中心 (TPEX)　　本報告僅供參考，不構成投資建議')
    for run in footer_p.runs:
        run.font.size = Pt(8)
        run.font.color.rgb = RGBColor(0x64, 0x74, 0x8B)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)

    fname = f"stock_report_{datetime.now().strftime('%Y%m%d_%H%M')}.docx"
    return send_file(buf,
                     mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                     as_attachment=True, download_name=fname)


if __name__ == '__main__':
    os.makedirs('static', exist_ok=True)
    print("=" * 50)
    print("🚀 台股均量追蹤系統 v3 啟動（雲端部署版）")
    print("📡 資料來源：TWSE + TPEX 公開 API")
    print("🌐 請開啟瀏覽器前往：http://localhost:5001")
    print("=" * 50)
    app.run(host='0.0.0.0', port=5001, debug=False, threaded=True)
