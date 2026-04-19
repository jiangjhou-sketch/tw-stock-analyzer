# 台股均量追蹤系統 v4

## 本機啟動

```bash
pip install -r requirements.txt
python app.py
# → http://localhost:5001
```

## 部署到 Railway

```bash
git init && git add . && git commit -m "init"
git remote add origin https://github.com/你的帳號/tw-stock.git
git push -u origin main
# → railway.app 連接 repo，自動部署
```

## 功能

- TWSE 上市 + TPEX 上櫃漲幅排行
- 5日均量 > 20日均量，連續 2～5 日篩選
- KD / MACD / 布林通道技術指標
- 四種掃描範圍選擇
- 匯出 PDF / Word 報告

## TWSE 備援說明

若 TWSE 直連 403（雲端部署常見），自動切換 yfinance 批次下載備援。
TPEX 使用 OpenAPI，無此問題。
