#!/usr/bin/env python3
"""
台達電 (2308) 每日交易報告
每天早上自動抓取資料並寄送 Email
"""

import yfinance as yf
import smtplib
import json
import os
import sys
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# ── 設定區 ────────────────────────────────────────────────
CONFIG_FILE = os.path.join(os.path.dirname(__file__), "tsmc_config.json")
TICKER = "2308.TW"
STOCK_NAME = "台達電"
# ──────────────────────────────────────────────────────────


def load_config():
    if os.environ.get("GMAIL_APP_PASSWORD"):
        return {
            "gmail_user":         os.environ.get("GMAIL_USER", "allensu0507@gmail.com"),
            "gmail_app_password": os.environ["GMAIL_APP_PASSWORD"],
            "recipient_email":    os.environ.get("RECIPIENT_EMAIL", "allensu0507@gmail.com"),
        }
    if not os.path.exists(CONFIG_FILE):
        print(f"❌ 找不到設定檔：{CONFIG_FILE}")
        print("   或設定環境變數 GMAIL_APP_PASSWORD / GMAIL_USER / RECIPIENT_EMAIL")
        sys.exit(1)
    with open(CONFIG_FILE, "r") as f:
        return json.load(f)


def fetch_stock_data():
    stock = yf.Ticker(TICKER)
    return stock.history(period="30d")


def calc_rsi(prices, period=14):
    delta = prices.diff()
    gain = delta.where(delta > 0, 0).rolling(window=period).mean()
    loss = (-delta.where(delta < 0, 0)).rolling(window=period).mean()
    rs = gain / loss
    return 100 - (100 / (1 + rs))


def generate_html(hist):
    if len(hist) < 2:
        raise ValueError("資料不足，無法產生報告")

    latest = hist.iloc[-1]
    prev   = hist.iloc[-2]
    closes = hist["Close"]

    close_price = round(float(latest["Close"]), 2)
    open_price  = round(float(latest["Open"]),  2)
    high_price  = round(float(latest["High"]),  2)
    low_price   = round(float(latest["Low"]),   2)
    volume      = int(latest["Volume"])
    prev_close  = round(float(prev["Close"]),   2)

    change     = round(close_price - prev_close, 2)
    change_pct = round(change / prev_close * 100, 2)

    ma5  = round(float(closes.rolling(5).mean().iloc[-1]),  2)
    ma20 = round(float(closes.rolling(20).mean().iloc[-1]), 2)
    rsi  = round(float(calc_rsi(closes).iloc[-1]), 2)

    date_str = hist.index[-1].strftime("%Y-%m-%d")

    if close_price > ma5 > ma20:
        trend, trend_color = "多頭排列 ↑", "#28a745"
    elif close_price < ma5 < ma20:
        trend, trend_color = "空頭排列 ↓", "#dc3545"
    else:
        trend, trend_color = "盤整中 →", "#e6a817"

    if rsi > 70:
        rsi_signal, rsi_color = "超買", "#dc3545"
    elif rsi < 30:
        rsi_signal, rsi_color = "超賣", "#28a745"
    else:
        rsi_signal, rsi_color = "正常", "#17a2b8"

    arrow       = "▲" if change >= 0 else "▼"
    price_color = "#c0392b" if change >= 0 else "#27ae60"

    html = f"""<!DOCTYPE html>
<html lang="zh-TW">
<head>
  <meta charset="UTF-8">
  <style>
    body      {{ font-family: Arial, "微軟正黑體", sans-serif; max-width:580px; margin:0 auto; padding:20px; background:#f0f2f5; color:#333; }}
    .wrap     {{ background:#fff; border-radius:10px; overflow:hidden; box-shadow:0 2px 8px rgba(0,0,0,.1); }}
    .header   {{ background:#2d6a4f; color:#fff; padding:22px 20px; text-align:center; }}
    .header h2{{ margin:0; font-size:20px; }}
    .header p {{ margin:6px 0 0; font-size:13px; opacity:.8; }}
    .price-box{{ text-align:center; padding:20px; border-bottom:1px solid #eee; }}
    .price    {{ font-size:40px; font-weight:700; color:{price_color}; }}
    .change   {{ font-size:18px; color:{price_color}; margin-top:4px; }}
    .section  {{ padding:16px 20px; border-bottom:1px solid #eee; }}
    .section h3{{ margin:0 0 12px; font-size:15px; color:#2d6a4f; }}
    table     {{ width:100%; border-collapse:collapse; }}
    td        {{ padding:7px 4px; font-size:14px; }}
    .lbl      {{ color:#888; width:42%; }}
    .val      {{ font-weight:600; }}
    .badge    {{ display:inline-block; padding:2px 9px; border-radius:12px; color:#fff; font-size:12px; font-weight:600; }}
    .footer   {{ text-align:center; padding:14px; font-size:11px; color:#aaa; }}
  </style>
</head>
<body>
<div class="wrap">
  <div class="header">
    <h2>📊 {STOCK_NAME} ({TICKER.replace('.TW','')}) 每日交易報告</h2>
    <p>{date_str}　資料來源：Yahoo Finance</p>
  </div>

  <div class="price-box">
    <div class="price">NT$ {close_price:,.1f}</div>
    <div class="change">{arrow} {abs(change):.2f}　({change_pct:+.2f}%)</div>
  </div>

  <div class="section">
    <h3>📈 價格資訊</h3>
    <table>
      <tr><td class="lbl">昨日收盤</td><td class="val">NT$ {prev_close:,.1f}</td>
          <td class="lbl">今日開盤</td><td class="val">NT$ {open_price:,.1f}</td></tr>
      <tr><td class="lbl">最高</td><td class="val">NT$ {high_price:,.1f}</td>
          <td class="lbl">最低</td><td class="val">NT$ {low_price:,.1f}</td></tr>
      <tr><td class="lbl">成交量</td><td class="val" colspan="3">{volume:,} 股　（{volume/1000:.0f} 千股）</td></tr>
    </table>
  </div>

  <div class="section">
    <h3>🔧 技術指標</h3>
    <table>
      <tr><td class="lbl">MA5</td><td class="val">NT$ {ma5:,.1f}</td>
          <td class="lbl">MA20</td><td class="val">NT$ {ma20:,.1f}</td></tr>
      <tr><td class="lbl">RSI (14)</td>
          <td class="val">{rsi:.1f}　<span class="badge" style="background:{rsi_color}">{rsi_signal}</span></td>
          <td class="lbl">趨勢</td>
          <td class="val"><span class="badge" style="background:{trend_color}">{trend}</span></td></tr>
    </table>
  </div>

  <div class="footer">本報告由自動化系統產生，僅供參考，不構成投資建議。</div>
</div>
</body>
</html>"""

    subject = f"【{STOCK_NAME} {TICKER.replace('.TW','')}】{date_str}　NT${close_price:,.1f}　({change_pct:+.2f}%)"
    return html, subject


def send_email(cfg, subject, html_body):
    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"]    = cfg["gmail_user"]
    msg["To"]      = cfg["recipient_email"]
    msg.attach(MIMEText(html_body, "html", "utf-8"))

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(cfg["gmail_user"], cfg["gmail_app_password"])
        server.sendmail(cfg["gmail_user"], cfg["recipient_email"], msg.as_string())
    print(f"✅ 報告已寄出至 {cfg['recipient_email']}")


def main():
    cfg = load_config()

    print(f"🔄 正在取得 {STOCK_NAME} ({TICKER}) 資料...")
    hist = fetch_stock_data()

    if hist.empty:
        print("❌ 無法取得股票資料，請確認網路連線")
        sys.exit(1)

    html_body, subject = generate_html(hist)
    print(f"📝 報告主旨：{subject}")

    send_email(cfg, subject, html_body)


if __name__ == "__main__":
    main()
