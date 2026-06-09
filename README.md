# 📈 Nifty 50 OHLC Data Downloader

A clean Streamlit dashboard to download **20 years of daily OHLC data** for Nifty 50, exported as a formatted, editable Excel file.

## Features

- ✅ Up to 20 years of daily Open · High · Low · Close · Volume
- ✅ Derived columns: Daily Change (₹), Change (%), Day Range
- ✅ Interactive date range selector
- ✅ Live Nifty close chart
- ✅ Summary stats (ATH, ATL, best/worst day, win rate)
- ✅ One-click Excel download (colour-coded, auto-filter, freeze panes)
- ✅ Two Excel sheets: full data + summary statistics

---

## 🚀 Run Locally

```bash
# 1. Clone the repo
git clone https://github.com/YOUR_USERNAME/nifty-ohlc-dashboard.git
cd nifty-ohlc-dashboard

# 2. Install dependencies
pip install -r requirements.txt

# 3. Launch the app
streamlit run app.py
```

The app will open at `http://localhost:8501`

---

## ☁️ Deploy on Streamlit Community Cloud (Free)

1. Push this repo to GitHub
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. Click **New app** → select your repo → set `app.py` as the main file
4. Click **Deploy** — done!

---

## 📁 File Structure

```
nifty-ohlc-dashboard/
├── app.py              # Main Streamlit app
├── requirements.txt    # Python dependencies
└── README.md
```

---

## 📊 Excel Output

| Sheet | Contents |
|-------|----------|
| **Nifty OHLC Data** | Full daily OHLC with formatted table, auto-filter, freeze panes, colour-coded change columns |
| **Summary Statistics** | All-time high/low, avg close, best/worst day, total positive vs negative sessions |

---

## Data Source

Yahoo Finance via [`yfinance`](https://github.com/ranaroussi/yfinance) — ticker `^NSEI`

> For educational and research purposes only. Not financial advice.
