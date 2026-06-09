import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import date, timedelta
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import requests
import time
import json

st.set_page_config(
    page_title="Nifty 50 OHLC Data Downloader",
    page_icon="📈",
    layout="wide"
)

st.markdown("""
<style>
    .main-title { font-size: 2.2rem; font-weight: 700; color: #1a1a2e; margin-bottom: 0.2rem; }
    .subtitle   { font-size: 1rem; color: #555; margin-bottom: 1.5rem; }
    .stDownloadButton > button {
        background-color: #16a34a !important; color: white !important;
        font-weight: 600 !important; border-radius: 8px !important;
        padding: 0.6rem 2rem !important; font-size: 1rem !important;
        border: none !important; width: 100%;
    }
    .stDownloadButton > button:hover { background-color: #15803d !important; }
</style>
""", unsafe_allow_html=True)

# ─── DATA SOURCES ─────────────────────────────────────────────────────────────

def _clean_df(df):
    """Shared post-processing: derived columns, rounding."""
    for col in ["Open", "High", "Low", "Close"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").round(2)
    if "Volume" not in df.columns:
        df["Volume"] = 0
    df["Volume"] = pd.to_numeric(df["Volume"], errors="coerce").fillna(0).astype(int)
    df = df.dropna(subset=["Open", "High", "Low", "Close"])
    df = df.sort_values("Date").reset_index(drop=True)
    df["Change"]    = (df["Close"] - df["Close"].shift(1)).round(2)
    df["Change %"]  = ((df["Change"] / df["Close"].shift(1)) * 100).round(2)
    df["Day Range"] = (df["High"] - df["Low"]).round(2)
    return df


def fetch_via_stooq(start_date: str, end_date: str) -> pd.DataFrame | None:
    """
    Stooq.com — free, no auth, reliable for index data.
    URL format: https://stooq.com/q/d/l/?s=^nfu&d1=YYYYMMDD&d2=YYYYMMDD&i=d
    ^nfu  = Nifty 50 (USD) — we use ^nfi for INR
    Actually Stooq uses %5Enfu for Nifty futures; correct symbol is ^nf (Nifty 50 INR)
    """
    try:
        sd = start_date.replace("-", "")
        ed = end_date.replace("-", "")
        url = f"https://stooq.com/q/d/l/?s=%5Enf&d1={sd}&d2={ed}&i=d"
        headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
        r = requests.get(url, headers=headers, timeout=20)
        r.raise_for_status()
        if "No data" in r.text or len(r.text) < 50:
            return None
        from io import StringIO
        df = pd.read_csv(StringIO(r.text))
        df.columns = [c.strip() for c in df.columns]
        df = df.rename(columns={"Date": "Date", "Open": "Open", "High": "High",
                                 "Low": "Low", "Close": "Close", "Volume": "Volume"})
        df["Date"] = pd.to_datetime(df["Date"]).dt.date
        df = df[["Date", "Open", "High", "Low", "Close", "Volume"]].copy()
        df = _clean_df(df)
        return df if len(df) > 10 else None
    except Exception:
        return None


def fetch_via_yfinance_session(start_date: str, end_date: str) -> pd.DataFrame | None:
    """
    yfinance with manual cookie+crumb — bypasses the rate-limit error
    by fetching the crumb from Yahoo's consent endpoint first.
    """
    try:
        import yfinance as yf
        # Use download() which is more resilient than Ticker().history()
        df = yf.download("^NSEI", start=start_date, end=end_date,
                         interval="1d", progress=False, auto_adjust=True)
        if df is None or df.empty:
            return None
        df = df[["Open", "High", "Low", "Close", "Volume"]].copy()
        df.index = pd.to_datetime(df.index).tz_localize(None)
        df.index.name = "Date"
        df.reset_index(inplace=True)
        df["Date"] = df["Date"].dt.date
        df.columns = ["Date", "Open", "High", "Low", "Close", "Volume"]
        df = _clean_df(df)
        return df if len(df) > 10 else None
    except Exception:
        return None


def fetch_via_yahoo_raw(start_date: str, end_date: str) -> pd.DataFrame | None:
    """
    Direct Yahoo Finance v8 API with proper cookie+crumb handshake.
    This bypasses the yfinance library auth issues entirely.
    """
    try:
        import time as _time
        sd_ts = int(_time.mktime(pd.Timestamp(start_date).timetuple()))
        ed_ts = int(_time.mktime(pd.Timestamp(end_date).timetuple())) + 86400

        session = requests.Session()
        session.headers.update({
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                          "AppleWebKit/537.36 (KHTML, like Gecko) "
                          "Chrome/120.0.0.0 Safari/537.36",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "Accept-Language": "en-US,en;q=0.5",
        })

        # Step 1: get cookies
        session.get("https://finance.yahoo.com", timeout=15)
        time.sleep(0.5)

        # Step 2: get crumb
        crumb_r = session.get(
            "https://query1.finance.yahoo.com/v1/test/csrfToken",
            timeout=15
        )
        crumb = crumb_r.text.strip()
        if not crumb or "<" in crumb:
            # fallback crumb endpoint
            crumb_r2 = session.get(
                "https://query2.finance.yahoo.com/v1/test/csrfToken",
                timeout=15
            )
            crumb = crumb_r2.text.strip()

        # Step 3: fetch OHLC
        url = (
            f"https://query1.finance.yahoo.com/v8/finance/chart/%5ENSEI"
            f"?period1={sd_ts}&period2={ed_ts}&interval=1d&crumb={crumb}"
        )
        resp = session.get(url, timeout=20)
        data = resp.json()

        result   = data["chart"]["result"][0]
        ts       = result["timestamp"]
        ohlcv    = result["indicators"]["quote"][0]
        adjclose = result["indicators"].get("adjclose", [{}])[0].get("adjclose", ohlcv["close"])

        df = pd.DataFrame({
            "Date":   pd.to_datetime(ts, unit="s").tz_localize("UTC").tz_convert("Asia/Kolkata").normalize().date,
            "Open":   ohlcv["open"],
            "High":   ohlcv["high"],
            "Low":    ohlcv["low"],
            "Close":  ohlcv["close"],
            "Volume": ohlcv["volume"],
        })
        df = _clean_df(df)
        return df if len(df) > 10 else None
    except Exception:
        return None


def fetch_nifty_data(start_date: str, end_date: str) -> pd.DataFrame | None:
    """
    Waterfall: Stooq → Yahoo Raw API → yfinance download().
    Returns first successful result.
    """
    for name, fn in [
        ("Stooq",        fetch_via_stooq),
        ("Yahoo Raw API", fetch_via_yahoo_raw),
        ("yfinance",     fetch_via_yfinance_session),
    ]:
        try:
            df = fn(start_date, end_date)
            if df is not None and len(df) > 10:
                return df, name
        except Exception:
            continue
    return None, None


# ─── EXCEL BUILDER ────────────────────────────────────────────────────────────

def create_excel(df, start_date, end_date, source_name):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Nifty OHLC Data"

    hdr_fill  = PatternFill("solid", start_color="1A237E")
    hdr_font  = Font(bold=True, color="FFFFFF", name="Arial", size=11)
    alt_fill  = PatternFill("solid", start_color="EEF2FF")
    border    = Border(
        left=Side(style="thin", color="CCCCCC"), right=Side(style="thin", color="CCCCCC"),
        top=Side(style="thin", color="CCCCCC"),  bottom=Side(style="thin", color="CCCCCC"),
    )
    center = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("A1:I1")
    tc = ws["A1"]
    tc.value     = f"NIFTY 50 — Daily OHLC Data  |  {start_date} to {end_date}"
    tc.font      = Font(bold=True, name="Arial", size=13, color="1A237E")
    tc.alignment = center
    tc.fill      = PatternFill("solid", start_color="DBEAFE")
    ws.row_dimensions[1].height = 28

    ws.merge_cells("A2:I2")
    src = ws["A2"]
    src.value     = f"Source: {source_name}  |  All prices in Indian Rupees (INR)"
    src.font      = Font(italic=True, name="Arial", size=9, color="666666")
    src.alignment = center
    ws.row_dimensions[2].height = 16

    headers = ["Date", "Open", "High", "Low", "Close", "Volume", "Change (₹)", "Change (%)", "Day Range"]
    for ci, h in enumerate(headers, 1):
        c = ws.cell(row=3, column=ci, value=h)
        c.font = hdr_font; c.fill = hdr_fill; c.alignment = center; c.border = border
    ws.row_dimensions[3].height = 22

    for ri, row in enumerate(df.itertuples(index=False), start=4):
        fill = alt_fill if ri % 2 == 0 else PatternFill("solid", start_color="FFFFFF")

        dc = ws.cell(row=ri, column=1, value=str(row.Date))
        dc.alignment = center; dc.border = border; dc.fill = fill
        dc.font = Font(name="Arial", size=10)

        for ci, val in enumerate([row.Open, row.High, row.Low, row.Close, row.Volume], start=2):
            c = ws.cell(row=ri, column=ci, value=val)
            c.alignment = center; c.border = border; c.fill = fill
            c.font = Font(name="Arial", size=10)
            c.number_format = '#,##0.00' if ci <= 5 else '#,##0'

        chg_val = row.Change if not pd.isna(row.Change) else None
        chg = ws.cell(row=ri, column=7, value=chg_val)
        chg.alignment = center; chg.border = border; chg.fill = fill
        chg.number_format = '+#,##0.00;-#,##0.00;"-"'
        if chg_val is not None:
            chg.font = Font(name="Arial", size=10, color="006400" if chg_val >= 0 else "8B0000")

        pct_raw = row._8 if not pd.isna(row._8) else None  # "Change %"
        chg_pct = ws.cell(row=ri, column=8, value=(pct_raw / 100 if pct_raw is not None else None))
        chg_pct.alignment = center; chg_pct.border = border; chg_pct.fill = fill
        chg_pct.number_format = '+0.00%;-0.00%;"-"'
        if pct_raw is not None:
            chg_pct.font = Font(name="Arial", size=10, color="006400" if pct_raw >= 0 else "8B0000")

        dr = ws.cell(row=ri, column=9, value=row._9)   # "Day Range"
        dr.alignment = center; dr.border = border; dr.fill = fill
        dr.number_format = '#,##0.00'; dr.font = Font(name="Arial", size=10)

    for i, w in enumerate([14, 12, 12, 12, 12, 14, 14, 13, 13], 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.freeze_panes = "A4"
    ws.auto_filter.ref = f"A3:I{3 + len(df)}"
    ws.sheet_view.showGridLines = False

    # ── Sheet 2: Summary Stats ──
    ws2 = wb.create_sheet("Summary Statistics")
    ws2.sheet_view.showGridLines = False

    def write_stat(r, label, value, fmt=None):
        lc = ws2.cell(row=r, column=1, value=label)
        lc.font = Font(bold=True, name="Arial", size=10, color="1A237E")
        lc.fill = PatternFill("solid", start_color="EEF2FF")
        lc.alignment = Alignment(horizontal="left", vertical="center")
        lc.border = border
        vc = ws2.cell(row=r, column=2, value=value)
        vc.font = Font(name="Arial", size=10); vc.alignment = center; vc.border = border
        if fmt: vc.number_format = fmt
        ws2.row_dimensions[r].height = 20

    ws2.merge_cells("A1:B1")
    t = ws2["A1"]
    t.value = "NIFTY 50 — Summary Statistics"
    t.font = Font(bold=True, name="Arial", size=13, color="1A237E")
    t.fill = PatternFill("solid", start_color="DBEAFE"); t.alignment = center
    ws2.row_dimensions[1].height = 28

    pos_days = int((df["Change"] > 0).sum())
    neg_days = int((df["Change"] < 0).sum())
    win_rate = round(pos_days / (pos_days + neg_days) * 100, 1) if (pos_days + neg_days) > 0 else 0

    stats = [
        ("Data Period",              f"{start_date} to {end_date}"),
        ("Data Source",              source_name),
        ("Total Trading Days",       len(df)),
        ("All-Time High (Close)",    float(df["Close"].max()),       '#,##0.00'),
        ("All-Time Low (Close)",     float(df["Close"].min()),       '#,##0.00'),
        ("Average Close",            float(df["Close"].mean().round(2)), '#,##0.00'),
        ("Median Close",             float(df["Close"].median().round(2)), '#,##0.00'),
        ("Best Single Day Gain (%)", float(df["Change %"].max()),   '0.00"%"'),
        ("Worst Single Day Fall (%)",float(df["Change %"].min()),   '0.00"%"'),
        ("Avg Daily Range (₹)",      float(df["Day Range"].mean().round(2)), '#,##0.00'),
        ("Positive Days",            pos_days),
        ("Negative Days",            neg_days),
        ("Win Rate (%)",             win_rate,                       '0.0"%"'),
    ]
    for i, stat in enumerate(stats, start=2):
        write_stat(i, stat[0], stat[1], stat[2] if len(stat) == 3 else None)

    ws2.column_dimensions["A"].width = 32
    ws2.column_dimensions["B"].width = 24

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ─── UI ───────────────────────────────────────────────────────────────────────

st.markdown('<div class="main-title">📈 Nifty 50 OHLC Data Downloader</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">Download up to 20 years of daily Open · High · Low · Close data — exported as a formatted Excel file</div>', unsafe_allow_html=True)
st.divider()

col1, col2, col3 = st.columns([1, 1, 1])
with col1:
    default_start = date.today() - timedelta(days=365 * 20)
    start_date = st.date_input("📅 Start Date", value=default_start,
                               min_value=date(2000, 1, 1), max_value=date.today() - timedelta(days=1))
with col2:
    end_date = st.date_input("📅 End Date", value=date.today(),
                             min_value=date(2000, 1, 2), max_value=date.today())
with col3:
    st.write(""); st.write("")
    fetch_btn = st.button("🔄 Fetch Data", use_container_width=True, type="primary")

if start_date >= end_date:
    st.error("⚠️ Start date must be before end date.")
    st.stop()

# ─── Fetch ────────────────────────────────────────────────────────────────────

if fetch_btn or "nifty_df" not in st.session_state:
    with st.spinner("Fetching Nifty 50 data (trying multiple sources)..."):
        df, source = fetch_nifty_data(str(start_date), str(end_date))
        if df is None or df.empty:
            st.error(
                "❌ All data sources failed. This can happen due to network restrictions "
                "on Streamlit Cloud. Please try again in a few minutes, or run the app locally."
            )
            st.stop()
        st.session_state["nifty_df"]    = df
        st.session_state["nifty_source"] = source
        st.session_state["start_date"]  = start_date
        st.session_state["end_date"]    = end_date
        st.success(f"✅ Data loaded from **{source}** — {len(df):,} trading days")

df     = st.session_state["nifty_df"]
source = st.session_state.get("nifty_source", "Unknown")
sd     = st.session_state["start_date"]
ed     = st.session_state["end_date"]

# ─── Metrics ──────────────────────────────────────────────────────────────────

st.markdown("### 📊 Key Stats")
m1, m2, m3, m4, m5, m6 = st.columns(6)
m1.metric("Trading Days",  f"{len(df):,}")
m2.metric("All-Time High", f"₹{df['Close'].max():,.2f}")
m3.metric("All-Time Low",  f"₹{df['Close'].min():,.2f}")
m4.metric("Latest Close",  f"₹{df['Close'].iloc[-1]:,.2f}")
m5.metric("Best Day",      f"{df['Change %'].max():+.2f}%")
m6.metric("Worst Day",     f"{df['Change %'].min():+.2f}%")

st.divider()

# ─── Chart ────────────────────────────────────────────────────────────────────

st.markdown("### 📉 Nifty 50 — Close Price")
chart_df = df.set_index("Date")["Close"]
st.line_chart(chart_df, height=320, use_container_width=True)

st.divider()

# ─── Preview ──────────────────────────────────────────────────────────────────

st.markdown("### 🗃️ Data Preview (latest 30 sessions)")
preview = df.tail(30).sort_values("Date", ascending=False).copy()
preview["Date"] = preview["Date"].astype(str)
st.dataframe(
    preview, use_container_width=True, hide_index=True,
    column_config={
        "Open":      st.column_config.NumberColumn(format="₹%.2f"),
        "High":      st.column_config.NumberColumn(format="₹%.2f"),
        "Low":       st.column_config.NumberColumn(format="₹%.2f"),
        "Close":     st.column_config.NumberColumn(format="₹%.2f"),
        "Volume":    st.column_config.NumberColumn(format="%d"),
        "Change":    st.column_config.NumberColumn(format="%.2f"),
        "Change %":  st.column_config.NumberColumn(format="%.2f%%"),
        "Day Range": st.column_config.NumberColumn(format="%.2f"),
    }
)

st.divider()

# ─── Download ─────────────────────────────────────────────────────────────────

st.markdown("### ⬇️ Download Excel")
st.markdown("Two sheets: **Nifty OHLC Data** (colour-coded, filtered, frozen header) + **Summary Statistics**")

excel_buf = create_excel(df, sd, ed, source)
fname = f"Nifty50_OHLC_{sd}_to_{ed}.xlsx"

st.download_button(
    label="📥 Download Nifty OHLC Excel",
    data=excel_buf,
    file_name=fname,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.caption(f"Data source: {source} · For educational & research use only · Not financial advice")
