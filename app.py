import streamlit as st
import pandas as pd
from io import BytesIO, StringIO
from datetime import date, timedelta
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import requests
import time
import urllib.parse

st.set_page_config(page_title="NSE/BSE OHLC Downloader", page_icon="📈", layout="wide")

# ─── CATALOGUES ───────────────────────────────────────────────────────────────

INDICES = {
    "── BROAD MARKET ──":       None,
    "Nifty 50":                 ("^NSEI",              "%5Enf",    "Nifty 50"),
    "Sensex (BSE 30)":          ("^BSESN",             "%5Ebsesn", "Sensex"),
    "Nifty Next 50":            ("^NSMIDCP",           None,       "Nifty Next 50"),
    "Nifty 100":                ("^CNX100",            None,       "Nifty 100"),
    "Nifty 200":                ("^CNX200",            None,       "Nifty 200"),
    "Nifty 500":                ("^CNX500",            None,       "Nifty 500"),
    "── SECTORAL ──":           None,
    "Bank Nifty":               ("^NSEBANK",           None,       "Bank Nifty"),
    "Nifty Fin Services":       ("NIFTY_FIN_SERVICE.NS", None,     "Nifty Fin Services"),
    "Nifty IT":                 ("^CNXIT",             None,       "Nifty IT"),
    "Nifty Pharma":             ("^CNXPHARMA",         None,       "Nifty Pharma"),
    "Nifty Auto":               ("^CNXAUTO",           None,       "Nifty Auto"),
    "Nifty FMCG":               ("^CNXFMCG",           None,       "Nifty FMCG"),
    "Nifty Metal":              ("^CNXMETAL",          None,       "Nifty Metal"),
    "Nifty Realty":             ("^CNXREALTY",         None,       "Nifty Realty"),
    "Nifty Energy":             ("^CNXENERGY",         None,       "Nifty Energy"),
    "Nifty Infra":              ("^CNXINFRA",          None,       "Nifty Infra"),
    "── CAP SIZE ──":           None,
    "Nifty Midcap 100":         ("^NSEMDCP100",        None,       "Nifty Midcap 100"),
    "Nifty Smallcap 100":       ("^CNXSC",             None,       "Nifty Smallcap 100"),
    "── VOLATILITY ──":         None,
    "India VIX":                ("^INDIAVIX",          None,       "India VIX"),
}

# All 50 Nifty 50 constituents (Yahoo Finance .NS tickers)
NIFTY50_STOCKS = {
    "── NIFTY 50 STOCKS ──": None,
    "Reliance Industries":    ("RELIANCE.NS",  None, "Reliance Industries"),
    "HDFC Bank":              ("HDFCBANK.NS",  None, "HDFC Bank"),
    "ICICI Bank":             ("ICICIBANK.NS", None, "ICICI Bank"),
    "Infosys":                ("INFY.NS",      None, "Infosys"),
    "TCS":                    ("TCS.NS",       None, "TCS"),
    "Bharti Airtel":          ("BHARTIARTL.NS",None, "Bharti Airtel"),
    "SBI":                    ("SBIN.NS",      None, "SBI"),
    "Hindustan Unilever":     ("HINDUNILVR.NS",None, "Hindustan Unilever"),
    "ITC":                    ("ITC.NS",       None, "ITC"),
    "Kotak Mahindra Bank":    ("KOTAKBANK.NS", None, "Kotak Mahindra Bank"),
    "L&T":                    ("LT.NS",        None, "L&T"),
    "Bajaj Finance":          ("BAJFINANCE.NS",None, "Bajaj Finance"),
    "HCL Technologies":       ("HCLTECH.NS",   None, "HCL Technologies"),
    "Axis Bank":              ("AXISBANK.NS",  None, "Axis Bank"),
    "Maruti Suzuki":          ("MARUTI.NS",    None, "Maruti Suzuki"),
    "Sun Pharma":             ("SUNPHARMA.NS", None, "Sun Pharma"),
    "Wipro":                  ("WIPRO.NS",     None, "Wipro"),
    "Titan Company":          ("TITAN.NS",     None, "Titan Company"),
    "UltraTech Cement":       ("ULTRACEMCO.NS",None, "UltraTech Cement"),
    "Asian Paints":           ("ASIANPAINT.NS",None, "Asian Paints"),
    "Tech Mahindra":          ("TECHM.NS",     None, "Tech Mahindra"),
    "Nestle India":           ("NESTLEIND.NS", None, "Nestle India"),
    "Power Grid":             ("POWERGRID.NS", None, "Power Grid"),
    "NTPC":                   ("NTPC.NS",      None, "NTPC"),
    "Bajaj Finserv":          ("BAJAJFINSV.NS",None, "Bajaj Finserv"),
    "JSW Steel":              ("JSWSTEEL.NS",  None, "JSW Steel"),
    "Tata Motors":            ("TATAMOTORS.NS",None, "Tata Motors"),
    "Tata Steel":             ("TATASTEEL.NS", None, "Tata Steel"),
    "M&M":                    ("M&M.NS",       None, "M&M"),
    "Cipla":                  ("CIPLA.NS",     None, "Cipla"),
    "Dr Reddys Labs":         ("DRREDDY.NS",   None, "Dr Reddys Labs"),
    "Eicher Motors":          ("EICHERMOT.NS", None, "Eicher Motors"),
    "Grasim Industries":      ("GRASIM.NS",    None, "Grasim Industries"),
    "Hindalco":               ("HINDALCO.NS",  None, "Hindalco"),
    "Hero MotoCorp":          ("HEROMOTOCO.NS",None, "Hero MotoCorp"),
    "IndusInd Bank":          ("INDUSINDBK.NS",None, "IndusInd Bank"),
    "ONGC":                   ("ONGC.NS",      None, "ONGC"),
    "Tata Consumer":          ("TATACONSUM.NS",None, "Tata Consumer"),
    "Divis Laboratories":     ("DIVISLAB.NS",  None, "Divis Laboratories"),
    "Adani Ports":            ("ADANIPORTS.NS",None, "Adani Ports"),
    "Adani Enterprises":      ("ADANIENT.NS",  None, "Adani Enterprises"),
    "Coal India":             ("COALINDIA.NS", None, "Coal India"),
    "BEL":                    ("BEL.NS",       None, "BEL"),
    "Shriram Finance":        ("SHRIRAMFIN.NS",None, "Shriram Finance"),
    "SBI Life Insurance":     ("SBILIFE.NS",   None, "SBI Life Insurance"),
    "HDFC Life Insurance":    ("HDFCLIFE.NS",  None, "HDFC Life Insurance"),
    "Trent":                  ("TRENT.NS",     None, "Trent"),
    "Zomato":                 ("ZOMATO.NS",    None, "Zomato"),
    "Jio Financial Services": ("JIOFIN.NS",    None, "Jio Financial Services"),
    "LTIMindtree":            ("LTIM.NS",      None, "LTIMindtree"),
}

# Merged catalogue: indices first, then stocks
ALL_INSTRUMENTS = {**INDICES, **NIFTY50_STOCKS}

INTERVAL_MAP = {
    "Daily":   "1d",
    "Weekly":  "1wk",
    "Monthly": "1mo",
}

# ─── STYLES ───────────────────────────────────────────────────────────────────

st.markdown("""
<style>
    .main-title { font-size: 2.1rem; font-weight: 700; color: #1a1a2e; margin-bottom: 0.1rem; }
    .subtitle   { font-size: 0.95rem; color: #555; margin-bottom: 1rem; }
    .stDownloadButton > button {
        background-color: #16a34a !important; color: white !important;
        font-weight: 600 !important; border-radius: 8px !important;
        padding: 0.6rem 2rem !important; font-size: 1rem !important;
        border: none !important; width: 100%;
    }
    .stDownloadButton > button:hover { background-color: #15803d !important; }
</style>
""", unsafe_allow_html=True)

# ─── DATA HELPERS ─────────────────────────────────────────────────────────────

def _clean_df(df, interval_label="Daily"):
    """Normalise types, add derived columns."""
    period = {"Daily": "Day", "Weekly": "Week", "Monthly": "Month"}.get(interval_label, "Period")
    for col in ["Open", "High", "Low", "Close"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").round(2)
    if "Volume" not in df.columns:
        df["Volume"] = 0
    df["Volume"] = pd.to_numeric(df["Volume"], errors="coerce").fillna(0).astype(int)
    df = df.dropna(subset=["Open", "High", "Low", "Close"])
    df = df.sort_values("Date").reset_index(drop=True)
    df[f"{period} Change (₹)"]  = (df["Close"] - df["Close"].shift(1)).round(2)
    df[f"{period} Change (%)"]  = ((df[f"{period} Change (₹)"] / df["Close"].shift(1)) * 100).round(2)
    df["High-Low Range"]        = (df["High"] - df["Low"]).round(2)
    df["Open-Close Diff (₹)"]   = (df["Close"] - df["Open"]).abs().round(2)
    return df


def fetch_via_stooq(start_date, end_date, stooq_sym=None, **_):
    if not stooq_sym:
        return None
    try:
        sd  = start_date.replace("-", "")
        ed  = end_date.replace("-", "")
        url = f"https://stooq.com/q/d/l/?s={stooq_sym}&d1={sd}&d2={ed}&i=d"
        r   = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=20)
        r.raise_for_status()
        if "No data" in r.text or len(r.text) < 50:
            return None
        df = pd.read_csv(StringIO(r.text))
        df.columns = [c.strip() for c in df.columns]
        df["Date"] = pd.to_datetime(df["Date"]).dt.date
        df = df[["Date", "Open", "High", "Low", "Close", "Volume"]].copy()
        return df if len(df) > 5 else None
    except Exception:
        return None


def fetch_via_yahoo_raw(start_date, end_date, yahoo_ticker="^NSEI", yf_interval="1d", **_):
    try:
        sd_ts = int(pd.Timestamp(start_date).timestamp())
        ed_ts = int(pd.Timestamp(end_date).timestamp()) + 86400
        enc   = urllib.parse.quote(yahoo_ticker)
        sess  = requests.Session()
        sess.headers.update({
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
            "Accept": "text/html,application/xhtml+xml,*/*;q=0.8",
        })
        sess.get("https://finance.yahoo.com", timeout=15)
        time.sleep(0.4)
        for ep in ["https://query1.finance.yahoo.com/v1/test/csrfToken",
                   "https://query2.finance.yahoo.com/v1/test/csrfToken"]:
            crumb = sess.get(ep, timeout=12).text.strip()
            if crumb and "<" not in crumb:
                break
        url  = (f"https://query1.finance.yahoo.com/v8/finance/chart/{enc}"
                f"?period1={sd_ts}&period2={ed_ts}&interval={yf_interval}&crumb={crumb}")
        data = sess.get(url, timeout=20).json()
        res  = data["chart"]["result"][0]
        ts   = res["timestamp"]
        q    = res["indicators"]["quote"][0]
        df   = pd.DataFrame({
            "Date":   pd.to_datetime(ts, unit="s").tz_localize("UTC")
                        .tz_convert("Asia/Kolkata").normalize().date,
            "Open":   q["open"], "High": q["high"],
            "Low":    q["low"],  "Close": q["close"], "Volume": q["volume"],
        })
        return df if len(df) > 5 else None
    except Exception:
        return None


def fetch_via_yfinance(start_date, end_date, yahoo_ticker="^NSEI", yf_interval="1d", **_):
    try:
        import yfinance as yf
        df = yf.download(yahoo_ticker, start=start_date, end=end_date,
                         interval=yf_interval, progress=False, auto_adjust=True)
        if df is None or df.empty:
            return None
        if isinstance(df.columns, pd.MultiIndex):
            df.columns = df.columns.get_level_values(0)
        df = df[["Open", "High", "Low", "Close", "Volume"]].copy()
        df.index = pd.to_datetime(df.index).tz_localize(None)
        df.index.name = "Date"
        df.reset_index(inplace=True)
        df["Date"] = df["Date"].dt.date
        df.columns = ["Date", "Open", "High", "Low", "Close", "Volume"]
        return df if len(df) > 5 else None
    except Exception:
        return None


def fetch_data(instrument_label: str, start_date: str, end_date: str,
               interval_label: str = "Daily") -> tuple:
    """Waterfall across 3 sources. Returns (df, source_name)."""
    meta = ALL_INSTRUMENTS.get(instrument_label)
    if meta is None:
        return None, None
    yahoo_ticker, stooq_sym, _ = meta
    yf_interval = INTERVAL_MAP[interval_label]

    # Stooq only has daily data
    sources = []
    if interval_label == "Daily":
        sources.append(("Stooq", fetch_via_stooq,
                        {"stooq_sym": stooq_sym}))
    sources += [
        ("Yahoo Raw API", fetch_via_yahoo_raw,
         {"yahoo_ticker": yahoo_ticker, "yf_interval": yf_interval}),
        ("yfinance",      fetch_via_yfinance,
         {"yahoo_ticker": yahoo_ticker, "yf_interval": yf_interval}),
    ]

    for name, fn, kwargs in sources:
        try:
            raw = fn(start_date, end_date, **kwargs)
            if raw is not None and len(raw) > 5:
                df = _clean_df(raw.copy(), interval_label)
                return df, name
        except Exception:
            continue
    return None, None


# ─── EXCEL BUILDER ────────────────────────────────────────────────────────────

def create_excel(df, start_date, end_date, source_name,
                 instrument_label="Nifty 50", interval_label="Daily"):
    wb   = openpyxl.Workbook()
    ws   = wb.active
    tab  = f"{instrument_label[:25]} {interval_label}"[:31]
    ws.title = tab

    HDR_FILL  = PatternFill("solid", start_color="1A237E")
    HDR_FONT  = Font(bold=True, color="FFFFFF", name="Arial", size=10)
    ALT_FILL  = PatternFill("solid", start_color="EEF2FF")
    WHT_FILL  = PatternFill("solid", start_color="FFFFFF")
    BORDER    = Border(
        left=Side(style="thin", color="CCCCCC"), right=Side(style="thin", color="CCCCCC"),
        top=Side(style="thin",  color="CCCCCC"), bottom=Side(style="thin", color="CCCCCC"),
    )
    CENTER = Alignment(horizontal="center", vertical="center")

    # ── Title rows ──
    period = {"Daily":"Day","Weekly":"Week","Monthly":"Month"}.get(interval_label,"Period")
    n_cols = 10   # Date OHLC Volume ChgAbs ChgPct HiLoRange O-C-Diff
    span = get_column_letter(n_cols)

    ws.merge_cells(f"A1:{span}1")
    tc = ws["A1"]
    tc.value = (f"{instrument_label} — {interval_label} OHLC  |  "
                f"{start_date} to {end_date}")
    tc.font      = Font(bold=True, name="Arial", size=12, color="1A237E")
    tc.alignment = CENTER
    tc.fill      = PatternFill("solid", start_color="DBEAFE")
    ws.row_dimensions[1].height = 26

    ws.merge_cells(f"A2:{span}2")
    src = ws["A2"]
    src.value     = f"Source: {source_name}  |  Prices in INR (₹)  |  Interval: {interval_label}"
    src.font      = Font(italic=True, name="Arial", size=9, color="666666")
    src.alignment = CENTER
    ws.row_dimensions[2].height = 15

    # ── Headers ──
    chg_col   = f"{period} Change (₹)"
    chgpct_col= f"{period} Change (%)"
    headers   = ["Date", "Open", "High", "Low", "Close", "Volume",
                 chg_col, chgpct_col, "High-Low Range", "Open-Close Diff (₹)"]

    for ci, h in enumerate(headers, 1):
        c = ws.cell(row=3, column=ci, value=h)
        c.font = HDR_FONT; c.fill = HDR_FILL
        c.alignment = CENTER; c.border = BORDER
    ws.row_dimensions[3].height = 22

    # ── Data rows ──
    for ri, (_, row) in enumerate(df.iterrows(), start=4):
        fill = ALT_FILL if ri % 2 == 0 else WHT_FILL
        base_font = Font(name="Arial", size=10)

        def cell(col, val, fmt=None, color=None):
            c = ws.cell(row=ri, column=col, value=val)
            c.alignment = CENTER; c.border = BORDER; c.fill = fill
            c.font = Font(name="Arial", size=10, color=color or "000000")
            if fmt:
                c.number_format = fmt
            return c

        cell(1,  str(row["Date"]))
        cell(2,  float(row["Open"])  if pd.notna(row["Open"])  else None, '#,##0.00')
        cell(3,  float(row["High"])  if pd.notna(row["High"])  else None, '#,##0.00')
        cell(4,  float(row["Low"])   if pd.notna(row["Low"])   else None, '#,##0.00')
        cell(5,  float(row["Close"]) if pd.notna(row["Close"]) else None, '#,##0.00')
        cell(6,  int(row["Volume"])  if pd.notna(row["Volume"]) else None, '#,##0')

        # Change ₹ — green/red
        chg = float(row[chg_col]) if pd.notna(row[chg_col]) else None
        cell(7, chg, '+#,##0.00;-#,##0.00;"-"',
             "006400" if chg and chg >= 0 else ("8B0000" if chg else "000000"))

        # Change % — stored as fraction for Excel percentage format
        pct = float(row[chgpct_col]) if pd.notna(row[chgpct_col]) else None
        cell(8, pct / 100 if pct is not None else None,
             '+0.00%;-0.00%;"-"',
             "006400" if pct and pct >= 0 else ("8B0000" if pct else "000000"))

        # High-Low Range — neutral
        hl = float(row["High-Low Range"]) if pd.notna(row["High-Low Range"]) else None
        cell(9, hl, '#,##0.00')

        # Open-Close Diff — absolute, highlight in amber if > avg
        oc = float(row["Open-Close Diff (₹)"]) if pd.notna(row["Open-Close Diff (₹)"]) else None
        c10 = cell(10, oc, '#,##0.00')
        # Bold if unusually large body (> 1.5× average)
        avg_oc = df["Open-Close Diff (₹)"].mean()
        if oc and avg_oc and oc > 1.5 * avg_oc:
            c10.font = Font(name="Arial", size=10, bold=True, color="7B3F00")

    # ── Column widths ──
    for i, w in enumerate([13, 11, 11, 11, 11, 13, 14, 13, 14, 16], 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.freeze_panes = "A4"
    ws.auto_filter.ref = f"A3:{span}{3 + len(df)}"
    ws.sheet_view.showGridLines = False

    # ── Sheet 2: Summary Stats ──
    ws2 = wb.create_sheet("Summary Statistics")
    ws2.sheet_view.showGridLines = False

    def write_stat(r, label, value, fmt=None):
        lc = ws2.cell(row=r, column=1, value=label)
        lc.font = Font(bold=True, name="Arial", size=10, color="1A237E")
        lc.fill = PatternFill("solid", start_color="EEF2FF")
        lc.alignment = Alignment(horizontal="left", vertical="center")
        lc.border = BORDER
        vc = ws2.cell(row=r, column=2, value=value)
        vc.font = Font(name="Arial", size=10)
        vc.alignment = CENTER; vc.border = BORDER
        if fmt: vc.number_format = fmt
        ws2.row_dimensions[r].height = 20

    ws2.merge_cells("A1:B1")
    t = ws2["A1"]
    t.value = f"{instrument_label} — Summary Statistics ({interval_label})"
    t.font  = Font(bold=True, name="Arial", size=12, color="1A237E")
    t.fill  = PatternFill("solid", start_color="DBEAFE")
    t.alignment = CENTER
    ws2.row_dimensions[1].height = 26

    pos = int((df[chg_col] > 0).sum())
    neg = int((df[chg_col] < 0).sum())
    wr  = round(pos / (pos + neg) * 100, 1) if (pos + neg) > 0 else 0
    avg_oc_val = float(df["Open-Close Diff (₹)"].mean().round(2))

    stats = [
        ("Instrument",                    instrument_label),
        ("Interval",                      interval_label),
        ("Period",                        f"{start_date} to {end_date}"),
        ("Data Source",                   source_name),
        ("Total Bars",                    len(df)),
        ("All-Time High (Close)",         float(df["Close"].max()),           '#,##0.00'),
        ("All-Time Low (Close)",          float(df["Close"].min()),           '#,##0.00'),
        ("Latest Close",                  float(df["Close"].iloc[-1]),        '#,##0.00'),
        ("Average Close",                 float(df["Close"].mean().round(2)), '#,##0.00'),
        (f"Best {period} Gain (%)",       float(df[chgpct_col].max()),        '0.00"%"'),
        (f"Worst {period} Fall (%)",      float(df[chgpct_col].min()),        '0.00"%"'),
        ("Avg High-Low Range (₹)",        float(df["High-Low Range"].mean().round(2)), '#,##0.00'),
        ("Avg Open-Close Diff (₹)",       avg_oc_val,                         '#,##0.00'),
        ("Positive Periods",              pos),
        ("Negative Periods",              neg),
        ("Win Rate (%)",                  wr,                                 '0.0"%"'),
    ]
    for i, s in enumerate(stats, start=2):
        write_stat(i, s[0], s[1], s[2] if len(s) == 3 else None)

    ws2.column_dimensions["A"].width = 34
    ws2.column_dimensions["B"].width = 26

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ─── UI ───────────────────────────────────────────────────────────────────────

st.markdown('<div class="main-title">📈 NSE / BSE OHLC Data Downloader</div>',
            unsafe_allow_html=True)
st.markdown('<div class="subtitle">Indices · Nifty 50 Stocks &nbsp;|&nbsp; '
            'Daily · Weekly · Monthly &nbsp;|&nbsp; Up to 20 years &nbsp;|&nbsp; '
            'Download as formatted Excel</div>', unsafe_allow_html=True)
st.divider()

# ── Controls ──
c1, c2, c3, c4, c5 = st.columns([1.8, 0.9, 1, 1, 0.8])

with c1:
    # Split selectbox into two: category first, then instrument
    cat = st.radio("Category", ["Indices", "Nifty 50 Stocks"], horizontal=True,
                   label_visibility="collapsed")

with c2:
    interval_label = st.selectbox("⏱ Interval", list(INTERVAL_MAP.keys()), index=0)

with c3:
    default_start = date.today() - timedelta(days=365 * 20)
    start_date = st.date_input("📅 From", value=default_start,
                               min_value=date(2000, 1, 1),
                               max_value=date.today() - timedelta(days=1))

with c4:
    end_date = st.date_input("📅 To", value=date.today(),
                             min_value=date(2000, 1, 2), max_value=date.today())

with c5:
    st.write(""); st.write("")
    fetch_btn = st.button("🔄 Fetch", use_container_width=True, type="primary")

# Instrument selector based on category
if cat == "Indices":
    valid_keys = [k for k, v in INDICES.items() if v is not None]
    instrument_label = st.selectbox("🏦 Select Index", valid_keys, index=0)
else:
    valid_keys = [k for k, v in NIFTY50_STOCKS.items() if v is not None]
    instrument_label = st.selectbox("🏢 Select Stock", valid_keys, index=0)

if start_date >= end_date:
    st.error("⚠️ Start date must be before end date.")
    st.stop()

# ── Fetch ──
cache_key = f"{instrument_label}|{interval_label}|{start_date}|{end_date}"
if fetch_btn or st.session_state.get("cache_key") != cache_key:
    meta = ALL_INSTRUMENTS.get(instrument_label)
    display_name = meta[2] if meta else instrument_label
    with st.spinner(f"Fetching {display_name} ({interval_label}) …"):
        df, source = fetch_data(instrument_label, str(start_date), str(end_date), interval_label)
    if df is None or df.empty:
        st.error("❌ All data sources failed. Check your internet connection or try a narrower date range.")
        st.stop()
    st.session_state.update({
        "df": df, "source": source, "sd": start_date, "ed": end_date,
        "instrument_label": instrument_label, "interval_label": interval_label,
        "cache_key": cache_key,
    })
    st.success(f"✅ **{display_name}** ({interval_label}) loaded from **{source}** — {len(df):,} bars")

if "df" not in st.session_state:
    st.info("👆 Select an instrument and click **Fetch** to load data.")
    st.stop()

df             = st.session_state["df"]
source         = st.session_state["source"]
sd             = st.session_state["sd"]
ed             = st.session_state["ed"]
instrument_label = st.session_state["instrument_label"]
interval_label = st.session_state["interval_label"]
meta           = ALL_INSTRUMENTS.get(instrument_label)
display_name   = meta[2] if meta else instrument_label
period         = {"Daily":"Day","Weekly":"Week","Monthly":"Month"}.get(interval_label,"Period")
chg_col        = f"{period} Change (₹)"
chgpct_col     = f"{period} Change (%)"

# ── Metrics ──
st.markdown("### 📊 Key Stats")
m1, m2, m3, m4, m5, m6, m7 = st.columns(7)
m1.metric("Total Bars",       f"{len(df):,}")
m2.metric("All-Time High",    f"₹{df['Close'].max():,.2f}")
m3.metric("All-Time Low",     f"₹{df['Close'].min():,.2f}")
m4.metric("Latest Close",     f"₹{df['Close'].iloc[-1]:,.2f}")
m5.metric(f"Best {period}",   f"{df[chgpct_col].max():+.2f}%")
m6.metric(f"Worst {period}",  f"{df[chgpct_col].min():+.2f}%")
m7.metric("Avg O-C Diff",     f"₹{df['Open-Close Diff (₹)'].mean():,.2f}")

st.divider()

# ── Chart ──
st.markdown(f"### 📉 {display_name} — Close Price ({interval_label})")
st.line_chart(df.set_index("Date")["Close"], height=300, use_container_width=True)

st.divider()

# ── Preview ──
n_preview = 30
st.markdown(f"### 🗃️ Latest {n_preview} Bars")
preview = df.tail(n_preview).sort_values("Date", ascending=False).copy()
preview["Date"] = preview["Date"].astype(str)
col_cfg = {
    "Open":  st.column_config.NumberColumn(format="₹%.2f"),
    "High":  st.column_config.NumberColumn(format="₹%.2f"),
    "Low":   st.column_config.NumberColumn(format="₹%.2f"),
    "Close": st.column_config.NumberColumn(format="₹%.2f"),
    "Volume": st.column_config.NumberColumn(format="%d"),
    chg_col:  st.column_config.NumberColumn(format="%.2f"),
    chgpct_col: st.column_config.NumberColumn(format="%.2f%%"),
    "High-Low Range":      st.column_config.NumberColumn(format="%.2f"),
    "Open-Close Diff (₹)": st.column_config.NumberColumn(format="%.2f"),
}
st.dataframe(preview, use_container_width=True, hide_index=True, column_config=col_cfg)

st.divider()

# ── Download ──
st.markdown("### ⬇️ Download Excel")
st.markdown("Two sheets: **OHLC Data** (colour-coded, auto-filter, frozen header) "
            "+ **Summary Statistics**")

excel_buf = create_excel(df, sd, ed, source, instrument_label, interval_label)
safe      = instrument_label.replace(" ", "_").replace("(","").replace(")","").replace("/","-")
fname     = f"{safe}_{interval_label}_{sd}_to_{ed}.xlsx"

st.download_button(
    label=f"📥 Download {display_name} {interval_label} Excel",
    data=excel_buf, file_name=fname,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

st.caption(f"Source: {source}  ·  Educational use only  ·  Not financial advice")
