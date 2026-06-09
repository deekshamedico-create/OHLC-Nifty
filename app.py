import streamlit as st
import yfinance as yf
import pandas as pd
from io import BytesIO
from datetime import date, timedelta
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="Nifty 50 OHLC Data Downloader",
    page_icon="📈",
    layout="wide"
)

# --- Styling ---
st.markdown("""
<style>
    .main-title {
        font-size: 2.2rem;
        font-weight: 700;
        color: #1a1a2e;
        margin-bottom: 0.2rem;
    }
    .subtitle {
        font-size: 1rem;
        color: #555;
        margin-bottom: 1.5rem;
    }
    .metric-card {
        background: #f0f4ff;
        border-radius: 12px;
        padding: 1rem 1.5rem;
        border-left: 4px solid #3b82f6;
    }
    .stDownloadButton > button {
        background-color: #16a34a !important;
        color: white !important;
        font-weight: 600 !important;
        border-radius: 8px !important;
        padding: 0.6rem 2rem !important;
        font-size: 1rem !important;
        border: none !important;
        width: 100%;
    }
    .stDownloadButton > button:hover {
        background-color: #15803d !important;
    }
</style>
""", unsafe_allow_html=True)


def fetch_nifty_data(start_date, end_date):
    ticker = yf.Ticker("^NSEI")
    df = ticker.history(start=start_date, end=end_date, interval="1d")
    if df.empty:
        return None
    df = df[["Open", "High", "Low", "Close", "Volume"]].copy()
    df.index = pd.to_datetime(df.index).tz_localize(None)
    df.index.name = "Date"
    df.reset_index(inplace=True)
    df["Date"] = df["Date"].dt.date
    df.columns = ["Date", "Open", "High", "Low", "Close", "Volume"]
    for col in ["Open", "High", "Low", "Close"]:
        df[col] = df[col].round(2)
    df["Change"] = (df["Close"] - df["Close"].shift(1)).round(2)
    df["Change %"] = ((df["Change"] / df["Close"].shift(1)) * 100).round(2)
    df["Day Range"] = (df["High"] - df["Low"]).round(2)
    return df


def create_excel(df, start_date, end_date):
    wb = openpyxl.Workbook()

    # --- Sheet 1: OHLC Data ---
    ws = wb.active
    ws.title = "Nifty OHLC Data"

    header_fill = PatternFill("solid", start_color="1A237E")
    header_font = Font(bold=True, color="FFFFFF", name="Arial", size=11)
    alt_fill = PatternFill("solid", start_color="EEF2FF")
    border = Border(
        left=Side(style="thin", color="CCCCCC"),
        right=Side(style="thin", color="CCCCCC"),
        top=Side(style="thin", color="CCCCCC"),
        bottom=Side(style="thin", color="CCCCCC"),
    )
    center = Alignment(horizontal="center", vertical="center")

    # Title row
    ws.merge_cells("A1:I1")
    title_cell = ws["A1"]
    title_cell.value = f"NIFTY 50 — Daily OHLC Data  |  {start_date} to {end_date}"
    title_cell.font = Font(bold=True, name="Arial", size=13, color="1A237E")
    title_cell.alignment = center
    title_cell.fill = PatternFill("solid", start_color="DBEAFE")
    ws.row_dimensions[1].height = 28

    # Source note
    ws.merge_cells("A2:I2")
    src = ws["A2"]
    src.value = "Source: Yahoo Finance (^NSEI)  |  All prices in Indian Rupees (INR)"
    src.font = Font(italic=True, name="Arial", size=9, color="666666")
    src.alignment = center

    ws.row_dimensions[2].height = 16

    headers = ["Date", "Open", "High", "Low", "Close", "Volume", "Change (₹)", "Change (%)", "Day Range"]
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=3, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = border
    ws.row_dimensions[3].height = 22

    # Data rows
    for row_idx, row in enumerate(df.itertuples(index=False), start=4):
        is_alt = (row_idx % 2 == 0)
        row_fill = alt_fill if is_alt else PatternFill("solid", start_color="FFFFFF")

        date_cell = ws.cell(row=row_idx, column=1, value=str(row.Date))
        date_cell.alignment = center
        date_cell.border = border
        date_cell.fill = row_fill
        date_cell.font = Font(name="Arial", size=10)

        for col_idx, val in enumerate([row.Open, row.High, row.Low, row.Close, row.Volume], start=2):
            c = ws.cell(row=row_idx, column=col_idx, value=val)
            c.alignment = center
            c.border = border
            c.fill = row_fill
            c.font = Font(name="Arial", size=10)
            if col_idx <= 5:
                c.number_format = '#,##0.00'
            else:
                c.number_format = '#,##0'

        # Change ₹
        chg = ws.cell(row=row_idx, column=7, value=row.Change)
        chg.alignment = center
        chg.border = border
        chg.fill = row_fill
        chg.number_format = '+#,##0.00;-#,##0.00;0.00'
        if isinstance(row.Change, float):
            chg.font = Font(name="Arial", size=10,
                            color="006400" if row.Change >= 0 else "8B0000")

        # Change %
        chg_pct = ws.cell(row=row_idx, column=8, value=row._8)
        chg_pct.alignment = center
        chg_pct.border = border
        chg_pct.fill = row_fill
        chg_pct.number_format = '+0.00%;-0.00%;0.00%'
        if isinstance(row._8, float):
            chg_pct.value = row._8 / 100
            chg_pct.font = Font(name="Arial", size=10,
                                color="006400" if row._8 >= 0 else "8B0000")

        # Day Range
        dr = ws.cell(row=row_idx, column=9, value=row._9)
        dr.alignment = center
        dr.border = border
        dr.fill = row_fill
        dr.number_format = '#,##0.00'
        dr.font = Font(name="Arial", size=10)

    # Column widths
    col_widths = [14, 12, 12, 12, 12, 14, 14, 13, 13]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.freeze_panes = "A4"
    ws.auto_filter.ref = f"A3:I{3 + len(df)}"
    ws.sheet_view.showGridLines = False

    # --- Sheet 2: Summary Stats ---
    ws2 = wb.create_sheet("Summary Statistics")
    ws2.sheet_view.showGridLines = False

    def write_stat(row, label, value, fmt=None):
        lc = ws2.cell(row=row, column=1, value=label)
        lc.font = Font(bold=True, name="Arial", size=10, color="1A237E")
        lc.fill = PatternFill("solid", start_color="EEF2FF")
        lc.alignment = Alignment(horizontal="left", vertical="center")
        lc.border = border

        vc = ws2.cell(row=row, column=2, value=value)
        vc.font = Font(name="Arial", size=10)
        vc.alignment = center
        vc.border = border
        if fmt:
            vc.number_format = fmt
        ws2.row_dimensions[row].height = 20

    ws2.merge_cells("A1:B1")
    t = ws2["A1"]
    t.value = "NIFTY 50 — Summary Statistics"
    t.font = Font(bold=True, name="Arial", size=13, color="1A237E")
    t.fill = PatternFill("solid", start_color="DBEAFE")
    t.alignment = center
    ws2.row_dimensions[1].height = 28

    stats = [
        ("Data Period", f"{start_date} to {end_date}"),
        ("Total Trading Days", len(df)),
        ("All-Time High (Close)", df["Close"].max(), '#,##0.00'),
        ("All-Time Low (Close)", df["Close"].min(), '#,##0.00'),
        ("Average Close", df["Close"].mean().round(2), '#,##0.00'),
        ("Median Close", df["Close"].median().round(2), '#,##0.00'),
        ("Highest Volume Day", str(df.loc[df["Volume"].idxmax(), "Date"])),
        ("Peak Volume", df["Volume"].max(), '#,##0'),
        ("Best Single Day Gain (%)", df["Change %"].max(), '0.00"%"'),
        ("Worst Single Day Fall (%)", df["Change %"].min(), '0.00"%"'),
        ("Avg Daily Range (₹)", df["Day Range"].mean().round(2), '#,##0.00'),
        ("Positive Days", (df["Change"] > 0).sum()),
        ("Negative Days", (df["Change"] < 0).sum()),
    ]

    for i, stat in enumerate(stats, start=2):
        if len(stat) == 3:
            write_stat(i, stat[0], stat[1], stat[2])
        else:
            write_stat(i, stat[0], stat[1])

    ws2.column_dimensions["A"].width = 30
    ws2.column_dimensions["B"].width = 22

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ─── UI ───────────────────────────────────────────────────────────────────────

st.markdown('<div class="main-title">📈 Nifty 50 OHLC Data Downloader</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">Download daily Open · High · Low · Close data for Nifty 50 — up to 20 years of history</div>', unsafe_allow_html=True)

st.divider()

col1, col2, col3 = st.columns([1, 1, 1])

with col1:
    max_start = date.today() - timedelta(days=1)
    default_start = date.today() - timedelta(days=365 * 20)
    start_date = st.date_input(
        "📅 Start Date",
        value=default_start,
        min_value=date(2000, 1, 1),
        max_value=max_start
    )

with col2:
    end_date = st.date_input(
        "📅 End Date",
        value=date.today(),
        min_value=date(2000, 1, 2),
        max_value=date.today()
    )

with col3:
    st.write("")
    st.write("")
    fetch_btn = st.button("🔄 Fetch Data", use_container_width=True, type="primary")

if start_date >= end_date:
    st.error("⚠️ Start date must be before end date.")
    st.stop()

# ─── Fetch & Display ──────────────────────────────────────────────────────────

if fetch_btn or "nifty_df" not in st.session_state:
    with st.spinner("Fetching data from Yahoo Finance..."):
        df = fetch_nifty_data(str(start_date), str(end_date))
        if df is None or df.empty:
            st.error("❌ No data returned. Please try a different date range.")
            st.stop()
        st.session_state["nifty_df"] = df
        st.session_state["start_date"] = start_date
        st.session_state["end_date"] = end_date

df = st.session_state["nifty_df"]
sd = st.session_state["start_date"]
ed = st.session_state["end_date"]

# Metrics row
st.markdown("### 📊 Key Stats")
m1, m2, m3, m4, m5 = st.columns(5)
m1.metric("Trading Days", f"{len(df):,}")
m2.metric("All-Time High", f"₹{df['Close'].max():,.2f}")
m3.metric("All-Time Low", f"₹{df['Close'].min():,.2f}")
m4.metric("Best Day", f"{df['Change %'].max():+.2f}%")
m5.metric("Worst Day", f"{df['Change %'].min():+.2f}%")

st.divider()

# Chart
st.markdown("### 📉 Nifty 50 Close Price")
chart_df = df.set_index("Date")["Close"]
st.line_chart(chart_df, height=320, use_container_width=True)

st.divider()

# Data Table
st.markdown("### 🗃️ Data Preview (last 30 rows)")
preview = df.tail(30).sort_values("Date", ascending=False).copy()
preview["Date"] = preview["Date"].astype(str)
st.dataframe(
    preview,
    use_container_width=True,
    hide_index=True,
    column_config={
        "Open":   st.column_config.NumberColumn(format="₹%.2f"),
        "High":   st.column_config.NumberColumn(format="₹%.2f"),
        "Low":    st.column_config.NumberColumn(format="₹%.2f"),
        "Close":  st.column_config.NumberColumn(format="₹%.2f"),
        "Volume": st.column_config.NumberColumn(format="%d"),
        "Change": st.column_config.NumberColumn(format="%.2f"),
        "Change %": st.column_config.NumberColumn(format="%.2f%%"),
        "Day Range": st.column_config.NumberColumn(format="%.2f"),
    }
)

st.divider()

# Download
st.markdown("### ⬇️ Download Excel")
st.markdown("The Excel file includes **two sheets**: full OHLC data with colour-coded daily changes + a summary statistics sheet.")

with st.spinner("Preparing Excel file..."):
    excel_buf = create_excel(df, sd, ed)

fname = f"Nifty50_OHLC_{sd}_to_{ed}.xlsx"
st.download_button(
    label="📥 Download Nifty OHLC Excel",
    data=excel_buf,
    file_name=fname,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.caption("Data sourced from Yahoo Finance via `yfinance`. For educational and research use only.")
