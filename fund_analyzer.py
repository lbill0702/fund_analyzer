"""
Fund Performance Analyzer — Production-Ready Streamlit App
===========================================================
A professional interactive dashboard for comparing investment fund performance
from heterogeneous CSV data sources.
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import io
from difflib import get_close_matches
from typing import Optional, Tuple, List

# ─────────────────────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Fund Performance Analyzer",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────────────────────────────────────
# CUSTOM CSS — Refined dark finance aesthetic
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Serif+Display:ital@0;1&family=DM+Mono:wght@400;500&family=Instrument+Sans:wght@400;500;600&display=swap');

/* ── Root palette ── */
:root {
    --bg:          #0b0f14;
    --surface:     #121820;
    --border:      #1e2a36;
    --accent:      #00d4aa;
    --accent2:     #f7c948;
    --text:        #dce6f0;
    --muted:       #6b7f90;
    --danger:      #e05c5c;
    --radius:      8px;
}

/* ── App shell ── */
.stApp { background: var(--bg); color: var(--text); }
[data-testid="stSidebar"] {
    background: var(--surface) !important;
    border-right: 1px solid var(--border);
}
[data-testid="stSidebar"] * { color: var(--text) !important; }

/* ── Typography ── */
h1, h2, h3 {
    font-family: 'DM Serif Display', Georgia, serif !important;
    color: var(--text) !important;
    letter-spacing: -0.02em;
}
p, li, label, span, div {
    font-family: 'Instrument Sans', sans-serif !important;
}
code, .stCode * {
    font-family: 'DM Mono', monospace !important;
}

/* ── Header banner ── */
.app-header {
    padding: 2rem 0 1.5rem;
    border-bottom: 1px solid var(--border);
    margin-bottom: 2rem;
}
.app-header h1 {
    font-size: 2.6rem;
    margin: 0;
    background: linear-gradient(135deg, var(--accent) 0%, var(--accent2) 100%);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    background-clip: text;
}
.app-header p {
    color: var(--muted);
    font-size: 0.95rem;
    margin-top: 0.4rem;
}

/* ── Stat cards ── */
.stat-card {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    padding: 1.2rem 1.4rem;
    margin-bottom: 0.75rem;
}
.stat-card .label {
    font-size: 0.72rem;
    text-transform: uppercase;
    letter-spacing: 0.12em;
    color: var(--muted);
    margin-bottom: 0.3rem;
}
.stat-card .value {
    font-family: 'DM Mono', monospace !important;
    font-size: 1.5rem;
    font-weight: 500;
    color: var(--text);
}
.stat-card .value.positive { color: var(--accent); }
.stat-card .value.negative { color: var(--danger); }

/* ── Metric override ── */
[data-testid="stMetric"] {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    padding: 1rem 1.2rem;
}
[data-testid="stMetricLabel"] { color: var(--muted) !important; font-size: 0.8rem !important; }
[data-testid="stMetricValue"] { font-family: 'DM Mono', monospace !important; color: var(--text) !important; }

/* ── Tab styling ── */
[data-testid="stTabs"] button {
    font-family: 'Instrument Sans', sans-serif !important;
    font-weight: 500;
    color: var(--muted) !important;
    border-bottom: 2px solid transparent !important;
    background: transparent !important;
    transition: all 0.2s;
}
[data-testid="stTabs"] button[aria-selected="true"] {
    color: var(--accent) !important;
    border-bottom-color: var(--accent) !important;
}

/* ── File uploader ── */
[data-testid="stFileUploader"] {
    border: 1.5px dashed var(--border) !important;
    border-radius: var(--radius);
    background: var(--surface) !important;
}

/* ── Buttons ── */
.stButton button {
    background: var(--accent) !important;
    color: #0b0f14 !important;
    border: none !important;
    font-weight: 600 !important;
    border-radius: var(--radius) !important;
    font-family: 'Instrument Sans', sans-serif !important;
}

/* ── DataFrame tables ── */
[data-testid="stDataFrame"] {
    border: 1px solid var(--border);
    border-radius: var(--radius);
}

/* ── Divider ── */
hr { border-color: var(--border) !important; }

/* ── Empty state ── */
.empty-state {
    text-align: center;
    padding: 4rem 2rem;
    color: var(--muted);
}
.empty-state .icon { font-size: 3rem; margin-bottom: 1rem; }
.empty-state h3 { color: var(--muted) !important; font-size: 1.3rem; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# BLACK SWAN EVENTS
# ─────────────────────────────────────────────────────────────────────────────
BLACK_SWAN_EVENTS = [
    {
        "name": "Dot-com Crash",
        "start": "2000-03-10",
        "end": "2002-10-09",
        "color": "rgba(224, 92, 92, 0.12)",
        "label_color": "#e05c5c",
    },
    {
        "name": "Global Financial Crisis",
        "start": "2007-10-09",
        "end": "2009-03-09",
        "color": "rgba(247, 201, 72, 0.10)",
        "label_color": "#f7c948",
    },
    {
        "name": "COVID-19 Crash",
        "start": "2020-02-20",
        "end": "2020-04-07",
        "color": "rgba(160, 100, 220, 0.12)",
        "label_color": "#c084fc",
    },
    {
        "name": "Russia/Energy Crisis",
        "start": "2022-02-24",
        "end": "2022-10-14",
        "color": "rgba(250, 130, 60, 0.12)",
        "label_color": "#fb923c",
    },
]

# ─────────────────────────────────────────────────────────────────────────────
# COLUMN DETECTION (fuzzy heuristics)
# ─────────────────────────────────────────────────────────────────────────────
DATE_KEYWORDS = [
    "date", "day", "trading day", "trade_day", "tradeday",
    "time", "period", "as of", "asof", "timestamp",
]
PRICE_KEYWORDS = [
    "nav", "price", "close", "value", "net asset value",
    "nav_price", "closing price", "adj close", "adjusted close",
    "unit price", "fund price", "last price", "px", "settle",
]


def _best_match(columns: List[str], keywords: List[str]) -> Optional[str]:
    """Return the column name most likely matching a set of keywords."""
    cols_lower = {c.lower().strip(): c for c in columns}
    # 1. exact keyword match
    for kw in keywords:
        if kw in cols_lower:
            return cols_lower[kw]
    # 2. substring match (keyword inside column name)
    for kw in keywords:
        for col_l, col in cols_lower.items():
            if kw in col_l:
                return col
    # 3. fuzzy match
    matches = get_close_matches(
        keywords[0], list(cols_lower.keys()), n=1, cutoff=0.55
    )
    if matches:
        return cols_lower[matches[0]]
    return None


def detect_columns(df: pd.DataFrame) -> Tuple[Optional[str], Optional[str]]:
    """Auto-detect the date and price columns from a raw DataFrame."""
    cols = df.columns.tolist()
    date_col = _best_match(cols, DATE_KEYWORDS)
    price_col = _best_match(cols, PRICE_KEYWORDS)
    return date_col, price_col


# ─────────────────────────────────────────────────────────────────────────────
# DATA INGESTION & NORMALIZATION
# ─────────────────────────────────────────────────────────────────────────────

def _parse_blackrock_spreadsheetml(file_bytes: bytes) -> Tuple[Optional[pd.DataFrame], str]:
    """
    Parse a BlackRock-style SpreadsheetML .xls file (XML disguised as .xls).
    Extracts daily NAV + monthly total-return sections, stitches into a
    single total-return daily series.
    """
    import re as _re

    content = file_bytes.decode("utf-8-sig", errors="replace")
    cells = _re.findall(r"<ss:Data[^>]*>(.*?)</ss:Data>", content, _re.DOTALL)
    cells = [c.strip() for c in cells]

    date_re = _re.compile(r"^\d{2}-\w{3}-\d{4}$")

    # ── Section A: Daily NAV ─────────────────────────────────────────────────
    daily_records = []
    start_idx = next((i for i, c in enumerate(cells) if c == "NAV per Share"), None)
    if start_idx is not None:
        i = start_idx + 3
        while i < len(cells) - 2:
            d = cells[i]
            if date_re.match(d):
                try:
                    nav = float(cells[i + 1])
                    daily_records.append({"Date": d, "Price": nav})
                    i += 4
                except Exception:
                    break
            else:
                break

    df_daily = pd.DataFrame(daily_records)
    if not df_daily.empty:
        df_daily["Date"] = pd.to_datetime(df_daily["Date"], format="%d-%b-%Y")
        df_daily = df_daily.sort_values("Date").reset_index(drop=True)

    # ── Section B: Monthly Total (NAV) Return ────────────────────────────────
    monthly_records = []
    mth_start = next(
        (i for i, c in enumerate(cells) if "Monthly Total (NAV) Return" in c), None
    )
    if mth_start is not None:
        i = mth_start + 1
        while i < len(cells) - 1:
            d = cells[i]
            if _re.match(r"^\d{2}-\w", d):
                try:
                    ret = float(cells[i + 1])
                    monthly_records.append({"Date": d, "Monthly_Return": ret / 100})
                    i += 2
                except Exception:
                    break
            else:
                break

    df_m = pd.DataFrame(monthly_records)
    if not df_m.empty:
        df_m["Date"] = pd.to_datetime(df_m["Date"], format="%d-%b-%Y", errors="coerce")
        df_m = df_m.dropna(subset=["Date"]).sort_values("Date").reset_index(drop=True)

    if df_m.empty:
        return None, "❌ Could not find Monthly Total Return data in this file."

    # ── Build synthetic daily series from monthly returns ────────────────────
    df_m["CumIdx"] = (1 + df_m["Monthly_Return"]).cumprod()

    if not df_daily.empty:
        anchor_date = df_daily["Date"].min()
        last_before = df_m[df_m["Date"] < anchor_date]
        cum_at_anchor = last_before["CumIdx"].iloc[-1] if not last_before.empty else df_m["CumIdx"].iloc[0]
        scale = df_daily.iloc[0]["Price"] / cum_at_anchor
    else:
        scale = 1.0
        anchor_date = pd.Timestamp("2099-01-01")

    rows = []
    prev_cum = 1.0
    prev_date = df_m["Date"].iloc[0] - pd.DateOffset(months=1)
    for _, row in df_m.iterrows():
        end_date, end_cum = row["Date"], row["CumIdx"]
        bdays = pd.bdate_range(prev_date + pd.Timedelta(days=1), end_date)
        if len(bdays) > 0:
            interp = np.linspace(prev_cum, end_cum, len(bdays) + 1)[1:]
            for dt, ci in zip(bdays, interp):
                rows.append({"Date": dt, "Price": ci * scale})
        prev_cum, prev_date = end_cum, end_date

    df_synth = pd.DataFrame(rows)
    if df_synth.empty:
        return None, "❌ Could not build daily series from monthly data."

    if not df_daily.empty:
        df_out = pd.concat(
            [df_synth[df_synth["Date"] < anchor_date], df_daily], ignore_index=True
        )
    else:
        df_out = df_synth

    df_out = df_out.sort_values("Date").drop_duplicates("Date").reset_index(drop=True)

    # ── Section C: Distribution (dividend) history ───────────────────────────
    div_records = []
    div_start = next(
        (i for i, c in enumerate(cells) if c == "Distribution"), None
    )
    if div_start is not None:
        i = div_start + 1
        while i < len(cells) - 1:
            d = cells[i]
            if _re.match(r"^\d{2}-\w", d):
                try:
                    amt = float(cells[i + 1])
                    div_records.append({"Ex_Date": d, "Distribution": amt})
                    i += 2
                except Exception:
                    break
            else:
                break

    df_div = pd.DataFrame(div_records)
    if not df_div.empty:
        df_div["Ex_Date"] = pd.to_datetime(df_div["Ex_Date"], format="%d-%b-%Y", errors="coerce")
        df_div = df_div.dropna(subset=["Ex_Date"]).sort_values("Ex_Date").reset_index(drop=True)

    info = (
        f"SpreadsheetML · {len(df_m)} monthly returns + "
        f"{len(df_daily)} daily NAV rows stitched"
        + (f" + {len(df_div)} dividend records" if not df_div.empty else "")
    )
    return df_out, info, df_div


def _read_raw_dataframe(uploaded_file) -> Tuple[Optional[pd.DataFrame], str]:
    """Route an uploaded file to the right reader. Returns (raw_df, fmt_desc)."""
    import io as _io
    name = uploaded_file.name
    ext  = name.rsplit(".", 1)[-1].lower() if "." in name else ""
    file_bytes = uploaded_file.read()
    uploaded_file.seek(0)

    if ext in ("xls", "xlsx", "xlsm"):
        sniff = file_bytes[:500]
        if b"<?xml" in sniff or b"ss:Workbook" in sniff:
            df_out, info, df_div = _parse_blackrock_spreadsheetml(file_bytes)
            return df_out, f"BlackRock SpreadsheetML ({info})", df_div
        try:
            engine = "xlrd" if ext == "xls" else "openpyxl"
            df = pd.read_excel(_io.BytesIO(file_bytes), engine=engine)
            return df, f"{ext.upper()} (Excel)", pd.DataFrame()
        except Exception as e:
            return None, f"❌ Could not read Excel file: {e}", pd.DataFrame()

    try:
        df = pd.read_csv(_io.StringIO(file_bytes.decode("utf-8-sig", errors="replace")))
        return df, "CSV", pd.DataFrame()
    except Exception as e:
        return None, f"❌ Could not read file: {e}", pd.DataFrame()


def load_and_normalize(uploaded_file) -> Tuple[Optional[pd.DataFrame], str]:
    """
    Parse one uploaded file (CSV / XLS / XLSX / SpreadsheetML) into the
    standard long-format DataFrame.
    """
    raw, fmt_info, df_div = _read_raw_dataframe(uploaded_file)

    # SpreadsheetML path returns a pre-built Date/Price DataFrame
    if isinstance(raw, pd.DataFrame) and "SpreadsheetML" in fmt_info:
        df = raw.copy()
        df["Date"]  = pd.to_datetime(df["Date"], errors="coerce")
        df["Price"] = pd.to_numeric(df["Price"], errors="coerce")
        df.dropna(subset=["Date", "Price"], inplace=True)
        df.sort_values("Date", inplace=True)
        df.reset_index(drop=True, inplace=True)
        if df.empty:
            return None, f"❌ No valid rows after cleaning ({fmt_info})."
        fund_name = (
            uploaded_file.name.rsplit(".", 1)[0]
            .replace("_fund", "").replace("-", " ").strip()
        )
        df["Fund_Name"] = fund_name
        df["Normalized_Price"] = df["Price"] / df["Price"].iloc[0]
        msg = (
            f"✅ **{fund_name}** — {fmt_info}  \n"
            f"&nbsp;&nbsp;&nbsp;&nbsp;{len(df):,} rows · "
            f"{df['Date'].min().date()} → {df['Date'].max().date()}"
        )
        return df[["Date", "Fund_Name", "Price", "Normalized_Price"]], msg, df_div

    if raw is None:
        return None, fmt_info, pd.DataFrame()

    date_col, price_col = detect_columns(raw)

    if date_col is None:
        return None, "❌ Could not detect a Date column."
    if price_col is None:
        return None, "❌ Could not detect a Price/NAV column."

    df = raw[[date_col, price_col]].copy()
    df.columns = ["Date", "Price"]

    # Parse dates
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    # Parse prices (strip currency symbols, commas)
    df["Price"] = (
        df["Price"]
        .astype(str)
        .str.replace(r"[^\d.\-]", "", regex=True)
    )
    df["Price"] = pd.to_numeric(df["Price"], errors="coerce")

    # Drop rows missing critical fields
    df.dropna(subset=["Date", "Price"], inplace=True)
    df.sort_values("Date", inplace=True)
    df.reset_index(drop=True, inplace=True)

    if df.empty:
        return None, "❌ No valid rows remain after cleaning."

    # ── Fund name resolution (priority order) ────────────────────────────────
    # 1. Dedicated Fund_Code column  (e.g. "DBB88")
    # 2. Fund_Name column in the CSV (e.g. Chinese/English long name)
    # 3. Any column whose header contains "code" or "fund"
    # 4. Fallback: filename without extension
    fund_name = None
    name_source = "filename"

    raw_cols_lower = {c.lower().strip(): c for c in raw.columns}

    # Priority 1 — explicit Fund_Code
    if "fund_code" in raw_cols_lower:
        col = raw_cols_lower["fund_code"]
        val = raw[col].dropna().astype(str).iloc[0].strip() if not raw[col].dropna().empty else ""
        if val:
            fund_name = val
            name_source = f"`{col}` column"

    # Priority 2 — explicit Fund_Name (only if it has meaningful content)
    if fund_name is None and "fund_name" in raw_cols_lower:
        col = raw_cols_lower["fund_name"]
        val = raw[col].dropna().astype(str).iloc[0].strip() if not raw[col].dropna().empty else ""
        if val:
            fund_name = val
            name_source = f"`{col}` column"

    # Priority 3 — any column containing "code" or "ticker" or "symbol"
    if fund_name is None:
        for kw in ("code", "ticker", "symbol", "isin"):
            for col_l, col in raw_cols_lower.items():
                if kw in col_l and col_l not in (date_col.lower() if date_col else "", price_col.lower() if price_col else ""):
                    val = raw[col].dropna().astype(str).iloc[0].strip() if not raw[col].dropna().empty else ""
                    if val and len(val) < 40:  # codes are short; skip long text fields
                        fund_name = val
                        name_source = f"`{col}` column"
                        break
            if fund_name:
                break

    # Priority 4 — filename fallback
    if not fund_name:
        fund_name = uploaded_file.name.rsplit(".", 1)[0]
        name_source = "filename"

    df["Fund_Name"] = fund_name

    # Normalized price (first price = 1.0)
    first_price = df["Price"].iloc[0]
    df["Normalized_Price"] = df["Price"] / first_price

    detected_msg = (
        f"✅ **{fund_name}** — label from {name_source} · "
        f"detected `{date_col}` → Date, `{price_col}` → Price  \n"
        f"&nbsp;&nbsp;&nbsp;&nbsp;{len(df):,} rows · "
        f"{df['Date'].min().date()} → {df['Date'].max().date()}"
    )
    return df[["Date", "Fund_Name", "Price", "Normalized_Price"]], detected_msg, pd.DataFrame()


# ─────────────────────────────────────────────────────────────────────────────
# FREQUENCY DETECTION
# ─────────────────────────────────────────────────────────────────────────────
def detect_frequency(dates: pd.Series) -> int:
    """Estimate annualisation factor: 252 daily, 52 weekly, 12 monthly."""
    deltas = dates.sort_values().diff().dropna().dt.days
    median_gap = deltas.median()
    if median_gap <= 3:
        return 252
    if median_gap <= 10:
        return 52
    if median_gap <= 35:
        return 12
    return 4  # quarterly fallback


# ─────────────────────────────────────────────────────────────────────────────
# FINANCIAL ANALYTICS ENGINE
# ─────────────────────────────────────────────────────────────────────────────
WINDOWS = {
    "1Y": 1,
    "3Y": 3,
    "5Y": 5,
    "10Y": 10,
}


def compute_stats(df_fund: pd.DataFrame) -> dict:
    """
    Compute rolling-window statistics for a single fund.
    df_fund must have columns: Date, Price, Normalized_Price
    """
    df = df_fund.sort_values("Date").copy()
    df["Return"] = df["Price"].pct_change()
    freq = detect_frequency(df["Date"])
    max_date = df["Date"].max()
    results = {}

    for label, years in WINDOWS.items():
        cutoff = max_date - pd.DateOffset(years=years)
        window = df[df["Date"] >= cutoff].copy()
        if len(window) < 5:
            results[label] = None
            continue

        # Annualized volatility
        period_returns = window["Price"].pct_change().dropna()
        vol = period_returns.std() * np.sqrt(freq)

        # Cumulative performance
        first_px = window["Price"].iloc[0]
        last_px = window["Price"].iloc[-1]
        cum_perf = (last_px / first_px) - 1.0

        results[label] = {
            "annualized_volatility": vol,
            "cumulative_return": cum_perf,
            "n_obs": len(window),
            "freq": freq,
        }

    # Calendar year returns
    df["Year"] = df["Date"].dt.year
    cal_years = {}
    for year, grp in df.groupby("Year"):
        grp = grp.sort_values("Date")
        if len(grp) < 2:
            continue
        cal_ret = (grp["Price"].iloc[-1] / grp["Price"].iloc[0]) - 1.0
        cal_years[int(year)] = cal_ret
    results["calendar_years"] = cal_years

    return results


# ─────────────────────────────────────────────────────────────────────────────
# PLOTLY CHART BUILDER
# ─────────────────────────────────────────────────────────────────────────────
FUND_COLORS = [
    "#00d4aa", "#f7c948", "#60a5fa", "#f472b6",
    "#a78bfa", "#fb923c", "#34d399", "#f87171",
]


def build_performance_chart(
    combined: pd.DataFrame,
    show_black_swan: bool,
    fund_color_map: dict,
) -> go.Figure:
    fig = go.Figure()

    for fund in combined["Fund_Name"].unique():
        fd = combined[combined["Fund_Name"] == fund].sort_values("Date")
        color = fund_color_map.get(fund, "#00d4aa")
        fig.add_trace(
            go.Scatter(
                x=fd["Date"],
                y=fd["Normalized_Price"],
                mode="lines",
                name=fund,
                line=dict(color=color, width=2.2),
                hovertemplate=(
                    "<b>%{fullData.name}</b><br>"
                    "Date: %{x|%b %d, %Y}<br>"
                    "Growth: <b>%{y:.4f}x</b><br>"
                    "Return: %{customdata:.2%}<extra></extra>"
                ),
                customdata=(fd["Normalized_Price"] - 1).values,
            )
        )

    if show_black_swan:
        x_min = combined["Date"].min()
        x_max = combined["Date"].max()
        for event in BLACK_SWAN_EVENTS:
            ev_start = pd.Timestamp(event["start"])
            ev_end = pd.Timestamp(event["end"])
            # Only add if the event overlaps the data range
            if ev_end < x_min or ev_start > x_max:
                continue
            fig.add_vrect(
                x0=max(ev_start, x_min),
                x1=min(ev_end, x_max),
                fillcolor=event["color"],
                line_width=0,
                annotation_text=event["name"],
                annotation_position="top left",
                annotation=dict(
                    font_size=9.5,
                    font_color=event["label_color"],
                    font_family="DM Mono",
                    showarrow=False,
                ),
                layer="below",
            )

    fig.update_layout(
        template="plotly_dark",
        paper_bgcolor="rgba(11,15,20,0)",
        plot_bgcolor="rgba(18,24,32,0.7)",
        font=dict(family="Instrument Sans", color="#dce6f0"),
        legend=dict(
            bgcolor="rgba(18,24,32,0.85)",
            bordercolor="#1e2a36",
            borderwidth=1,
            font=dict(size=12),
        ),
        hovermode="x unified",
        margin=dict(l=10, r=10, t=30, b=10),
        xaxis=dict(
            gridcolor="#1e2a36",
            showgrid=True,
            zeroline=False,
            tickfont=dict(family="DM Mono", size=11),
        ),
        yaxis=dict(
            gridcolor="#1e2a36",
            showgrid=True,
            zeroline=False,
            tickformat=".2f",
            tickfont=dict(family="DM Mono", size=11),
            title=dict(text="Cumulative Growth (base = 1.0)", font=dict(size=12)),
        ),
        height=520,
    )
    return fig


# MSCI ACWI Min Volatility benchmark — scraped directly from BlackRock Highcharts
MSCI_ACWI_MINVOL = {
    2016: 7.43, 2017: 17.93, 2018: -1.56, 2019: 21.05, 2020: 2.69,
    2021: 13.94, 2022: -10.31, 2023: 7.74, 2024: 11.37, 2025: 10.65,
}

# BlackRock chart colours (exact from Highcharts SVG inspection)
BR_FUND_COLOR      = "#8db600"   # olive-green  — Total Return series
BR_BENCHMARK_COLOR = "#4472c4"   # steel-blue   — Benchmark 1 series
BR_BG              = "#2a2a2a"   # dark charcoal background
BR_PLOT_BG         = "#1e1e1e"   # slightly darker plot area
BR_GRID            = "#3a3a3a"   # subtle gridlines
BR_AXIS_TEXT       = "#cccccc"   # light axis labels
BR_ZERO_LINE       = "#666666"   # zero-line


def build_calendar_chart(
    stats_map: dict,
    fund_color_map: dict,
    show_benchmark: bool = True,
) -> go.Figure:
    """
    BlackRock-style grouped bar chart:
    – Dark background matching BlackRock Performance tab
    – Fund bars in olive-green, Benchmark bars in steel-blue
    – Percentage Y-axis, year X-axis
    – Optional MSCI ACWI Min Vol benchmark overlay
    – Data table rendered below chart via Plotly annotations
    """
    all_years = sorted(
        {
            year
            for stats in stats_map.values()
            if stats and stats.get("calendar_years")
            for year in stats["calendar_years"]
        }
    )
    if not all_years:
        return None

    fig = go.Figure()

    # ── Fund bars ────────────────────────────────────────────────────────────
    fund_colors = [BR_FUND_COLOR, "#5ba3e0", "#e0a85b", "#c084fc", "#fb923c"]
    for idx, (fund, stats) in enumerate(stats_map.items()):
        if not stats or not stats.get("calendar_years"):
            continue
        cy = stats["calendar_years"]
        ys = [cy.get(y) for y in all_years]
        bar_colors = [
            (fund_colors[idx % len(fund_colors)] if v is not None and v >= 0
             else "#cc3333")
            for v in ys
        ]
        fig.add_trace(go.Bar(
            x=[str(y) for y in all_years],
            y=ys,
            name=fund,
            marker=dict(
                color=bar_colors,
                line=dict(width=0),
            ),
            text=[f"{v:+.1f}%" if v is not None else "" for v in ys],
            textposition="outside",
            textfont=dict(family="Courier New, monospace", size=9.5, color="#cccccc"),
            hovertemplate=(
                "<b>%{fullData.name}</b><br>"
                "%{x}: <b>%{y:.2f}%</b><extra></extra>"
            ),
            cliponaxis=False,
        ))

    # ── Benchmark bars ───────────────────────────────────────────────────────
    if show_benchmark:
        bm_ys = [MSCI_ACWI_MINVOL.get(y) for y in all_years]
        bm_colors = [
            BR_BENCHMARK_COLOR if (v is not None and v >= 0) else "#8888cc"
            for v in bm_ys
        ]
        fig.add_trace(go.Bar(
            x=[str(y) for y in all_years],
            y=bm_ys,
            name="Benchmark (MSCI ACWI Min Vol)",
            marker=dict(color=bm_colors, line=dict(width=0)),
            text=[f"{v:+.1f}%" if v is not None else "" for v in bm_ys],
            textposition="outside",
            textfont=dict(family="Courier New, monospace", size=9.5, color="#aaaacc"),
            hovertemplate=(
                "<b>MSCI ACWI Min Vol</b><br>"
                "%{x}: <b>%{y:.2f}%</b><extra></extra>"
            ),
            cliponaxis=False,
        ))

    fig.update_layout(
        paper_bgcolor=BR_BG,
        plot_bgcolor=BR_PLOT_BG,
        font=dict(family="Instrument Sans, sans-serif", color=BR_AXIS_TEXT, size=12),
        barmode="group",
        bargap=0.28,
        bargroupgap=0.06,
        legend=dict(
            orientation="h",
            x=0.5, xanchor="center",
            y=-0.18, yanchor="top",
            bgcolor="rgba(0,0,0,0)",
            font=dict(size=11, color=BR_AXIS_TEXT),
            itemsizing="constant",
        ),
        xaxis=dict(
            showgrid=False,
            zeroline=False,
            tickfont=dict(family="Instrument Sans", size=11, color=BR_AXIS_TEXT),
            linecolor=BR_GRID,
        ),
        yaxis=dict(
            gridcolor=BR_GRID,
            gridwidth=0.5,
            ticksuffix="%",
            tickfont=dict(family="Courier New, monospace", size=10, color=BR_AXIS_TEXT),
            title=dict(text="Values", font=dict(size=11, color=BR_AXIS_TEXT), standoff=8),
            zeroline=True,
            zerolinecolor=BR_ZERO_LINE,
            zerolinewidth=1.2,
            showline=False,
        ),
        margin=dict(l=50, r=20, t=30, b=80),
        height=460,
    )
    return fig


# ─────────────────────────────────────────────────────────────────────────────
# STATS TABLE RENDERER  — portal-style (funds × years / windows)
# ─────────────────────────────────────────────────────────────────────────────
def _pct_cell(value: float | None, bold: bool = False) -> str:
    """Render a percentage value as a coloured HTML table cell."""
    if value is None:
        return "<td style='text-align:right;color:#6b7f90;'>—</td>"
    color = "#00d4aa" if value >= 0 else "#e05c5c"
    weight = "font-weight:600;" if bold else ""
    return (
        f"<td style='text-align:right;color:{color};{weight}"
        f"font-family:DM Mono,monospace;font-size:0.88rem;'>"
        f"{value:+.2%}</td>"
    )


def render_calendar_table(stats_map: dict, fund_color_map: dict):
    """
    Portal-style table: funds as rows, calendar years as columns.
    Matches the Manulife 'Performance – calendar year' layout.
    """
    # Gather all years that appear across any fund, sort descending (most recent first)
    all_years = sorted(
        {
            yr
            for stats in stats_map.values()
            if stats and stats.get("calendar_years")
            for yr in stats["calendar_years"]
        },
        reverse=True,
    )
    if not all_years:
        st.info("Not enough data to compute calendar year returns.")
        return

    # Header row
    year_headers = "".join(
        f"<th style='text-align:right;color:#6b7f90;font-weight:500;"
        f"font-size:0.82rem;padding:0.6rem 1rem;white-space:nowrap;'>{y}</th>"
        for y in all_years
    )
    header = (
        "<thead><tr>"
        "<th style='text-align:left;color:#6b7f90;font-weight:500;"
        "font-size:0.82rem;padding:0.6rem 1rem;min-width:220px;'>"
        "Fund (Code) · Currency</th>"
        + year_headers
        + "<th style='text-align:right;color:#6b7f90;font-weight:500;"
        "font-size:0.82rem;padding:0.6rem 1rem;'>Unit Price</th>"
        "</tr></thead>"
    )

    # Body rows — one per fund
    body_rows = []
    for fund in stats_map:
        stats = stats_map[fund]
        color = fund_color_map.get(fund, "#00d4aa")
        cy = stats.get("calendar_years", {}) if stats else {}

        # Fund name cell with colour dot
        name_cell = (
            f"<td style='padding:0.75rem 1rem;'>"
            f"<span style='display:inline-block;width:8px;height:8px;"
            f"border-radius:50%;background:{color};margin-right:0.5rem;"
            f"vertical-align:middle;'></span>"
            f"<span style='font-weight:500;font-size:0.9rem;color:#dce6f0;'>{fund}</span>"
            f"</td>"
        )

        year_cells = "".join(_pct_cell(cy.get(y)) for y in all_years)

        # Latest unit price from combined (passed via closure — we look it up from stats_map)
        body_rows.append(
            f"<tr style='border-top:1px solid #1e2a36;'>"
            f"{name_cell}{year_cells}"
            f"<td style='text-align:right;color:#6b7f90;font-size:0.82rem;"
            f"font-family:DM Mono,monospace;padding:0.75rem 1rem;'>—</td>"
            f"</tr>"
        )

    body = "<tbody>" + "".join(body_rows) + "</tbody>"
    table_html = (
        "<div style='overflow-x:auto;border:1px solid #1e2a36;border-radius:8px;"
        "background:#121820;'>"
        f"<table style='width:100%;border-collapse:collapse;'>{header}{body}</table>"
        "</div>"
    )
    st.markdown(table_html, unsafe_allow_html=True)


def render_rolling_table(stats_map: dict, fund_color_map: dict):
    """
    Rolling-window table: funds as rows, windows (1Y/3Y/5Y/10Y) as column pairs
    showing Cumulative Return + Annualized Volatility.
    """
    window_labels = list(WINDOWS.keys())  # ["1Y","3Y","5Y","10Y"]

    # Build double header: window label spanning 2 sub-columns
    window_headers = "".join(
        f"<th colspan='2' style='text-align:center;color:#6b7f90;font-weight:500;"
        f"font-size:0.82rem;padding:0.4rem 1rem;border-bottom:1px solid #1e2a36;"
        f"white-space:nowrap;'>{lbl}</th>"
        for lbl in window_labels
    )
    sub_headers = "".join(
        "<th style='text-align:right;color:#6b7f90;font-weight:400;"
        "font-size:0.75rem;padding:0.3rem 0.6rem;'>Cum. Ret.</th>"
        "<th style='text-align:right;color:#6b7f90;font-weight:400;"
        "font-size:0.75rem;padding:0.3rem 0.6rem;'>Ann. Vol.</th>"
        for _ in window_labels
    )
    header = (
        "<thead>"
        "<tr>"
        "<th rowspan='2' style='text-align:left;color:#6b7f90;font-weight:500;"
        "font-size:0.82rem;padding:0.6rem 1rem;min-width:200px;'>Fund</th>"
        + window_headers
        + "</tr>"
        "<tr>" + sub_headers + "</tr>"
        "</thead>"
    )

    body_rows = []
    for fund in stats_map:
        stats = stats_map[fund]
        color = fund_color_map.get(fund, "#00d4aa")
        name_cell = (
            f"<td style='padding:0.75rem 1rem;'>"
            f"<span style='display:inline-block;width:8px;height:8px;"
            f"border-radius:50%;background:{color};margin-right:0.5rem;"
            f"vertical-align:middle;'></span>"
            f"<span style='font-weight:500;font-size:0.9rem;color:#dce6f0;'>{fund}</span>"
            f"</td>"
        )
        window_cells = ""
        for lbl in window_labels:
            w = stats.get(lbl) if stats else None
            window_cells += _pct_cell(w["cumulative_return"] if w else None)
            # Volatility is always positive — show in neutral colour
            if w:
                window_cells += (
                    f"<td style='text-align:right;color:#dce6f0;"
                    f"font-family:DM Mono,monospace;font-size:0.88rem;'>"
                    f"{w['annualized_volatility']:.2%}</td>"
                )
            else:
                window_cells += "<td style='text-align:right;color:#6b7f90;'>—</td>"

        body_rows.append(
            f"<tr style='border-top:1px solid #1e2a36;'>{name_cell}{window_cells}</tr>"
        )

    body = "<tbody>" + "".join(body_rows) + "</tbody>"
    table_html = (
        "<div style='overflow-x:auto;border:1px solid #1e2a36;border-radius:8px;"
        "background:#121820;'>"
        f"<table style='width:100%;border-collapse:collapse;'>{header}{body}</table>"
        "</div>"
    )
    st.markdown(table_html, unsafe_allow_html=True)



# ─────────────────────────────────────────────────────────────────────────────
# DIVIDEND CHART BUILDER
# ─────────────────────────────────────────────────────────────────────────────
def build_dividend_chart(
    all_divs: list, fund_color_map: dict, nav_df: pd.DataFrame
) -> go.Figure:
    """
    Three-panel dividend chart:
      Panel 1: Monthly distribution amount per unit (bars) + trailing 12m total (dotted line)
      Panel 2: Annualised dividend yield = (distribution / unit_price) × 12  (line per fund)
      Panel 3: Annual total distribution (bars, full years only)
    """
    if not all_divs:
        return None

    combined_div = pd.concat(all_divs, ignore_index=True)
    combined_div["Year"] = combined_div["Ex_Date"].dt.year
    funds = combined_div["Fund_Name"].unique().tolist()

    # ── Compute monthly yield for each fund ───────────────────────────────────
    # For each ex-date, look up the NAV price on that date (or nearest prior day).
    # Yield = (distribution / nav_price) * 12
    nav_indexed = (
        nav_df[["Date", "Fund_Name", "Price"]]
        .sort_values("Date")
        .set_index(["Fund_Name", "Date"])
    )

    def get_nav_on(fund, date):
        try:
            fund_nav = nav_indexed.loc[fund]
            # Use the last available price on or before ex_date
            candidates = fund_nav[fund_nav.index <= date]
            if candidates.empty:
                return np.nan
            return float(candidates.iloc[-1]["Price"])
        except Exception:
            return np.nan

    combined_div["NAV_on_ExDate"] = combined_div.apply(
        lambda r: get_nav_on(r["Fund_Name"], r["Ex_Date"]), axis=1
    )
    combined_div["Ann_Yield"] = (
        combined_div["Distribution"] / combined_div["NAV_on_ExDate"] * 12
    )

    fig = make_subplots(
        rows=3, cols=1,
        row_heights=[0.38, 0.32, 0.30],
        vertical_spacing=0.08,
        subplot_titles=(
            "Monthly Distribution per Unit (USD)",
            "Annualised Dividend Yield  [distribution ÷ unit price × 12]",
            "Annual Total Distribution (USD)",
        ),
    )

    # ── Panel 1: distribution bars + trailing-12m line ────────────────────────
    for fund in funds:
        fd = combined_div[combined_div["Fund_Name"] == fund].sort_values("Ex_Date")
        color = fund_color_map.get(fund, "#00d4aa")

        fig.add_trace(
            go.Bar(
                x=fd["Ex_Date"],
                y=fd["Distribution"],
                name=fund,
                marker_color=color,
                opacity=0.80,
                hovertemplate=(
                    "<b>%{fullData.name}</b><br>"
                    "Ex-Date: %{x|%b %Y}<br>"
                    "Distribution: <b>USD %{y:.4f}</b><extra></extra>"
                ),
                legendgroup=fund,
            ),
            row=1, col=1,
        )

        # Trailing 12-month rolling total
        fd_monthly = (
            fd.set_index("Ex_Date")
            .resample("ME")["Distribution"]
            .sum()
            .reset_index()
        )
        fd_monthly["Trail12m"] = fd_monthly["Distribution"].rolling(12, min_periods=6).sum()
        fig.add_trace(
            go.Scatter(
                x=fd_monthly["Ex_Date"],
                y=fd_monthly["Trail12m"],
                name=f"{fund} · trailing 12m",
                line=dict(color=color, width=1.6, dash="dot"),
                opacity=0.65,
                hovertemplate=(
                    "Trailing 12m: <b>USD %{y:.4f}</b><extra></extra>"
                ),
                legendgroup=fund,
                showlegend=False,
            ),
            row=1, col=1,
        )

    # ── Panel 2: annualised yield line ────────────────────────────────────────
    for fund in funds:
        fd = combined_div[
            (combined_div["Fund_Name"] == fund) & combined_div["Ann_Yield"].notna()
        ].sort_values("Ex_Date")
        color = fund_color_map.get(fund, "#00d4aa")

        fig.add_trace(
            go.Scatter(
                x=fd["Ex_Date"],
                y=fd["Ann_Yield"],
                name=f"{fund} · yield",
                mode="lines+markers",
                line=dict(color=color, width=2),
                marker=dict(size=4, color=color),
                fill="tozeroy",
                fillcolor=color.replace(")", ",0.08)").replace("rgb", "rgba")
                    if color.startswith("rgb") else color + "14",
                hovertemplate=(
                    "<b>%{fullData.name}</b><br>"
                    "Ex-Date: %{x|%b %Y}<br>"
                    "NAV: USD %{customdata[0]:.4f}<br>"
                    "Distribution: USD %{customdata[1]:.4f}<br>"
                    "Ann. Yield: <b>%{y:.2%}</b><extra></extra>"
                ),
                customdata=np.stack(
                    [fd["NAV_on_ExDate"].values, fd["Distribution"].values], axis=1
                ),
                legendgroup=fund,
                showlegend=False,
            ),
            row=2, col=1,
        )

        # Horizontal avg yield reference line
        avg_yield = fd["Ann_Yield"].mean()
        fig.add_hline(
            y=avg_yield,
            line_dash="dash",
            line_color=color,
            line_width=1,
            opacity=0.4,
            annotation_text=f"avg {avg_yield:.1%}",
            annotation_font=dict(size=9, color=color, family="DM Mono"),
            annotation_position="right",
            row=2, col=1,
        )

    # ── Panel 3: annual total bars ────────────────────────────────────────────
    for fund in funds:
        fd = combined_div[combined_div["Fund_Name"] == fund]
        annual = fd.groupby("Year")["Distribution"].sum().reset_index()
        full_yrs = fd.groupby("Year")["Distribution"].count()
        full_yrs = full_yrs[full_yrs >= 10].index
        annual_full = annual[annual["Year"].isin(full_yrs)]
        color = fund_color_map.get(fund, "#00d4aa")

        fig.add_trace(
            go.Bar(
                x=annual_full["Year"].astype(str),
                y=annual_full["Distribution"],
                name=fund,
                marker_color=color,
                opacity=0.85,
                text=[f"${v:.3f}" for v in annual_full["Distribution"]],
                textposition="outside",
                textfont=dict(family="DM Mono", size=10, color="#dce6f0"),
                hovertemplate=(
                    "<b>%{fullData.name}</b><br>"
                    "Year: %{x}<br>"
                    "Annual total: <b>USD %{y:.4f}</b><extra></extra>"
                ),
                legendgroup=fund,
                showlegend=False,
            ),
            row=3, col=1,
        )

    # ── Layout ────────────────────────────────────────────────────────────────
    fig.update_layout(
        template="plotly_dark",
        paper_bgcolor="rgba(11,15,20,0)",
        plot_bgcolor="rgba(18,24,32,0.7)",
        font=dict(family="Instrument Sans", color="#dce6f0"),
        barmode="group",
        bargap=0.12,
        legend=dict(
            bgcolor="rgba(18,24,32,0.85)",
            bordercolor="#1e2a36",
            borderwidth=1,
            font=dict(size=12),
        ),
        hovermode="x unified",
        margin=dict(l=10, r=80, t=40, b=10),
        height=860,
    )
    for row in [1, 3]:
        fig.update_xaxes(gridcolor="#1e2a36", tickfont=dict(family="DM Mono", size=11), row=row, col=1)
        fig.update_yaxes(
            gridcolor="#1e2a36",
            tickfont=dict(family="DM Mono", size=11),
            tickprefix="$",
            tickformat=".3f",
            row=row, col=1,
        )
    # Yield axis: percent format
    fig.update_xaxes(gridcolor="#1e2a36", tickfont=dict(family="DM Mono", size=11), row=2, col=1)
    fig.update_yaxes(
        gridcolor="#1e2a36",
        tickfont=dict(family="DM Mono", size=11),
        tickformat=".1%",
        title=dict(text="Ann. Yield", font=dict(size=11)),
        row=2, col=1,
    )
    return fig

# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown(
        "<h2 style='font-family:DM Serif Display,serif;font-size:1.4rem;"
        "margin-bottom:0.3rem;'>Fund Analyzer</h2>"
        "<p style='font-size:0.78rem;color:#6b7f90;margin-top:0;'>Upload CSV or Excel files to begin</p>",
        unsafe_allow_html=True,
    )
    st.markdown("---")

    st.markdown("**Upload Fund files**")
    uploaded_files = st.file_uploader(
        "Upload Fund files",
        type=["csv", "xls", "xlsx", "xlsm"],
        accept_multiple_files=True,
        help="CSV or Excel (incl. BlackRock SpreadsheetML .xls). Column headers auto-detected.",
        label_visibility="collapsed",
    )

    st.markdown("---")
    st.markdown("**Display Options**")
    show_black_swan = st.checkbox(
        "Show Black Swan Events",
        value=False,
        help="Overlay shaded regions for major market crises.",
    )
    show_benchmark = st.checkbox(
        "Show Benchmark (MSCI ACWI Min Vol)",
        value=True,
        help="Overlay MSCI All Country World Minimum Volatility Index on the calendar chart.",
    )
    show_raw = st.checkbox("Show Raw Data Table", value=False)

    st.markdown("---")
    st.markdown(
        "<p style='font-size:0.72rem;color:#6b7f90;line-height:1.6;'>"
        "Columns auto-detected via fuzzy matching.<br>"
        "Normalized price sets first observation = 1.0.<br>"
        "Volatility = σ(returns) × √freq (industry standard)."
        "</p>",
        unsafe_allow_html=True,
    )


# ─────────────────────────────────────────────────────────────────────────────
# MAIN HEADER
# ─────────────────────────────────────────────────────────────────────────────
st.markdown(
    "<div class='app-header'>"
    "<h1>Fund Performance Analyzer</h1>"
    "<p>Multi-fund comparison · Black Swan overlays · Rolling analytics · Calendar returns</p>"
    "</div>",
    unsafe_allow_html=True,
)

# ─────────────────────────────────────────────────────────────────────────────
# MAIN LOGIC
# ─────────────────────────────────────────────────────────────────────────────
if not uploaded_files:
    st.markdown(
        "<div class='empty-state'>"
        "<div class='icon'>📂</div>"
        "<h3>Upload fund CSV files using the sidebar to get started</h3>"
        "<p style='color:#6b7f90;font-size:0.9rem;'>Supports any CSV with date + price/NAV columns.<br>"
        "Supports CSV and Excel (incl. BlackRock SpreadsheetML .xls).<br>Column headers are automatically detected — no manual mapping needed.</p>"
        "</div>",
        unsafe_allow_html=True,
    )
    st.stop()

# ─── Load & normalise all files ───────────────────────────────────────────────
all_dfs = []
all_divs = []
load_messages = []

with st.spinner("Parsing and normalizing uploaded files…"):
    for f in uploaded_files:
        f.seek(0)
        df_fund, msg, df_div_file = load_and_normalize(f)
        load_messages.append(msg)
        if df_fund is not None:
            all_dfs.append(df_fund)
            if not df_div_file.empty:
                fund_name_for_div = df_fund["Fund_Name"].iloc[0]
                df_div_file["Fund_Name"] = fund_name_for_div
                all_divs.append(df_div_file)

# Show ingestion report
with st.expander("Ingestion Report", expanded=len(all_dfs) == 0):
    for msg in load_messages:
        st.markdown(msg)

if not all_dfs:
    st.error("No valid fund data could be loaded. Check the ingestion report above.")
    st.stop()

combined = pd.concat(all_dfs, ignore_index=True)
funds = combined["Fund_Name"].unique().tolist()
fund_color_map = {f: FUND_COLORS[i % len(FUND_COLORS)] for i, f in enumerate(funds)}

# ─── Summary metrics row ──────────────────────────────────────────────────────
col_metrics = st.columns(min(len(funds), 4))
for i, fund in enumerate(funds[:4]):
    fd = combined[combined["Fund_Name"] == fund].sort_values("Date")
    total_ret = (fd["Price"].iloc[-1] / fd["Price"].iloc[0]) - 1
    with col_metrics[i]:
        color = "#00d4aa" if total_ret >= 0 else "#e05c5c"
        sign = "+" if total_ret >= 0 else ""
        st.markdown(
            f"<div class='stat-card'>"
            f"<div class='label'>{fund}</div>"
            f"<div class='value' style='color:{color};font-size:1.3rem;'>"
            f"{sign}{total_ret:.2%}</div>"
            f"<div style='font-size:0.72rem;color:#6b7f90;margin-top:0.2rem;'>"
            f"Total return · {len(fd):,} obs</div>"
            f"</div>",
            unsafe_allow_html=True,
        )

st.markdown("<div style='height:0.5rem'></div>", unsafe_allow_html=True)

# ─── Compute stats for all funds ─────────────────────────────────────────────
with st.spinner("Computing financial analytics…"):
    stats_map = {}
    for fund in funds:
        fd = combined[combined["Fund_Name"] == fund].copy()
        stats_map[fund] = compute_stats(fd)

# ─── Tabs ─────────────────────────────────────────────────────────────────────
tab_chart, tab_stats, tab_div = st.tabs(["📈  Visual Performance", "📊  Detailed Statistics", "💰  Dividend History"])

# ── TAB 1: Visual Performance ─────────────────────────────────────────────────
with tab_chart:
    with st.spinner("Rendering chart…"):
        fig = build_performance_chart(combined, show_black_swan, fund_color_map)
    st.plotly_chart(fig, use_container_width=True)

    if show_black_swan:
        st.markdown("##### Black Swan Events")
        cols_bs = st.columns(len(BLACK_SWAN_EVENTS))
        for i, ev in enumerate(BLACK_SWAN_EVENTS):
            with cols_bs[i]:
                st.markdown(
                    f"<div style='border-left:3px solid {ev['label_color']};"
                    f"padding-left:0.6rem;font-size:0.8rem;'>"
                    f"<b style='color:{ev['label_color']};'>{ev['name']}</b><br>"
                    f"<span style='color:#6b7f90;font-size:0.72rem;'>{ev['start']} → {ev['end']}</span>"
                    f"</div>",
                    unsafe_allow_html=True,
                )

    st.markdown("##### Calendar Year Returns")
    cal_fig = build_calendar_chart(stats_map, fund_color_map, show_benchmark=show_benchmark)
    if cal_fig:
        st.plotly_chart(cal_fig, use_container_width=True)
    else:
        st.info("Not enough data to compute calendar year returns.")

    if show_raw:
        st.markdown("##### Raw Normalized Data")
        st.dataframe(
            combined.sort_values(["Fund_Name", "Date"]),
            use_container_width=True,
            hide_index=True,
        )

# ── TAB 2: Detailed Statistics ────────────────────────────────────────────────
with tab_stats:

    # ── Section 1: Calendar Year Performance (portal style) ───────────────────
    st.markdown("##### Performance — Calendar Year")
    st.caption("Annual return for each full calendar year. Green = positive, red = negative.")
    render_calendar_table(stats_map, fund_color_map)

    st.markdown("<div style='height:1.5rem'></div>", unsafe_allow_html=True)

    # ── Section 2: Rolling-Window table ───────────────────────────────────────
    st.markdown("##### Rolling-Window Cumulative Return & Volatility")
    st.caption(
        "Windows calculated back from the most recent date in the dataset. "
        "Volatility = σ(periodic returns) × √annualisation_factor (252 daily · 52 weekly · 12 monthly)."
    )
    render_rolling_table(stats_map, fund_color_map)

    st.markdown("---")
    st.markdown("##### Methodology Notes")
    st.markdown("""
- **Calendar Year Return**: `(year_end_price / year_start_price) − 1` for each full calendar year.
- **Cumulative Return**: `(last_price / first_price_in_window) − 1` over the rolling window.
- **Annualized Volatility**: `std(periodic_returns) × √freq` — industry-standard definition.
  Frequency auto-detected from median gap between observations.
- **Normalized Price** (chart): each fund's price divided by its first available price (base = 1.0).
- **Black Swan Regions**: translucent vertical bands via Plotly `add_vrect`, rendered below fund lines.
    """)

# ── TAB 3: Dividend History ───────────────────────────────────────────────────
with tab_div:
    if not all_divs:
        st.info(
            "No dividend data found in the uploaded files. "
            "Upload a BlackRock SpreadsheetML .xls file to see distribution history."
        )
    else:
        combined_div_all = pd.concat(all_divs, ignore_index=True)
        combined_div_all["Year"] = combined_div_all["Ex_Date"].dt.year

        # ── Summary stat cards ────────────────────────────────────────────────
        for fund in combined_div_all["Fund_Name"].unique():
            fd = combined_div_all[combined_div_all["Fund_Name"] == fund]
            color = fund_color_map.get(fund, "#00d4aa")
            latest = fd.sort_values("Ex_Date").iloc[-1]
            annual_latest = fd[fd["Year"] == fd["Year"].max()]["Distribution"].sum()
            trailing12 = fd.sort_values("Ex_Date").tail(12)["Distribution"].sum()
            # Latest yield = latest distribution / latest NAV * 12
            nav_for_fund = combined[combined["Fund_Name"] == fund].sort_values("Date")
            nav_on_latest_div = np.nan
            if not nav_for_fund.empty:
                prior = nav_for_fund[nav_for_fund["Date"] <= latest["Ex_Date"]]
                if not prior.empty:
                    nav_on_latest_div = float(prior.iloc[-1]["Price"])
            latest_yield = (latest["Distribution"] / nav_on_latest_div * 12) if not np.isnan(nav_on_latest_div) else None

            col1, col2, col3, col4, col5 = st.columns(5)
            with col1:
                st.markdown(
                    f"<div class='stat-card'><div class='label'>Latest Payment</div>"
                    f"<div class='value' style='color:{color};'>USD {latest['Distribution']:.4f}</div>"
                    f"<div style='font-size:0.72rem;color:#6b7f90;'>{latest['Ex_Date'].strftime('%d %b %Y')}</div></div>",
                    unsafe_allow_html=True,
                )
            with col2:
                yield_str = f"{latest_yield:.2%}" if latest_yield is not None else "—"
                nav_str = f"NAV USD {nav_on_latest_div:.4f}" if not np.isnan(nav_on_latest_div) else ""
                st.markdown(
                    f"<div class='stat-card'><div class='label'>Latest Ann. Yield</div>"
                    f"<div class='value' style='color:{color};'>{yield_str}</div>"
                    f"<div style='font-size:0.72rem;color:#6b7f90;'>{nav_str}</div></div>",
                    unsafe_allow_html=True,
                )
            with col3:
                st.markdown(
                    f"<div class='stat-card'><div class='label'>Trailing 12-Month Total</div>"
                    f"<div class='value' style='color:{color};'>USD {trailing12:.4f}</div>"
                    f"<div style='font-size:0.72rem;color:#6b7f90;'>per unit</div></div>",
                    unsafe_allow_html=True,
                )
            with col4:
                count = len(fd)
                st.markdown(
                    f"<div class='stat-card'><div class='label'>Total Payments</div>"
                    f"<div class='value' style='color:{color};'>{count}</div>"
                    f"<div style='font-size:0.72rem;color:#6b7f90;'>since inception</div></div>",
                    unsafe_allow_html=True,
                )
            with col5:
                avg = fd["Distribution"].mean()
                st.markdown(
                    f"<div class='stat-card'><div class='label'>Avg per Payment</div>"
                    f"<div class='value' style='color:{color};'>USD {avg:.4f}</div>"
                    f"<div style='font-size:0.72rem;color:#6b7f90;'>all-time mean</div></div>",
                    unsafe_allow_html=True,
                )

        st.markdown("<div style='height:0.5rem'></div>", unsafe_allow_html=True)

        # ── Dividend chart ────────────────────────────────────────────────────
        div_fig = build_dividend_chart(all_divs, fund_color_map, combined)
        if div_fig:
            st.plotly_chart(div_fig, use_container_width=True)

        # ── Raw dividend table ────────────────────────────────────────────────
        with st.expander("Full Dividend History Table"):
            display_div = combined_div_all.copy()
            # Merge NAV on ex-date for yield calculation
            nav_lookup = combined[["Date","Fund_Name","Price"]].rename(columns={"Date":"Ex_Date"})
            display_div = pd.merge_asof(
                display_div.sort_values("Ex_Date"),
                nav_lookup.sort_values("Ex_Date"),
                on="Ex_Date", by="Fund_Name", direction="backward"
            )
            display_div["Ann_Yield"] = display_div["Distribution"] / display_div["Price"] * 12
            display_div["Ex_Date_str"] = display_div["Ex_Date"].dt.strftime("%d %b %Y")
            display_div["Distribution_str"] = display_div["Distribution"].map("USD {:.4f}".format)
            display_div["NAV_str"] = display_div["Price"].map("USD {:.4f}".format)
            display_div["Yield_str"] = display_div["Ann_Yield"].map("{:.2%}".format)
            st.dataframe(
                display_div[["Fund_Name", "Ex_Date_str", "Year", "Distribution_str", "NAV_str", "Yield_str"]]
                .rename(columns={
                    "Fund_Name": "Fund", "Ex_Date_str": "Ex-Date", "Year": "Year",
                    "Distribution_str": "Distribution", "NAV_str": "Unit Price", "Yield_str": "Ann. Yield"
                })
                .sort_values(["Fund", "Ex-Date"], ascending=[True, False]),
                use_container_width=True,
                hide_index=True,
            )
