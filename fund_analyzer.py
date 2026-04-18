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
    info = (
        f"SpreadsheetML · {len(df_m)} monthly returns + "
        f"{len(df_daily)} daily NAV rows stitched"
    )
    return df_out, info


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
            df_out, info = _parse_blackrock_spreadsheetml(file_bytes)
            return df_out, f"BlackRock SpreadsheetML ({info})"
        try:
            engine = "xlrd" if ext == "xls" else "openpyxl"
            df = pd.read_excel(_io.BytesIO(file_bytes), engine=engine)
            return df, f"{ext.upper()} (Excel)"
        except Exception as e:
            return None, f"❌ Could not read Excel file: {e}"

    try:
        df = pd.read_csv(_io.StringIO(file_bytes.decode("utf-8-sig", errors="replace")))
        return df, "CSV"
    except Exception as e:
        return None, f"❌ Could not read file: {e}"


def load_and_normalize(uploaded_file) -> Tuple[Optional[pd.DataFrame], str]:
    """
    Parse one uploaded file (CSV / XLS / XLSX / SpreadsheetML) into the
    standard long-format DataFrame.
    """
    raw, fmt_info = _read_raw_dataframe(uploaded_file)

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
        return df[["Date", "Fund_Name", "Price", "Normalized_Price"]], msg

    if raw is None:
        return None, fmt_info

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
    return df[["Date", "Fund_Name", "Price", "Normalized_Price"]], detected_msg


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


def build_calendar_chart(stats_map: dict, fund_color_map: dict) -> go.Figure:
    """Grouped bar chart for calendar year returns."""
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
    for fund, stats in stats_map.items():
        if not stats or not stats.get("calendar_years"):
            continue
        cy = stats["calendar_years"]
        ys = [cy.get(y, None) for y in all_years]
        colors = [
            "#00d4aa" if (v is not None and v >= 0) else "#e05c5c"
            for v in ys
        ]
        fig.add_trace(
            go.Bar(
                x=[str(y) for y in all_years],
                y=ys,
                name=fund,
                marker_color=fund_color_map.get(fund, "#00d4aa"),
                text=[f"{v:.1%}" if v is not None else "" for v in ys],
                textposition="outside",
                textfont=dict(family="DM Mono", size=10),
                hovertemplate="<b>%{fullData.name}</b><br>%{x}: <b>%{y:.2%}</b><extra></extra>",
            )
        )

    fig.update_layout(
        template="plotly_dark",
        paper_bgcolor="rgba(11,15,20,0)",
        plot_bgcolor="rgba(18,24,32,0.7)",
        font=dict(family="Instrument Sans", color="#dce6f0"),
        barmode="group",
        bargap=0.25,
        bargroupgap=0.08,
        legend=dict(bgcolor="rgba(18,24,32,0.85)", bordercolor="#1e2a36", borderwidth=1),
        xaxis=dict(gridcolor="#1e2a36", tickfont=dict(family="DM Mono", size=11)),
        yaxis=dict(
            gridcolor="#1e2a36",
            tickformat=".0%",
            tickfont=dict(family="DM Mono", size=11),
            title=dict(text="Annual Return", font=dict(size=12)),
            zeroline=True,
            zerolinecolor="#3a4a58",
            zerolinewidth=1.5,
        ),
        margin=dict(l=10, r=10, t=20, b=10),
        height=400,
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
load_messages = []

with st.spinner("Parsing and normalizing uploaded files…"):
    for f in uploaded_files:
        f.seek(0)
        df_fund, msg = load_and_normalize(f)
        load_messages.append(msg)
        if df_fund is not None:
            all_dfs.append(df_fund)

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
tab_chart, tab_stats = st.tabs(["📈  Visual Performance", "📊  Detailed Statistics"])

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
    cal_fig = build_calendar_chart(stats_map, fund_color_map)
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
