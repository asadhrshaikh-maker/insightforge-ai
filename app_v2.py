# =========================================================
# Business Copilot v5.0 — Premium AI-Powered BI Dashboard
# ---------------------------------------------------------
# Install:  pip install "streamlit>=1.40" pandas numpy plotly openpyxl xlrd
# Run:      python -m streamlit run business_copilot.py
# =========================================================
from __future__ import annotations

import csv
import hashlib
import html
import re
from datetime import datetime
from io import BytesIO

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# ---------------------------------------------------------
# SETTINGS
# ---------------------------------------------------------
MAX_CATEGORIES = 50
ID_UNIQUE_RATIO = 0.95
DATE_MIN_SUCCESS = 0.90
MAX_PLOT_ROWS = 50_000
MAX_COLOR_GROUPS = 12
MAX_FILTER_OPTIONS = 500
MAX_CORR_COLS = 40
MAX_EXCEL_ROWS = 200_000
STRONG_R, MODERATE_R = 0.70, 0.50

ROW_COUNT = "(row count)"
AGGS = {"Sum": "sum", "Average": "mean", "Median": "median", "Count": "count", "Min": "min", "Max": "max"}
FREQS = {"Day": "D", "Week": "W", "Month": "MS", "Quarter": "QS"}
METRIC_HINTS = ("revenue", "sales", "amount", "total", "profit", "income", "cost", "value", "price")

ACCENT = {
    "blue":   "#3B82F6",
    "purple": "#8B5CF6",
    "green":  "#10B981",
    "orange": "#F59E0B",
    "red":    "#EF4444",
    "cyan":   "#06B6D4",
    "pink":   "#EC4899",
}

_DATE_RE = r"^\s*(?:\d{4}[-/.]\d{1,2}[-/.]\d{1,2}|\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4})"
_ID_NAME_RE = re.compile(r"(?:^|[\W_])(?:id|uuid|guid|idx|index)$|[a-z]ID$|^(?:id|uuid)", re.I)

st.set_page_config(
    page_title="Business Copilot AI",
    page_icon="🚀",
    layout="wide",
    initial_sidebar_state="expanded",
)

# =========================================================
# PREMIUM DARK GLASSMORPHISM CSS
# =========================================================
st.markdown("""
<style>
/* ---------- Premium gradient background ---------- */
.stApp {
    background: linear-gradient(180deg, #061326 0%, #0B1E36 50%, #061326 100%);
    color: #E5E7EB;
}

/* ---------- Global typography ---------- */
html, body, [class*="css"] {
    color: #E5E7EB !important;
    font-family: 'Inter', 'Segoe UI', system-ui, -apple-system, sans-serif;
}
h1, h2, h3, h4, h5, h6 {
    color: #FFFFFF !important;
    letter-spacing: -0.02em;
    font-weight: 700;
}
p, label, span, div { color: #CBD5E1 !important; }

/* ---------- Sidebar ---------- */
section[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #0B1E36 0%, #061326 100%);
    border-right: 1px solid rgba(59, 130, 246, 0.15);
}
section[data-testid="stSidebar"] * { color: #E5E7EB !important; }

/* ---------- Tabs ---------- */
.stTabs [data-baseweb="tab-list"] {
    gap: 8px;
    background: rgba(11, 30, 54, 0.6);
    padding: 8px;
    border-radius: 14px;
    border: 1px solid rgba(59, 130, 246, 0.15);
}
.stTabs [data-baseweb="tab"] {
    color: #94A3B8 !important;
    font-weight: 600;
    border-radius: 10px;
    padding: 10px 20px;
    transition: 0.25s;
}
.stTabs [data-baseweb="tab"]:hover {
    color: #FFFFFF !important;
    background: rgba(59, 130, 246, 0.12);
}
.stTabs [aria-selected="true"] {
    color: #FFFFFF !important;
    background: linear-gradient(135deg, #3B82F6, #8B5CF6) !important;
    box-shadow: 0 4px 20px rgba(59, 130, 246, 0.4);
}
.stTabs [data-baseweb="tab-highlight"] { display: none; }
.stTabs [data-baseweb="tab-border"] { display: none; }

/* ---------- Headings ---------- */
.app-header h1 {
    font-size: 44px;
    margin: 0;
    background: linear-gradient(90deg, #FFFFFF, #60A5FA);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    background-clip: text;
    font-weight: 800;
    letter-spacing: -0.03em;
}
.app-header p {
    color: #94A3B8 !important;
    font-size: 15px;
    margin-top: 6px;
}

/* ---------- KPI Cards (Glassmorphism) ---------- */
.kpi-card {
    background: linear-gradient(135deg, rgba(17,24,39,0.85), rgba(30,41,59,0.65));
    border: 1px solid rgba(59,130,246,0.20);
    border-radius: 18px;
    padding: 22px 20px;
    box-shadow: 0 10px 30px rgba(0,0,0,0.35), inset 0 1px 0 rgba(255,255,255,0.04);
    transition: all 0.3s ease;
    backdrop-filter: blur(12px);
    -webkit-backdrop-filter: blur(12px);
    position: relative;
    overflow: hidden;
    height: 100%;
}
.kpi-card:hover {
    transform: translateY(-4px);
    border-color: rgba(59,130,246,0.55);
    box-shadow: 0 18px 40px rgba(59,130,246,0.28);
}
.kpi-card::before {
    content: "";
    position: absolute;
    top: 0; left: 0; right: 0;
    height: 3px;
    background: linear-gradient(90deg, var(--glow, #3B82F6), transparent);
    opacity: 0.9;
}
.kpi-icon {
    font-size: 20px;
    margin-bottom: 8px;
    opacity: 0.85;
}
.kpi-title {
    color: #94A3B8 !important;
    font-size: 12px;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 1px;
    margin-bottom: 10px;
}
.kpi-value {
    color: #FFFFFF !important;
    font-size: 30px;
    font-weight: 800;
    line-height: 1.1;
    letter-spacing: -0.02em;
}
.kpi-subtitle {
    color: #64748B !important;
    font-size: 12px;
    margin-top: 10px;
    line-height: 1.4;
}

/* ---------- AI Summary Card ---------- */
.ai-summary-card {
    background: linear-gradient(135deg, rgba(30,41,59,0.85), rgba(17,24,39,0.7));
    border: 1px solid rgba(139,92,246,0.3);
    border-left: 4px solid #8B5CF6;
    border-radius: 18px;
    padding: 28px 32px;
    margin-bottom: 28px;
    box-shadow: 0 15px 40px rgba(139,92,246,0.15);
    backdrop-filter: blur(14px);
    -webkit-backdrop-filter: blur(14px);
}
.ai-summary-card h3 {
    margin-top: 0;
    color: #A78BFA !important;
    font-size: 18px;
    display: flex;
    align-items: center;
    gap: 10px;
    font-weight: 700;
}
.ai-summary-card p, .ai-summary-card li {
    color: #E5E7EB !important;
    font-size: 15px;
    line-height: 1.75;
    margin: 8px 0;
}
.ai-summary-card strong { color: #FFFFFF !important; }

/* ---------- Insight Cards ---------- */
.insight-card {
    background: linear-gradient(135deg, rgba(17,24,39,0.85), rgba(30,41,59,0.6));
    border: 1px solid rgba(148,163,184,0.15);
    border-radius: 16px;
    padding: 22px;
    height: 100%;
    box-shadow: 0 8px 24px rgba(0,0,0,0.28);
    backdrop-filter: blur(10px);
    transition: 0.25s;
}
.insight-card:hover {
    transform: translateY(-3px);
    border-color: rgba(59,130,246,0.35);
}
.insight-card h4 {
    font-size: 15px;
    margin: 0 0 12px 0;
    display: flex;
    align-items: center;
    gap: 8px;
    font-weight: 700;
}
.insight-card ul {
    margin: 0;
    padding-left: 18px;
}
.insight-card li {
    color: #CBD5E1 !important;
    font-size: 13.5px;
    line-height: 1.65;
    margin-bottom: 8px;
}
.insight-happened h4 { color: #10B981 !important; }
.insight-risks h4 { color: #F59E0B !important; }
.insight-opps h4 { color: #3B82F6 !important; }
.insight-recs h4 { color: #8B5CF6 !important; }

/* ---------- Section headers ---------- */
.section-title {
    font-size: 22px;
    font-weight: 700;
    color: #FFFFFF !important;
    margin: 32px 0 16px 0;
    padding-bottom: 8px;
    border-bottom: 1px solid rgba(59,130,246,0.15);
    letter-spacing: -0.01em;
}

/* ---------- Badges ---------- */
.badge {
    display: inline-block;
    padding: 3px 10px;
    border-radius: 20px;
    font-size: 11px;
    font-weight: 600;
    letter-spacing: 0.4px;
    text-transform: uppercase;
}
.badge-green { background: rgba(16,185,129,0.15); color: #34D399; border: 1px solid rgba(16,185,129,0.3); }
.badge-red   { background: rgba(239,68,68,0.15);  color: #F87171; border: 1px solid rgba(239,68,68,0.3); }
.badge-orange{ background: rgba(245,158,11,0.15); color: #FBBF24; border: 1px solid rgba(245,158,11,0.3); }
.badge-blue  { background: rgba(59,130,246,0.15); color: #60A5FA; border: 1px solid rgba(59,130,246,0.3); }

/* ---------- Buttons ---------- */
.stButton > button, .stDownloadButton > button {
    background: linear-gradient(135deg, #3B82F6, #8B5CF6);
    color: white !important;
    border: none;
    border-radius: 12px;
    padding: 10px 22px;
    font-weight: 600;
    box-shadow: 0 6px 20px rgba(59,130,246,0.3);
    transition: 0.25s;
}
.stButton > button:hover, .stDownloadButton > button:hover {
    transform: translateY(-2px);
    box-shadow: 0 10px 28px rgba(139,92,246,0.45);
}

/* ---------- Info / success / warning overrides ---------- */
.stAlert {
    background: rgba(30,41,59,0.6) !important;
    border-radius: 12px !important;
    border: 1px solid rgba(59,130,246,0.2) !important;
    color: #E5E7EB !important;
}

/* ---------- Fix Streamlit plotly container bg ---------- */
.js-plotly-plot .plotly .modebar { background: transparent !important; }

/* ---------- Dataframe ---------- */
[data-testid="stDataFrame"] {
    border-radius: 14px;
    overflow: hidden;
    border: 1px solid rgba(59,130,246,0.15);
}
</style>
""", unsafe_allow_html=True)


# =========================================================
# GENERIC HELPERS
# =========================================================
def fmt_num(x) -> str:
    if x is None or not np.isfinite(x):
        return "–"
    a = abs(x)
    if a >= 1e9:
        return f"{x / 1e9:,.2f}B"
    if a >= 1e6:
        return f"{x / 1e6:,.2f}M"
    if a >= 1e4 or float(x).is_integer():
        return f"{x:,.0f}"
    return f"{x:,.2f}"


def _premium_fig(fig, height=380):
    """Apply consistent premium dark chart styling."""
    fig.update_layout(
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(11,30,54,0.35)",
        font=dict(color="#E5E7EB", size=13, family="Inter, Segoe UI, sans-serif"),
        margin=dict(l=20, r=20, t=50, b=20),
        height=height,
        legend=dict(bgcolor="rgba(0,0,0,0)", font=dict(color="#CBD5E1")),
        title=dict(font=dict(size=15, color="#FFFFFF")),
        xaxis=dict(gridcolor="rgba(148,163,184,0.10)", zerolinecolor="rgba(148,163,184,0.15)"),
        yaxis=dict(gridcolor="rgba(148,163,184,0.10)", zerolinecolor="rgba(148,163,184,0.15)"),
    )
    return fig


def show_chart(fig, height=380):
    fig = _premium_fig(fig, height)
    try:
        st.plotly_chart(fig, width="stretch", config={"displaylogo": False})
    except TypeError:
        st.plotly_chart(fig, use_container_width=True, config={"displaylogo": False})


def show_df(data):
    try:
        st.dataframe(data, width="stretch", hide_index=True)
    except TypeError:
        st.dataframe(data, use_container_width=True, hide_index=True)


def _unique_names(cols) -> list:
    used, out = set(), []
    for c in cols:
        base = str(c).strip() or "Unnamed"
        name, i = base, 1
        while name in used:
            name = f"{base}_{i}"
            i += 1
        used.add(name)
        out.append(name)
    return out


# =========================================================
# LOADING & PROFILING
# =========================================================
@st.cache_data(show_spinner=False, max_entries=4)
def excel_sheets(name: str, data: bytes) -> list:
    try:
        with pd.ExcelFile(BytesIO(data)) as wb:
            return list(wb.sheet_names)
    except ImportError as err:
        raise ValueError(f"Reading '{name}' needs an extra package: {err}. Try: pip install openpyxl xlrd") from err
    except Exception as err:
        raise ValueError(f"Could not open '{name}' as an Excel workbook: {err}") from err


def _read_csv(name: str, data: bytes) -> pd.DataFrame:
    last_err = None
    for enc in ("utf-8-sig", "cp1252", "latin1"):
        try:
            sample = data[:65536].decode(enc)
        except UnicodeDecodeError as err:
            last_err = err
            continue
        try:
            sep = csv.Sniffer().sniff(sample, delimiters=",;\t|").delimiter
        except csv.Error:
            sep = ","
        try:
            return pd.read_csv(BytesIO(data), encoding=enc, sep=sep)
        except UnicodeDecodeError as err:
            last_err = err
    raise ValueError(f"Could not decode '{name}': {last_err}")


def _parse_dates(s: pd.Series, dayfirst: bool):
    nonnull = s.dropna().astype(str)
    if nonnull.empty or nonnull.head(1000).str.match(_DATE_RE).mean() < DATE_MIN_SUCCESS:
        return None, False
    parts = nonnull.str.extract(r"^\s*(\d{1,2})[-/.](\d{1,2})[-/.]\d{2,4}")
    first, second = pd.to_numeric(parts[0], errors="coerce"), pd.to_numeric(parts[1], errors="coerce")
    if ((first > 12).any() and not dayfirst) or ((second > 12).any() and dayfirst):
        return None, True
    parsed = pd.to_datetime(s, errors="coerce", format="mixed", dayfirst=dayfirst, utc=True).dt.tz_convert(None)
    ok = parsed.notna().sum() / len(nonnull) >= DATE_MIN_SUCCESS
    return (parsed if ok else None), True


def _looks_like_id(name, s, nn, nun) -> bool:
    if not pd.api.types.is_integer_dtype(s) or nn < 20 or nun != nn:
        return False
    return bool(_ID_NAME_RE.search(str(name))) or bool(s.dropna().is_monotonic_increasing)


def classify(df: pd.DataFrame) -> dict:
    roles = {"numeric": [], "datetime": [], "categorical": [], "text": [], "id": []}
    for c in df.columns:
        s = df[c]
        nn, nun = int(s.notna().sum()), int(s.nunique(dropna=True))
        if pd.api.types.is_bool_dtype(s):
            roles["categorical"].append(c)
        elif pd.api.types.is_datetime64_any_dtype(s):
            roles["datetime"].append(c)
        elif pd.api.types.is_numeric_dtype(s):
            roles["id" if _looks_like_id(c, s, nn, nun) else "numeric"].append(c)
        elif nun <= MAX_CATEGORIES:
            roles["categorical"].append(c)
        elif nn and nun / nn >= ID_UNIQUE_RATIO:
            roles["id"].append(c)
        else:
            roles["text"].append(c)
    return roles


def profile_table(df: pd.DataFrame, roles: dict) -> pd.DataFrame:
    role_of = {c: r for r, cols in roles.items() for c in cols}
    rows = []
    for c in df.columns:
        s = df[c]
        rows.append({
            "Column": c, "Detected as": role_of.get(c, "?"), "Dtype": str(s.dtype),
            "Missing %": round(float(s.isna().mean()) * 100, 1), "Unique": int(s.nunique(dropna=True)),
            "Example": str(s.dropna().iloc[0])[:40] if s.notna().any() else "",
        })
    return pd.DataFrame(rows)


@st.cache_data(show_spinner="Reading and profiling your data…", max_entries=4)
def load_dataset(name: str, data: bytes, sheet, dayfirst: bool):
    lower = name.lower()
    try:
        if lower.endswith((".csv", ".txt")):
            raw = _read_csv(name, data)
        elif lower.endswith((".xlsx", ".xls")):
            raw = pd.read_excel(BytesIO(data), sheet_name=sheet if sheet is not None else 0)
        else:
            raise ValueError(f"Unsupported file type: {name}")
    except pd.errors.EmptyDataError as err:
        raise ValueError(f"'{name}' is empty.") from err
    except ValueError:
        raise
    except ImportError as err:
        raise ValueError(f"Reading '{name}' needs an extra package: {err}. Try: pip install openpyxl xlrd") from err
    except Exception as err:
        raise ValueError(f"Could not read '{name}': {err}") from err

    notes = []
    df = raw.copy()
    renamed = _unique_names(df.columns)
    if renamed != [str(c) for c in df.columns]:
        notes.append("Column names were cleaned (whitespace trimmed, duplicates numbered).")
    df.columns = renamed

    before = df.shape
    df = df.dropna(how="all").dropna(axis=1, how="all").reset_index(drop=True)
    if df.shape != before:
        notes.append(f"Removed {before[0] - df.shape[0]} empty row(s) and {before[1] - df.shape[1]} empty column(s).")
    if df.empty:
        raise ValueError(f"'{name}' has no data rows.")

    for c in df.columns:
        s = df[c]
        if pd.api.types.is_datetime64_any_dtype(s):
            if getattr(s.dt, "tz", None) is not None:
                df[c] = s.dt.tz_convert(None)
        elif not pd.api.types.is_numeric_dtype(s) and not pd.api.types.is_bool_dtype(s):
            parsed, looked_like_dates = _parse_dates(s, dayfirst)
            if parsed is not None:
                df[c] = parsed
                notes.append(f"Column '{c}' was read as dates ({'DD/MM' if dayfirst else 'MM/DD'} order).")
            elif looked_like_dates:
                notes.append(f"Column '{c}' looks like dates but its day/month order does not match the "
                             "'Dates are DD/MM/YYYY' setting, so it was kept as text.")

    roles = classify(df)
    return df, roles, notes, profile_table(df, roles)


# =========================================================
# SMART DATASET DETECTION
# =========================================================
def detect_dataset_type(df: pd.DataFrame, roles: dict) -> tuple[str, str]:
    cols_lower = [c.lower() for c in df.columns]
    joined = " ".join(cols_lower)

    sales_kw = ["revenue", "sales", "order", "product", "customer", "discount", "quantity", "price"]
    finance_kw = ["expense", "budget", "cost", "invoice", "account", "ledger", "debit", "credit", "profit", "margin"]
    hr_kw = ["employee", "salary", "department", "hire", "attrition", "job", "role", "leave", "performance", "headcount"]
    survey_kw = ["response", "question", "answer", "score", "feedback", "rating", "survey", "nps"]
    ops_kw = ["ticket", "sla", "delivery", "shipment", "inventory", "supplier", "warehouse", "cycle", "backlog"]

    def _score(kws):
        return sum(1 for k in kws if any(k in c for c in cols_lower))

    scores = {
        "Sales":     _score(sales_kw),
        "Finance":   _score(finance_kw),
        "HR":        _score(hr_kw),
        "Survey":    _score(survey_kw),
        "Operations":_score(ops_kw),
    }
    best = max(scores, key=scores.get)
    if scores[best] < 2:
        best = "General"

    descriptions = {
        "Sales": "Looks like a sales dataset. Focus on revenue trends, product mix, regional performance, and customer concentration.",
        "Finance": "Looks like a finance dataset. Focus on expense control, margin analysis, budget vs actuals, and cash flow.",
        "HR": "Looks like an HR dataset. Focus on headcount, attrition, department productivity, and compensation distribution.",
        "Survey": "Looks like a survey dataset. Focus on NPS, satisfaction drivers, and response segmentation.",
        "Operations": "Looks like an operations dataset. Focus on SLA compliance, delivery times, inventory, and bottleneck analysis.",
        "General": "General business dataset. Focus on the primary metric, segment distribution, and trend movement.",
    }
    return best, descriptions[best]


# =========================================================
# AI INSIGHT ENGINE
# =========================================================
def compute_ai_metrics(fdf: pd.DataFrame, metric: str, segment: str, date_col: str) -> dict:
    """Compute all the numeric figures the AI engine needs."""
    m = {}
    m["total_rows"] = len(fdf)
    m["metric"] = metric
    m["segment"] = segment
    m["date_col"] = date_col

    if metric and metric in fdf.columns:
        m["total_metric"] = float(fdf[metric].sum())
        m["avg_metric"] = float(fdf[metric].mean())
        m["median_metric"] = float(fdf[metric].median())
    else:
        m["total_metric"] = None
        m["avg_metric"] = None
        m["median_metric"] = None

    # Data quality
    cells = fdf.shape[0] * fdf.shape[1]
    m["completeness"] = (1 - fdf.isna().sum().sum() / cells) * 100 if cells else 0.0

    # Segment concentration
    m["top_segment"] = None
    m["top_segment_value"] = None
    m["top_segment_share"] = None
    m["worst_segment"] = None
    m["worst_segment_value"] = None
    if metric and segment and segment in fdf.columns and metric in fdf.columns:
        grp = fdf.groupby(segment)[metric].sum().sort_values(ascending=False)
        if not grp.empty and grp.sum() > 0:
            m["top_segment"] = str(grp.index[0])
            m["top_segment_value"] = float(grp.iloc[0])
            m["top_segment_share"] = float(grp.iloc[0] / grp.sum() * 100)
            m["worst_segment"] = str(grp.index[-1])
            m["worst_segment_value"] = float(grp.iloc[-1])

    # Growth over last 30 days vs previous 30 days
    m["growth_pct"] = None
    m["current_period"] = None
    m["previous_period"] = None
    if metric and date_col and metric in fdf.columns and date_col in fdf.columns:
        d = fdf.dropna(subset=[date_col, metric]).copy()
        if not d.empty:
            last_date = d[date_col].max()
            cut = last_date - pd.Timedelta(days=30)
            cur = float(d.loc[d[date_col] > cut, metric].sum())
            prev = float(d.loc[(d[date_col] <= cut) & (d[date_col] > cut - pd.Timedelta(days=30)), metric].sum())
            m["current_period"] = cur
            m["previous_period"] = prev
            if prev > 0:
                m["growth_pct"] = ((cur - prev) / prev) * 100

    return m


def generate_ai_summary_paragraph(metrics: dict, dataset_type: str) -> str:
    """Management-friendly executive summary paragraph."""
    lines = []
    metric = metrics["metric"]
    mname = metric.capitalize() if metric else "Total"

    if metrics["total_metric"] is not None:
        lines.append(
            f"**{mname}** reached **{fmt_num(metrics['total_metric'])}** across "
            f"**{metrics['total_rows']:,}** transactions."
        )
    else:
        lines.append(f"Dataset contains **{metrics['total_rows']:,}** records.")

    if metrics["top_segment"] and metrics["top_segment_share"] is not None:
        lines.append(
            f"**{metrics['top_segment']}** contributed **{metrics['top_segment_share']:.1f}%** "
            f"of {metric or 'total volume'} and remains the leading {metrics['segment']}."
        )

    if metrics["growth_pct"] is not None:
        direction = "grew" if metrics["growth_pct"] >= 0 else "declined"
        lines.append(
            f"Performance {direction} **{abs(metrics['growth_pct']):.1f}%** "
            f"over the last 30 days versus the previous period."
        )

    if metrics["top_segment_share"] is not None:
        if metrics["top_segment_share"] > 40:
            lines.append("Performance is **highly concentrated** in one segment — concentration risk should be reviewed.")
        elif metrics["top_segment_share"] > 25:
            lines.append("Revenue is **moderately concentrated**, with the leading segment still dominating.")

    lines.append(f"*Dataset classification: **{dataset_type}**.*")
    lines.append("Management should investigate declining segments and reduce concentration risk where possible.")

    return " ".join(lines)


def build_ai_insights(metrics: dict, dataset_type: str) -> dict:
    """Return structured insight lists: happened / risks / opportunities / actions."""
    happened, risks, opps, actions = [], [], [], []
    metric = metrics["metric"] or "records"

    # --- What happened ---
    if metrics["total_metric"] is not None:
        happened.append(f"Total {metric}: **{fmt_num(metrics['total_metric'])}** across {metrics['total_rows']:,} records.")
    if metrics["top_segment"]:
        happened.append(
            f"**{metrics['top_segment']}** is the top {metrics['segment']} "
            f"({metrics['top_segment_share']:.1f}% share)."
        )
    if metrics["worst_segment"]:
        happened.append(
            f"**{metrics['worst_segment']}** is the weakest {metrics['segment']} "
            f"(value {fmt_num(metrics['worst_segment_value'])})."
        )
    if metrics["growth_pct"] is not None:
        word = "grew" if metrics["growth_pct"] >= 0 else "declined"
        happened.append(f"Last 30-day {metric} {word} **{abs(metrics['growth_pct']):.1f}%** vs prior period.")

    # --- Risks ---
    if metrics["top_segment_share"] and metrics["top_segment_share"] > 40:
        risks.append(
            f"**High concentration risk**: {metrics['top_segment']} drives "
            f"{metrics['top_segment_share']:.1f}% of {metric}."
        )
    if metrics["completeness"] < 95:
        risks.append(f"**Data quality risk**: dataset is only {metrics['completeness']:.1f}% complete.")
    if metrics["growth_pct"] is not None and metrics["growth_pct"] < -5:
        risks.append(f"**Declining trend**: {metric} dropped {abs(metrics['growth_pct']):.1f}% in the last 30 days.")
    if not risks:
        risks.append("No major risks detected in the current view.")

    # --- Opportunities ---
    if metrics["worst_segment"]:
        opps.append(f"Upside potential in **{metrics['worst_segment']}** — currently the weakest segment.")
    if metrics["growth_pct"] is not None and metrics["growth_pct"] > 0:
        opps.append("Momentum is positive — capitalise on the recent growth by scaling top performers.")
    if metrics["top_segment_share"] and metrics["top_segment_share"] < 30:
        opps.append("Revenue is well-diversified across segments — a stable base for expansion.")
    opps.append(f"AI recommends focusing on **{dataset_type}** specific drivers for deeper analysis.")
    if not opps:
        opps.append("Investigate secondary segments for untapped potential.")

    # --- Recommended actions ---
    if metrics["growth_pct"] is not None and metrics["growth_pct"] < 0:
        actions.append("Run root-cause analysis on the last 30-day decline.")
    if metrics["top_segment_share"] and metrics["top_segment_share"] > 40:
        actions.append(f"Diversify away from **{metrics['top_segment']}** to reduce concentration risk.")
    if metrics["completeness"] < 95:
        actions.append("Clean missing values before the next reporting cycle.")
    actions.append("Share this dashboard with leadership for weekly review.")
    actions.append("Set up alerts for KPI movements beyond ±10%.")
    if not actions:
        actions.append("Maintain current strategy and monitor KPIs weekly.")

    return {"happened": happened, "risks": risks, "opportunities": opps, "actions": actions}


# =========================================================
# RENDER: HEADER
# =========================================================
def render_header(name: str, sheet, n_filtered: int, n_total: int, loaded_at: datetime):
    sheet_txt = f" · sheet {html.escape(str(sheet))}" if sheet else ""
    st.markdown(
        f'<div class="app-header">'
        f'<h1>🚀 Business Copilot AI</h1>'
        f'<p>{html.escape(name)}{sheet_txt} · '
        f'<strong>{n_filtered:,}</strong> of {n_total:,} records · '
        f'loaded {loaded_at:%d-%b-%Y %H:%M}</p>'
        f'</div>',
        unsafe_allow_html=True,
    )


# =========================================================
# RENDER: AI EXECUTIVE SUMMARY CARD
# =========================================================
def render_ai_summary(summary_md: str):
    st.markdown(
        f'<div class="ai-summary-card">'
        f'<h3>🤖 AI Executive Summary</h3>'
        f'<p>{summary_md}</p>'
        f'</div>',
        unsafe_allow_html=True,
    )


# =========================================================
# RENDER: KPI CARDS
# =========================================================
def render_kpi_cards(cards: list):
    if not cards:
        return
    cols = st.columns(len(cards))
    for col, k in zip(cols, cards):
        glow = k.get("color", ACCENT["blue"])
        icon = k.get("icon", "📊")
        col.markdown(
            f'<div class="kpi-card" style="--glow:{glow};">'
            f'<div class="kpi-icon">{icon}</div>'
            f'<div class="kpi-title">{html.escape(k["title"])}</div>'
            f'<div class="kpi-value">{html.escape(k["value"])}</div>'
            f'<div class="kpi-subtitle">{k["subtitle"]}</div>'
            f'</div>',
            unsafe_allow_html=True,
        )


# =========================================================
# RENDER: VISUALIZATIONS (auto charts, no dropdowns)
# =========================================================
def render_visualizations(fdf, metric, segment, date_col):
    # ---------- Row 1: Trend + Distribution ----------
    st.markdown('<div class="section-title">📈 Performance Overview</div>', unsafe_allow_html=True)
    r1c1, r1c2 = st.columns([3, 2])

    with r1c1:
        if metric and date_col:
            trend = (
                fdf.dropna(subset=[date_col])
                .groupby(pd.Grouper(key=date_col, freq="MS"))[metric]
                .sum().reset_index().sort_values(date_col)
            )
            if not trend.empty:
                fig = px.area(
                    trend, x=date_col, y=metric,
                    title=f"{metric} — Monthly Trend",
                    template="plotly_dark",
                )
                fig.update_traces(
                    line_color=ACCENT["blue"], line_width=3,
                    fillcolor="rgba(59,130,246,0.15)",
                )
                show_chart(fig, height=360)
            else:
                st.info("Not enough date data for trend chart.")
        elif metric:
            fig = px.histogram(fdf, x=metric, nbins=40, template="plotly_dark",
                               title=f"{metric} Distribution")
            fig.update_traces(marker_color=ACCENT["blue"])
            show_chart(fig, height=360)

    with r1c2:
        if metric and segment:
            seg_data = (
                fdf.groupby(segment)[metric].sum()
                .reset_index().sort_values(metric, ascending=False).head(10)
            )
            if not seg_data.empty and seg_data[metric].sum() > 0:
                fig = px.pie(
                    seg_data, names=segment, values=metric, hole=0.55,
                    title=f"{segment} Contribution",
                    template="plotly_dark",
                    color_discrete_sequence=[
                        ACCENT["blue"], ACCENT["purple"], ACCENT["green"],
                        ACCENT["orange"], ACCENT["cyan"], ACCENT["pink"],
                        "#64748B", "#94A3B8", "#CBD5E1", "#F472B6",
                    ],
                )
                fig.update_traces(textposition="inside", textinfo="percent")
                show_chart(fig, height=360)

    # ---------- Row 2: Top 10 Contributors ----------
    st.markdown('<div class="section-title">🏆 Top Contributors</div>', unsafe_allow_html=True)

    cat_candidates = []
    for c in [segment, "Product", "Category", "Sub-Category", "Region", "Location",
              "Department", "Job Role", "Country", "State", "City", "Channel"]:
        if c and c in fdf.columns and c not in cat_candidates:
            cat_candidates.append(c)
    cat_candidates = cat_candidates[:4]

    if metric and cat_candidates:
        row1 = cat_candidates[:2]
        row2 = cat_candidates[2:4]
        for row in [row1, row2]:
            if not row:
                continue
            cols = st.columns(len(row))
            for col, cat_col in zip(cols, row):
                with col:
                    data = (
                        fdf.groupby(cat_col)[metric].sum()
                        .reset_index().sort_values(metric, ascending=False).head(10)
                    )
                    if not data.empty:
                        fig = px.bar(
                            data, x=metric, y=cat_col, orientation="h",
                            title=f"Top 10 {cat_col}",
                            template="plotly_dark",
                            color=metric,
                            color_continuous_scale=[[0, "#1E3A8A"], [1, "#60A5FA"]],
                        )
                        fig.update_layout(yaxis=dict(autorange="reversed"), coloraxis_showscale=False)
                        show_chart(fig, height=340)

    # ---------- Row 3: Correlation / Heatmap ----------
    num_cols = [c for c in fdf.columns if pd.api.types.is_numeric_dtype(fdf[c])][:8]
    if len(num_cols) >= 3:
        st.markdown('<div class="section-title">🔥 Correlation Matrix</div>', unsafe_allow_html=True)
        corr = fdf[num_cols].corr()
        fig = px.imshow(
            corr, text_auto=".2f", aspect="auto",
            color_continuous_scale=[[0, "#EF4444"], [0.5, "#111827"], [1, "#3B82F6"]],
            zmin=-1, zmax=1, template="plotly_dark",
        )
        show_chart(fig, height=420)


# =========================================================
# RENDER: AI INSIGHTS (4 cards)
# =========================================================
def render_ai_insights(insights: dict):
    st.markdown('<div class="section-title">🧠 AI Insights</div>', unsafe_allow_html=True)

    def _ul(items):
        return "<ul>" + "".join(f"<li>{i}</li>" for i in items) + "</ul>"

    c1, c2 = st.columns(2)
    with c1:
        st.markdown(
            f'<div class="insight-card insight-happened">'
            f'<h4>✅ What Happened</h4>{_ul(insights["happened"])}</div>',
            unsafe_allow_html=True,
        )
    with c2:
        st.markdown(
            f'<div class="insight-card insight-risks">'
            f'<h4>⚠ Risks</h4>{_ul(insights["risks"])}</div>',
            unsafe_allow_html=True,
        )

    st.markdown("<br>", unsafe_allow_html=True)

    c3, c4 = st.columns(2)
    with c3:
        st.markdown(
            f'<div class="insight-card insight-opps">'
            f'<h4>📈 Opportunities</h4>{_ul(insights["opportunities"])}</div>',
            unsafe_allow_html=True,
        )
    with c4:
        st.markdown(
            f'<div class="insight-card insight-recs">'
            f'<h4>🎯 Recommended Actions</h4>{_ul(insights["actions"])}</div>',
            unsafe_allow_html=True,
        )


# =========================================================
# RENDER: EXPORT
# =========================================================
def to_excel_bytes(sheets: dict) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for name, frame in sheets.items():
            frame.to_excel(writer, sheet_name=name[:31], index=False)
    return buf.getvalue()


def render_export(fdf, profile, insights, ns):
    st.markdown('<div class="section-title">📥 Export</div>', unsafe_allow_html=True)
    c1, c2 = st.columns(2)

    c1.download_button(
        "📥 Download filtered data (CSV)",
        fdf.to_csv(index=False).encode("utf-8-sig"),
        f"filtered_{datetime.now():%Y%m%d_%H%M}.csv",
        "text/csv",
        use_container_width=True,
    )

    with c2:
        if len(fdf) > MAX_EXCEL_ROWS:
            st.caption(f"Excel export limited to {MAX_EXCEL_ROWS:,} rows – use CSV.")
        else:
            if st.button("Prepare Excel report", key=f"{ns}_xl", use_container_width=True):
                flat_insights = (
                    [f"Happened: {re.sub(r'[*][*]', '', i)}" for i in insights["happened"]]
                    + [f"Risk: {re.sub(r'[*][*]', '', i)}" for i in insights["risks"]]
                    + [f"Opportunity: {re.sub(r'[*][*]', '', i)}" for i in insights["opportunities"]]
                    + [f"Action: {re.sub(r'[*][*]', '', i)}" for i in insights["actions"]]
                )
                frames = {
                    "Filtered Data": fdf,
                    "Column Profile": profile,
                    "AI Insights": pd.DataFrame({"Insight": flat_insights}),
                }
                st.session_state[f"{ns}_xlsx"] = to_excel_bytes(frames)
            if f"{ns}_xlsx" in st.session_state:
                st.download_button(
                    "📥 Download Excel report",
                    st.session_state[f"{ns}_xlsx"],
                    f"business_copilot_{datetime.now():%Y%m%d_%H%M}.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )


# =========================================================
# SOURCE + FILTERS (sidebar)
# =========================================================
def get_source():
    with st.sidebar:
        st.markdown("### 🚀 Business Copilot AI")
        st.markdown("---")
        st.markdown("#### 📂 Upload")
        up = st.file_uploader("Upload dataset", type=["csv", "txt", "xlsx", "xls"], label_visibility="collapsed")
        use_sample = st.toggle("Use sample sales data", value=False)
    if up is not None:
        return up.name, up.getvalue()
    if use_sample:
        rng = np.random.default_rng(42)
        n = 1500
        units = rng.integers(1, 9, n)
        price = rng.gamma(4, 25, n).round(2)
        discount = rng.choice([0, 0.05, 0.1, 0.2], n)
        sample = pd.DataFrame({
            "Order ID": np.arange(10000, 10000 + n),
            "Order Date": (pd.Timestamp("2026-01-01") + pd.to_timedelta(rng.integers(0, 300, n), unit="D")).strftime("%Y-%m-%d"),
            "Region": rng.choice(["North", "South", "East", "West"], n, p=[.35, .25, .25, .15]),
            "Category": rng.choice(["Electronics", "Apparel", "Home", "Beauty", "Sports"], n),
            "Product": rng.choice(["Laptop", "Phone", "Shirt", "Desk", "Cream", "Ball", "Monitor", "Shoes", "Lamp", "Serum"], n),
            "Units": units,
            "Unit Price": price,
            "Discount": discount,
            "Revenue": (units * price * (1 - discount)).round(2),
        })
        sample.loc[rng.choice(n, 40, replace=False), "Region"] = np.nan
        return "sample_sales.csv", sample.to_csv(index=False).encode("utf-8")
    return None


def render_filters(df: pd.DataFrame, roles: dict, ns: str):
    st.sidebar.markdown("---")
    st.sidebar.markdown("#### 🔎 Filters")
    options = roles["categorical"] + roles["numeric"] + roles["datetime"]
    chosen = st.sidebar.multiselect("Filter on columns", options, key=f"{ns}_cols")
    mask = pd.Series(True, index=df.index)
    active = []
    for col in chosen:
        s, key = df[col], f"{ns}_f_{col}"
        if col in roles["categorical"]:
            opts = s.value_counts().index.tolist()[:MAX_FILTER_OPTIONS]
            picked = st.sidebar.multiselect(f"{col} (empty = all)", opts, key=key)
            if picked:
                mask &= s.isin(picked)
                active.append(f"{col}: {len(picked)} value(s)")
        elif col in roles["numeric"]:
            is_int = pd.api.types.is_integer_dtype(s)
            lo, hi = (int(s.min()), int(s.max())) if is_int else (float(s.min()), float(s.max()))
            if not lo < hi:
                continue
            rng = st.sidebar.slider(col, lo, hi, (lo, hi), key=key)
            if tuple(rng) != (lo, hi):
                mask &= s.between(rng[0], rng[1])
                active.append(f"{col}: {fmt_num(rng[0])} – {fmt_num(rng[1])}")
        else:
            lo, hi = s.min().date(), s.max().date()
            if lo == hi:
                continue
            rng = st.sidebar.date_input(col, (lo, hi), min_value=lo, max_value=hi, key=key)
            if isinstance(rng, (tuple, list)) and len(rng) == 2 and tuple(rng) != (lo, hi):
                mask &= (s >= pd.Timestamp(rng[0])) & (s < pd.Timestamp(rng[1]) + pd.Timedelta(days=1))
                active.append(f"{col}: {rng[0]} → {rng[1]}")
    return df[mask], active


def pick_index(cols: list, hints=METRIC_HINTS) -> int:
    for hint in hints:
        for i, c in enumerate(cols):
            if hint in c.lower():
                return i
    return 0


# =========================================================
# DASHBOARD TAB
# =========================================================
def render_dashboard(fdf, metrics, insights, summary_md, segment, date_col, metric, dataset_type):
    # 1. AI Summary
    render_ai_summary(summary_md)

    # 2. KPI cards
    cards = [
        {
            "title": "Records",
            "value": f"{metrics['total_rows']:,}",
            "subtitle": f"dataset: <span class='badge badge-blue'>{dataset_type}</span>",
            "icon": "📊",
            "color": ACCENT["blue"],
        },
        {
            "title": "Data Quality",
            "value": f"{metrics['completeness']:.1f}%",
            "subtitle": "cells filled",
            "icon": "🩺",
            "color": ACCENT["green"] if metrics["completeness"] > 90 else ACCENT["orange"],
        },
    ]
    if metric and metrics["total_metric"] is not None:
        cards.append({
            "title": f"Total {metric}",
            "value": fmt_num(metrics["total_metric"]),
            "subtitle": f"avg {fmt_num(metrics['avg_metric'])}",
            "icon": "💰",
            "color": ACCENT["purple"],
        })
    if metrics["top_segment"]:
        cards.append({
            "title": f"Top {segment}",
            "value": metrics["top_segment"],
            "subtitle": f"{metrics['top_segment_share']:.1f}% share",
            "icon": "🏆",
            "color": ACCENT["pink"],
        })
    if metrics["growth_pct"] is not None:
        g = metrics["growth_pct"]
        cards.append({
            "title": "Growth (30d)",
            "value": f"{g:+.1f}%",
            "subtitle": "vs prior 30 days",
            "icon": "📈" if g >= 0 else "📉",
            "color": ACCENT["green"] if g >= 0 else ACCENT["red"],
        })

    render_kpi_cards(cards)

    # 3. Visualizations
    render_visualizations(fdf, metric, segment, date_col)

    # 4. AI Insight Cards
    render_ai_insights(insights)


# =========================================================
# MAIN
# =========================================================
def main():
    source = get_source()
    if source is None:
        st.markdown(
            '<div class="app-header"><h1>🚀 Business Copilot AI</h1>'
            '<p>Upload a dataset to unlock automated KPIs, AI insights, and interactive analytics.</p></div>',
            unsafe_allow_html=True,
        )
        st.info("👈 Upload a CSV or Excel file in the sidebar, or switch on the sample sales data.")
        st.stop()

    name, data = source

    sheet = None
    if name.lower().endswith((".xlsx", ".xls")):
        try:
            sheets = excel_sheets(name, data)
        except ValueError as err:
            st.error(str(err))
            st.stop()
        sheet = st.sidebar.selectbox("Sheet", sheets) if len(sheets) > 1 else sheets[0]

    st.sidebar.markdown("---")
    dayfirst = st.sidebar.checkbox(
        "Dates are DD/MM/YYYY", value=False,
        help="Only matters for ambiguous dates such as 03/04/2026.",
    )

    try:
        df, roles, notes, profile = load_dataset(name, data, sheet, dayfirst)
    except ValueError as err:
        st.error(str(err))
        st.stop()

    ns = hashlib.md5(f"{name}|{sheet}|{dayfirst}|{len(data)}".encode()).hexdigest()[:8]
    loaded_at = st.session_state.setdefault(f"{ns}_loaded", datetime.now())

    num, cat, dts = roles["numeric"], roles["categorical"], roles["datetime"]

    # Analysis settings
    st.sidebar.markdown("---")
    st.sidebar.markdown("#### 🎯 Analysis Settings")
    metric = st.sidebar.selectbox(
        "Primary metric", num,
        index=pick_index(num), key=f"{ns}_metric",
    ) if num else None

    seg_pref = next((i for i, c in enumerate(cat) if 2 <= df[c].nunique() <= 30), 0)
    segment = st.sidebar.selectbox(
        "Segment by", cat, index=seg_pref, key=f"{ns}_seg",
    ) if cat else None

    date_col = st.sidebar.selectbox(
        "Date column", dts, key=f"{ns}_date",
    ) if dts else None

    # Filters
    fdf, active = render_filters(df, roles, ns)

    # Header
    render_header(name, sheet, len(fdf), len(df), loaded_at)
    if active:
        st.caption("Active filters: " + " | ".join(active))

    if fdf.empty:
        st.warning("No rows match the current filters.")
        st.stop()

    # Dataset classification
    dataset_type, _ = detect_dataset_type(df, roles)

    # AI metrics
    metrics = compute_ai_metrics(fdf, metric, segment, date_col)
    summary_md = generate_ai_summary_paragraph(metrics, dataset_type)
    insights = build_ai_insights(metrics, dataset_type)

    # ---- Tabs ----
    tab_dash, tab_ai, tab_data, tab_export = st.tabs([
        "🏠 Dashboard",
        "🤖 AI Insights",
        "🔍 Data Explorer",
        "📥 Export",
    ])

    with tab_dash:
        render_dashboard(fdf, metrics, insights, summary_md, segment, date_col, metric, dataset_type)

    with tab_ai:
        render_ai_summary(summary_md)
        render_ai_insights(insights)
        st.markdown('<div class="section-title">🎯 Recommended Next Steps</div>', unsafe_allow_html=True)
        st.markdown(
            f'<div class="insight-card insight-recs">'
            f'<h4>🧭 Strategic Priorities</h4>'
            f'<ul>'
            f'<li>Focus on <strong>{metrics["top_segment"] or "top segment"}</strong> — it drives the majority of {metric or "value"}.</li>'
            f'<li>Investigate <strong>{metrics["worst_segment"] or "weakest segment"}</strong> for recovery opportunities.</li>'
            f'<li>Maintain data quality above 95% for reliable insights.</li>'
            f'<li>Review dashboard weekly with leadership team.</li>'
            f'</ul></div>',
            unsafe_allow_html=True,
        )

    with tab_data:
        st.markdown('<div class="section-title">📋 Column Profile</div>', unsafe_allow_html=True)
        for n in notes:
            st.write(f"• {n}")
        show_df(profile)
        st.markdown('<div class="section-title">📄 Filtered Data</div>', unsafe_allow_html=True)
        st.caption(f"Previewing the first 1,000 of {len(fdf):,} rows.")
        show_df(fdf.head(1000))

    with tab_export:
        render_export(fdf, profile, insights, ns)


if __name__ == "__main__":
    main()