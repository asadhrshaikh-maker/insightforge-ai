
import io
import os
import json
import hashlib
import warnings
from datetime import datetime

import numpy as np
import pandas as pd
import plotly.express as px  # type: ignore[reportMissingTypeStubs]
import plotly.graph_objects as go  # type: ignore[reportMissingTypeStubs]
import streamlit as st

warnings.filterwarnings("ignore")

# ---- Optional heavy dependencies ----
try:
    from fpdf import FPDF
except Exception:
    pass
HAS_FPDF = "FPDF" in locals()

try:
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    HAS_PPTX = True
except Exception:
    HAS_PPTX = False

try:
    import requests
    HAS_REQUESTS = True
except Exception:
    HAS_REQUESTS = False

# =============================================================================
# 1. PAGE CONFIG  (must be first Streamlit command)
# =============================================================================
st.set_page_config(
    page_title="Business Copilot AI V3",
    page_icon="🚀",
    layout="wide",
    initial_sidebar_state="expanded",
)

# =============================================================================
# 2. GLOBAL CSS — Premium dark enterprise theme
# =============================================================================
st.markdown(
    """
<style>
:root{
  --bg-1:#050B18; --bg-2:#071326; --bg-3:#0B2344;
  --primary:#3B82F6; --purple:#8B5CF6; --green:#10B981;
  --orange:#F59E0B; --red:#EF4444;
  --text:#E5EEFF; --muted:#93A4C0;
}
.stApp{
  background:
    radial-gradient(1200px 800px at 8% -10%, #0B2344 0%, transparent 55%),
    radial-gradient(1000px 700px at 100% 0%, #071326 0%, transparent 55%),
    linear-gradient(180deg,#050B18 0%,#071326 55%,#0B2344 100%) !important;
  color:var(--text);
}
h1,h2,h3,h4,h5,h6{color:var(--text) !important;}
p,label,span,div{color:var(--text);}
section[data-testid="stSidebar"]{
  background:linear-gradient(180deg,#071326 0%,#050B18 100%) !important;
  border-right:1px solid rgba(59,130,246,.22);
}
.block-container{padding-top:1.8rem;padding-bottom:3rem;max-width:1550px;}

/* Hero */
.hero{
  background:linear-gradient(135deg,rgba(59,130,246,.24) 0%,rgba(139,92,246,.22) 50%,rgba(16,185,129,.18) 100%);
  border:1px solid rgba(139,92,246,.38);border-radius:22px;
  padding:32px 36px;margin-bottom:22px;
  box-shadow:0 20px 50px rgba(2,10,30,.55);
}
.hero h1{font-size:38px;margin:0 0 6px 0;}
.hero h3{color:#B9C8E6;font-weight:500;margin:0 0 16px 0;}
.hero ul{margin:0;padding-left:18px;}
.hero li{color:#D3DEF6;line-height:1.9;}

/* KPI card */
.kpi-card{
  background:linear-gradient(135deg,rgba(59,130,246,.14) 0%,rgba(139,92,246,.12) 100%);
  border:1px solid rgba(59,130,246,.38);border-radius:16px;
  padding:18px 20px;backdrop-filter:blur(10px);
  box-shadow:0 8px 24px rgba(2,10,30,.5);margin-bottom:12px;
}
.kpi-card .kpi-label{color:#93A4C0;font-size:11.5px;letter-spacing:1.1px;
  text-transform:uppercase;margin-bottom:6px;font-weight:600;}
.kpi-card .kpi-value{color:#E5EEFF;font-size:22px;font-weight:700;line-height:1.2;}
.kpi-card .kpi-hint{color:#93A4C0;font-size:11px;margin-top:4px;}

/* Section card */
.section-card{
  background:rgba(11,35,68,.55);border:1px solid rgba(59,130,246,.26);
  border-radius:16px;padding:20px 24px;margin-bottom:16px;
  backdrop-filter:blur(8px);box-shadow:0 8px 28px rgba(2,10,30,.35);
}
.section-card h4{margin-top:0;color:#E5EEFF;}

/* Insight chips */
.chip-info,.chip-warn,.chip-success,.chip-danger{
  border-radius:12px;padding:12px 16px;margin-bottom:10px;
  border-left:4px solid;background:rgba(11,35,68,.6);
  font-size:14.5px;line-height:1.65;
}
.chip-info{border-left-color:#3B82F6;}
.chip-warn{border-left-color:#F59E0B;}
.chip-success{border-left-color:#10B981;}
.chip-danger{border-left-color:#EF4444;}

/* Tabs */
.stTabs [data-baseweb="tab-list"]{gap:6px;flex-wrap:wrap;}
.stTabs [data-baseweb="tab"]{
  background:rgba(11,35,68,.6);border:1px solid rgba(59,130,246,.28);
  border-radius:10px;padding:10px 16px;color:#B9C8E6;font-weight:600;
}
.stTabs [aria-selected="true"]{
  background:linear-gradient(135deg,rgba(59,130,246,.38),rgba(139,92,246,.38));
  color:#FFF !important;border-color:rgba(139,92,246,.6);
}

/* Buttons */
.stButton > button,.stDownloadButton > button{
  background:linear-gradient(135deg,#3B82F6 0%,#8B5CF6 100%);
  color:#fff !important;border:none;border-radius:10px;
  padding:10px 20px;font-weight:600;
}
.stButton > button:hover,.stDownloadButton > button:hover{filter:brightness(1.12);}

/* Inputs */
.stTextInput input,.stNumberInput input,.stTextArea textarea{
  background:rgba(11,35,68,.85) !important;color:#E5EEFF !important;
  border-radius:10px !important;
}
.stSelectbox div[data-baseweb="select"] > div{
  background:rgba(11,35,68,.85) !important;color:#E5EEFF !important;
  border-radius:10px !important;
}

/* Chat */
.chat-user{
  background:rgba(59,130,246,.18);border:1px solid rgba(59,130,246,.45);
  border-radius:12px;padding:12px 16px;margin:8px 0 8px auto;
  max-width:82%;color:#E5EEFF;
}
.chat-bot{
  background:rgba(139,92,246,.14);border:1px solid rgba(139,92,246,.45);
  border-radius:12px;padding:12px 16px;margin:8px auto 8px 0;
  max-width:82%;color:#E5EEFF;
}

/* Dataframe corners */
.stDataFrame{border-radius:12px;overflow:hidden;}
</style>
""",
    unsafe_allow_html=True,
)

# =============================================================================
# 3. CONSTANTS + PLOTLY THEME
# =============================================================================
PLOTLY_LAYOUT = dict(
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(0,0,0,0)",
    font=dict(color="#E5EEFF", size=12),
    title_font=dict(color="#E5EEFF", size=16),
    xaxis=dict(gridcolor="rgba(59,130,246,0.15)",
               linecolor="rgba(59,130,246,0.3)"),
    yaxis=dict(gridcolor="rgba(59,130,246,0.15)",
               linecolor="rgba(59,130,246,0.3)"),
    legend=dict(bgcolor="rgba(0,0,0,0)"),
    margin=dict(l=20, r=20, t=55, b=20),
)

PALETTE = [
    "#3B82F6", "#8B5CF6", "#10B981", "#F59E0B", "#EF4444",
    "#06B6D4", "#EC4899", "#A3E635", "#6366F1", "#F97316",
]

MONEY_HINTS = [
    "revenue", "sales", "amount", "price", "cost", "profit",
    "expense", "budget", "income", "salary", "spend", "total",
    "value", "turnover", "margin", "gmv",
]

DATASET_KEYWORDS = {
    "Sales": ["revenue", "sales", "order", "product", "customer",
              "price", "quantity", "cart", "sku"],
    "Survey": ["survey", "interviewer", "respondent", "response",
               "question", "rating", "feedback", "score", "opinion"],
    "Finance": ["profit", "expense", "budget", "cost", "account",
                "invoice", "payment", "ledger", "debit", "credit"],
    "HR": ["employee", "department", "salary", "position", "hire",
           "manager", "attrition", "headcount", "designation", "leave"],
    "Operations": ["process", "operation", "machine", "downtime",
                   "cycle", "efficiency", "throughput", "latency"],
    "Inventory": ["inventory", "stock", "warehouse", "sku",
                  "supplier", "restock", "bin", "batch"],
    "Marketing": ["campaign", "click", "impression", "conversion",
                  "channel", "lead", "ad", "ctr", "spend"],
}

# =============================================================================
# 4. LOW-LEVEL HELPERS
# =============================================================================
def is_money_col(col) -> bool:
    name = str(col).lower()
    return any(h in name for h in MONEY_HINTS)


def format_number(value) -> str:
    try:
        n = float(value)
    except (TypeError, ValueError):
        return str(value)
    if pd.isna(n):
        return "N/A"
    a = abs(n)
    if a >= 1_000_000_000:
        return f"{n/1_000_000_000:,.2f}B"
    if a >= 1_000_000:
        return f"{n/1_000_000:,.2f}M"
    if a >= 1_000:
        return f"{n/1_000:,.1f}K"
    if a >= 1:
        return f"{n:,.2f}"
    return f"{n:,.4f}"


def sanitize_text(text) -> str:
    """Latin-1 safe text — required for fpdf2 core fonts."""
    if text is None:
        return ""
    if not isinstance(text, str):
        text = str(text)
    repl = {
        "₹": "Rs.", "€": "EUR", "£": "GBP", "¥": "JPY",
        "–": "-", "—": "-", "•": "-", "…": "...",
        "“": '"', "”": '"', "‘": "'", "’": "'",
        "→": "->", "←": "<-", "↑": "^", "↓": "v",
    }
    for k, v in repl.items():
        text = text.replace(k, v)
    # strip non-latin1 (e.g. emoji)
    out_chars = []
    for ch in text:
        try:
            ch.encode("latin-1")
            out_chars.append(ch)
        except UnicodeEncodeError:
            out_chars.append("")
    return "".join(out_chars)


def clean_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    seen = {}
    new_cols = []
    for c in df.columns:
        name = str(c).strip()
        if not name or name.lower().startswith("unnamed"):
            name = "Unnamed"
        if name in seen:
            seen[name] += 1
            name = f"{name}_{seen[name]}"
        else:
            seen[name] = 0
        new_cols.append(name)
    df.columns = new_cols
    return df


def detect_date_columns(df: pd.DataFrame) -> list:
    out = []
    for c in df.columns:
        s = df[c]
        if pd.api.types.is_datetime64_any_dtype(s):
            out.append(c)
            continue
        if pd.api.types.is_numeric_dtype(s) or pd.api.types.is_bool_dtype(s):
            continue
        sample = s.dropna().astype(str).head(120)
        if len(sample) == 0:
            continue
        try:
            try:
                parsed = pd.to_datetime(sample, errors="coerce", format="mixed")
            except (TypeError, ValueError):
                parsed = pd.to_datetime(sample, errors="coerce")
        except Exception:
            continue
        if pd.Series(parsed).notna().mean() < 0.75:
            continue
        name_hint = any(k in str(c).lower() for k in
                        ["date", "time", "day", "month", "year",
                         "period", "created", "updated", "timestamp"])
        try:
            has_sep = sample.str.contains(r"[-/:]", regex=True).mean() >= 0.5
        except Exception:
            has_sep = False
        if name_hint or has_sep:
            out.append(c)
    return out


def detect_numeric_columns(df: pd.DataFrame) -> list:
    return [c for c in df.columns
            if pd.api.types.is_numeric_dtype(df[c])
            and not pd.api.types.is_bool_dtype(df[c])]


def detect_categorical_columns(df: pd.DataFrame, exclude=()) -> list:
    out = []
    n = max(len(df), 1)
    for c in df.columns:
        if c in exclude:
            continue
        s = df[c]
        if (pd.api.types.is_numeric_dtype(s)
                or pd.api.types.is_datetime64_any_dtype(s)
                or pd.api.types.is_bool_dtype(s)):
            continue
        try:
            nu = int(s.nunique(dropna=True))
        except Exception:
            continue
        if 1 <= nu <= 60 and nu < n:
            out.append(c)
    return out


def detect_dataset_type(df: pd.DataFrame) -> str:
    cols_text = " ".join(df.columns.astype(str).str.lower())
    scores = {}
    for label, kws in DATASET_KEYWORDS.items():
        scores[label] = sum(1 for k in kws if k in cols_text)
    best_label, best_score = "General", 0
    for label, score in scores.items():
        if score > best_score:
            best_label, best_score = label, score
    return best_label if best_score >= 2 else "General"


def pick_primary_metric(df: pd.DataFrame, num_cols: list):
    priority = ["revenue", "sales", "amount", "total", "profit",
                "income", "price", "value", "cost", "quantity", "count"]
    for p in priority:
        for c in num_cols:
            if p in str(c).lower():
                return c
    return num_cols[0] if num_cols else None


def top_category(df: pd.DataFrame, cat_cols: list):
    for c in cat_cols:
        vc = df[c].dropna().astype(str).value_counts()
        if len(vc):
            return c, str(vc.index[0]), int(vc.iloc[0])
    return None, None, 0


def missing_percentage(df: pd.DataFrame) -> float:
    if df.empty or df.size == 0:
        return 0.0
    return float(df.isna().sum().sum()) / float(df.size) * 100.0


# =============================================================================
# 5. DATA LOADER (cached)
# =============================================================================
@st.cache_data(show_spinner=False)
def load_data(file_bytes: bytes, file_name: str) -> pd.DataFrame:
    name = str(file_name).lower()

    if name.endswith(".csv"):
        last_err = None
        for enc in ("utf-8", "utf-8-sig", "latin-1", "cp1252"):
            try:
                return clean_columns(pd.read_csv(io.BytesIO(file_bytes),
                                                 encoding=enc))
            except Exception as e:
                last_err = e
        try:
            return clean_columns(pd.read_csv(io.BytesIO(file_bytes),
                                             sep=None, engine="python",
                                             encoding="utf-8", errors="ignore"))
        except Exception:
            raise ValueError(f"Unable to read CSV. ({last_err})")

    try:
        return clean_columns(pd.read_excel(io.BytesIO(file_bytes)))
    except Exception as e:
        raise ValueError(f"Unable to read Excel file. ({e})")


# =============================================================================
# 6. KPI + SUMMARY
# =============================================================================
def compute_kpis(df, dataset_type, num_cols, cat_cols, date_cols) -> dict:
    metric = pick_primary_metric(df, num_cols)
    missing_pct = missing_percentage(df)
    dup_rows = int(df.duplicated().sum())
    dup_pct = (dup_rows / max(len(df), 1)) * 100.0

    # Date range
    date_range_str = "N/A"
    if date_cols:
        try:
            d = pd.to_datetime(df[date_cols[0]], errors="coerce").dropna()
            if len(d):
                date_range_str = f"{d.min():%d %b %Y} → {d.max():%d %b %Y}"
        except Exception:
            pass

    # Top category
    top_c, top_v, top_n = top_category(df, cat_cols)
    top_cat_str = f"{top_v}" if top_v else "N/A"

    # Primary metric value
    metric_value_str = "N/A"
    if metric is not None:
        try:
            total = pd.to_numeric(df[metric], errors="coerce").sum()
            prefix = "₹" if is_money_col(metric) else ""
            metric_value_str = f"{prefix}{format_number(total)}"
        except Exception:
            pass

    # Growth %
    growth_pct = None
    if date_cols and metric is not None:
        try:
            t = df[[date_cols[0], metric]].copy()
            t[date_cols[0]] = pd.to_datetime(t[date_cols[0]], errors="coerce")
            t[metric] = pd.to_numeric(t[metric], errors="coerce")
            t = t.dropna().sort_values(date_cols[0])
            if len(t) >= 8:
                half = len(t) // 2
                a = float(t.iloc[:half][metric].sum())
                b = float(t.iloc[half:][metric].sum())
                if a > 0:
                    growth_pct = (b - a) / a * 100.0
        except Exception:
            pass

    # Data quality score
    try:
        quality = 100.0 - missing_pct - (dup_pct * 0.5)
        quality = max(0.0, min(100.0, quality))
    except Exception:
        quality = 0.0

    return {
        "Records": len(df),
        "Columns": df.shape[1],
        "Missing %": f"{missing_pct:.2f}%",
        "Duplicate Rows": dup_rows,
        "Date Range": date_range_str,
        "Top Category": top_cat_str,
        "Primary Metric": metric if metric else "N/A",
        "Primary Value": metric_value_str,
        "Data Quality": f"{quality:.1f}/100",
        "Growth %": f"{growth_pct:+.1f}%" if growth_pct is not None else "N/A",
        # raw numeric for later use
        "_missing_pct": missing_pct,
        "_dup_pct": dup_pct,
        "_growth_pct": growth_pct,
        "_quality": quality,
        "_metric": metric,
    }


def generate_executive_summary(df, dataset_type, kpis,
                               cat_cols, date_cols) -> str:
    parts = []
    parts.append(
        f"The dataset contains **{len(df):,}** rows across "
        f"**{df.shape[1]}** columns and has been automatically "
        f"classified as a **{dataset_type} dataset**."
    )

    metric = kpis.get("_metric")
    if metric:
        parts.append(
            f"The primary measure is **{metric}**, totalling "
            f"**{kpis['Primary Value']}**."
        )

    top_cat = kpis.get("Top Category")
    if top_cat and top_cat != "N/A":
        try:
            cat_c, _, _ = top_category(df, cat_cols)
            share = 0.0
            if cat_c:
                vc = df[cat_c].dropna().astype(str).value_counts()
                if len(vc):
                    share = vc.iloc[0] / max(vc.sum(), 1) * 100
            parts.append(
                f"**{top_cat}** is the leading segment in **{cat_c}**, "
                f"contributing approximately **{share:.1f}%** of records."
            )
        except Exception:
            pass

    if kpis.get("_growth_pct") is not None:
        g = kpis["_growth_pct"]
        direction = "grew" if g >= 0 else "declined"
        parts.append(
            f"The primary metric **{direction} {abs(g):.1f}%** between the "
            f"earlier and later halves of the timeline."
        )
    else:
        parts.append(
            "Growth trend could not be computed — no reliable date column "
            "or insufficient data volume was detected."
        )

    mp = kpis.get("_missing_pct", 0)
    if mp < 2:
        parts.append(f"Data quality remains **high** with only "
                     f"**{mp:.2f}%** missing values.")
    elif mp < 10:
        parts.append(f"Data quality is **acceptable** — "
                     f"**{mp:.2f}%** of values are missing.")
    else:
        parts.append(f"Data quality requires attention — "
                     f"**{mp:.2f}%** of values are missing.")

    if top_cat and top_cat != "N/A":
        parts.append(
            f"**Recommendation:** consolidate performance in "
            f"**{top_cat}** while exploring under-performing segments."
        )
    else:
        parts.append(
            "**Recommendation:** enrich the dataset with categorical and "
            "date fields to enable deeper executive recommendations."
        )

    return " ".join(parts)


# =============================================================================
# 7. AI INSIGHTS ENGINE
# =============================================================================
def generate_ai_insights(df, num_cols, cat_cols, date_cols, kpis) -> dict:
    metric = pick_primary_metric(df, num_cols)
    cat_c, cat_v, cat_n = top_category(df, cat_cols)
    missing_pct = kpis.get("_missing_pct", 0.0)

    happened, risks, opps, recs = [], [], [], []

    # ---------- WHAT HAPPENED ----------
    if cat_c:
        share = cat_n / max(len(df), 1) * 100
        happened.append(
            f"Top contributor: **{cat_v}** in **{cat_c}** "
            f"({cat_n:,} records, {share:.1f}% share)."
        )
    if metric:
        try:
            series = pd.to_numeric(df[metric], errors="coerce").dropna()
            if len(series):
                hi, lo = series.max(), series.min()
                prefix = "₹" if is_money_col(metric) else ""
                happened.append(
                    f"Highest **{metric}** recorded: "
                    f"**{prefix}{format_number(hi)}**; "
                    f"lowest: **{prefix}{format_number(lo)}**."
                )
        except Exception:
            pass
    if kpis.get("_growth_pct") is not None:
        g = kpis["_growth_pct"]
        verb = "increase" if g >= 0 else "decrease"
        happened.append(
            f"Period-over-period **{verb}** of **{abs(g):.1f}%** in the "
            f"primary metric."
        )

    # ---------- RISKS ----------
    try:
        miss_series = df.isna().mean().sort_values(ascending=False)
        if len(miss_series) and float(miss_series.iloc[0]) > 0.15:
            risks.append(
                f"High missingness in **{miss_series.index[0]}** "
                f"({float(miss_series.iloc[0])*100:.1f}%) may bias analysis."
            )
    except Exception:
        pass

    if cat_c:
        try:
            vc = df[cat_c].dropna().astype(str).value_counts()
            if len(vc) >= 2:
                share_top = vc.iloc[0] / vc.sum() * 100
                if share_top > 40:
                    risks.append(
                        f"Concentration risk: **{vc.index[0]}** accounts for "
                        f"{share_top:.1f}% of **{cat_c}** — over-reliance on "
                        f"one segment."
                    )
        except Exception:
            pass

    if metric:
        try:
            s = pd.to_numeric(df[metric], errors="coerce").dropna()
            if len(s) >= 30:
                q1, q3 = s.quantile(0.25), s.quantile(0.75)
                iqr = q3 - q1
                outliers = int(((s < q1 - 1.5 * iqr) |
                                (s > q3 + 1.5 * iqr)).sum())
                if outliers > 0:
                    risks.append(
                        f"Detected **{outliers}** potential outliers in "
                        f"**{metric}** — verify or treat before reporting."
                    )
        except Exception:
            pass

    if kpis.get("_growth_pct") is not None and kpis["_growth_pct"] < -5:
        risks.append(
            f"Declining trend: primary metric dropped "
            f"**{abs(kpis['_growth_pct']):.1f}%** — investigate drivers."
        )

    if not risks:
        risks.append("No significant risks were detected in the current dataset.")

    # ---------- OPPORTUNITIES ----------
    if date_cols and metric and cat_c:
        try:
            date_col = date_cols[0]
            t = df[[date_col, metric, cat_c]].copy()
            t[date_col] = pd.to_datetime(t[date_col], errors="coerce")
            t[metric] = pd.to_numeric(t[metric], errors="coerce")
            t = t.dropna(subset=[date_col, metric])
            if len(t) >= 20:
                mid = t[date_col].median()
                early = t[t[date_col] < mid].groupby(cat_c)[metric].sum()
                late = t[t[date_col] >= mid].groupby(cat_c)[metric].sum()
                for cat in late.index:
                    if cat in early.index and early[cat] > 0:
                        g = (late[cat] - early[cat]) / early[cat] * 100
                        if g > 20:
                            opps.append(
                                f"Growth segment: **{cat}** in **{cat_c}** "
                                f"grew **{g:.1f}%** recently — scale investment."
                            )
                # underutilised
                if not opps:
                    vc = df[cat_c].dropna().astype(str).value_counts()
                    if len(vc) >= 3:
                        smallest = vc.tail(2)
                        for name, cnt in smallest.items():
                            opps.append(
                                f"Underutilised segment: **{name}** in "
                                f"**{cat_c}** has only {cnt} records — "
                                f"candidate for expansion."
                            )
        except Exception:
            pass

    if not opps:
        opps.append(
            "No emerging segments were identified — collect more "
            "segmented data to surface growth opportunities."
        )

    # ---------- RECOMMENDATIONS ----------
    if missing_pct > 10:
        recs.append(
            "Prioritise data cleaning — missingness exceeds 10% and "
            "undermines decision confidence."
        )
    if cat_v and cat_c:
        recs.append(
            f"Double down on **{cat_v}** in **{cat_c}** — the proven "
            f"top-performer for sustainable returns."
        )
    if kpis.get("_growth_pct") is not None:
        if kpis["_growth_pct"] < 0:
            recs.append(
                "Run a root-cause analysis on declining segments and "
                "adjust the operating plan this quarter."
            )
        else:
            recs.append(
                "Maintain momentum — institutionalise practices driving "
                "current growth."
            )
    recs.append(
        "Automate weekly executive review of the KPI cards and AI "
        "insights to catch trends early."
    )

    return {
        "happened": happened,
        "risks": risks,
        "opportunities": opps,
        "recommendations": recs,
    }


# =============================================================================
# 8. ROOT CAUSE ANALYSIS
# =============================================================================
def run_root_cause_analysis(df, num_cols, cat_cols, date_cols) -> dict:
    metric = pick_primary_metric(df, num_cols)
    if not metric:
        return {"ok": False,
                "message": "No numeric metric detected — cannot run root cause analysis."}
    if not date_cols:
        return {"ok": False,
                "message": "No date column detected — root cause analysis requires time-series data."}
    if not cat_cols:
        return {"ok": False,
                "message": "No categorical dimensions found — root cause analysis needs at least one category."}

    date_col = date_cols[0]
    try:
        base = df[[date_col, metric] + cat_cols[:5]].copy()
        base[date_col] = pd.to_datetime(base[date_col], errors="coerce")
        base[metric] = pd.to_numeric(base[metric], errors="coerce")
        base = base.dropna(subset=[date_col, metric])
        if len(base) < 20:
            return {"ok": False,
                    "message": "Not enough data points (min 20) to run analysis."}

        mid = base[date_col].median()
        early = base[base[date_col] < mid]
        late = base[base[date_col] >= mid]

        early_sum = float(early[metric].sum())
        late_sum = float(late[metric].sum())
        change = late_sum - early_sum
        change_pct = (change / early_sum * 100) if early_sum != 0 else 0.0

        drivers = []
        for c in cat_cols[:5]:
            if c not in base.columns:
                continue
            try:
                e_g = early.groupby(c)[metric].sum()
                l_g = late.groupby(c)[metric].sum()
                idx = e_g.index.union(l_g.index)
                diff = (l_g.reindex(idx, fill_value=0)
                        - e_g.reindex(idx, fill_value=0))
                diff = diff.sort_values()
                if len(diff) > 0:
                    drivers.append({
                        "dimension": c,
                        "decliner": (str(diff.index[0]), float(diff.iloc[0])),
                        "gainer": (str(diff.index[-1]), float(diff.iloc[-1])),
                        "all": diff,
                    })
            except Exception:
                continue

        # Score dimensions by absolute movement
        drivers = sorted(drivers,
                         key=lambda d: abs(d["gainer"][1]) + abs(d["decliner"][1]),
                         reverse=True)

        explanation_lines = []
        direction = "growth" if change >= 0 else "decline"
        explanation_lines.append(
            f"Between the two halves of the timeline, **{metric}** experienced "
            f"a **{direction} of {abs(change_pct):.1f}%** "
            f"(from {format_number(early_sum)} to {format_number(late_sum)})."
        )

        if drivers:
            d0 = drivers[0]
            explanation_lines.append(
                f"The dominant driver is **{d0['dimension']}**: "
                f"**{d0['gainer'][0]}** contributed the biggest positive change "
                f"(+{format_number(d0['gainer'][1])}) while "
                f"**{d0['decliner'][0]}** dragged the total "
                f"({format_number(d0['decliner'][1])})."
            )
            for d in drivers[1:3]:
                explanation_lines.append(
                    f"Secondary driver — **{d['dimension']}**: "
                    f"**{d['gainer'][0]}** gained "
                    f"{format_number(d['gainer'][1])}; "
                    f"**{d['decliner'][0]}** lost "
                    f"{format_number(d['decliner'][1])}."
                )

        return {
            "ok": True,
            "metric": metric,
            "early_total": early_sum,
            "late_total": late_sum,
            "change": change,
            "change_pct": change_pct,
            "drivers": drivers,
            "explanation": explanation_lines,
        }
    except Exception as e:
        return {"ok": False, "message": f"Analysis failed: {e}"}


# =============================================================================
# 9. CHART BUILDERS
# =============================================================================
def _apply_theme(fig):
    fig.update_layout(**PLOTLY_LAYOUT)
    return fig


def chart_bar(df, cat_c, metric):
    try:
        if metric:
            g = (df[[cat_c, metric]].dropna()
                 .assign(**{metric: lambda d: pd.to_numeric(d[metric], errors="coerce")})
                 .dropna()
                 .groupby(cat_c)[metric].sum()
                 .sort_values(ascending=False).head(10).reset_index())
            if g.empty:
                return None
            fig = px.bar(g, x=cat_c, y=metric, text=metric,
                         color=metric, color_continuous_scale="Blues",
                         title=f"Top 10 {cat_c} by {metric}")
            fig.update_traces(texttemplate="%{text:,.0f}",
                              textposition="outside")
            fig.update_layout(coloraxis_showscale=False)
        else:
            g = df[cat_c].value_counts().head(10).reset_index()
            g.columns = [cat_c, "Count"]
            fig = px.bar(g, x=cat_c, y="Count", text="Count",
                         color="Count", color_continuous_scale="Blues",
                         title=f"Top 10 {cat_c}")
            fig.update_layout(coloraxis_showscale=False)
        return _apply_theme(fig)
    except Exception:
        return None


def chart_pie(df, cat_c):
    try:
        vc = df[cat_c].dropna().astype(str).value_counts()
        if vc.empty:
            return None
        if len(vc) > 8:
            vc = pd.concat([vc.head(8), pd.Series({"Other": int(vc.iloc[8:].sum())})])
        g = vc.reset_index()
        g.columns = [cat_c, "Records"]
        fig = px.pie(g, names=cat_c, values="Records",
                     color_discrete_sequence=PALETTE,
                     title=f"{cat_c} — Pie")
        fig.update_traces(textinfo="percent+label")
        return _apply_theme(fig)
    except Exception:
        return None


def chart_donut(df, cat_c):
    try:
        vc = df[cat_c].dropna().astype(str).value_counts()
        if vc.empty:
            return None
        if len(vc) > 8:
            vc = pd.concat([vc.head(8), pd.Series({"Other": int(vc.iloc[8:].sum())})])
        g = vc.reset_index()
        g.columns = [cat_c, "Records"]
        fig = px.pie(g, names=cat_c, values="Records", hole=0.5,
                     color_discrete_sequence=PALETTE,
                     title=f"{cat_c} — Donut")
        fig.update_traces(textinfo="percent+label")
        return _apply_theme(fig)
    except Exception:
        return None


def chart_line(df, date_c, metric):
    try:
        tmp = df[[date_c] + ([metric] if metric else [])].copy()
        tmp[date_c] = pd.to_datetime(tmp[date_c], errors="coerce")
        tmp = tmp.dropna(subset=[date_c])
        if tmp.empty:
            return None
        span = (tmp[date_c].max() - tmp[date_c].min()).days
        if span <= 62:
            tmp["_p"] = tmp[date_c].dt.floor("D")
        elif span <= 730:
            tmp["_p"] = tmp[date_c].dt.to_period("W").dt.start_time
        else:
            tmp["_p"] = tmp[date_c].dt.to_period("M").dt.to_timestamp()
        if metric:
            tmp[metric] = pd.to_numeric(tmp[metric], errors="coerce")
            s = tmp.groupby("_p")[metric].sum()
            y = metric
        else:
            s = tmp.groupby("_p").size()
            y = "Records"
        agg = s.reset_index()
        agg.columns = ["Period", y]
        fig = px.line(agg, x="Period", y=y, markers=True,
                      title=f"{y} Trend Over Time")
        fig.update_traces(line=dict(width=3))
        return _apply_theme(fig)
    except Exception:
        return None


def chart_area(df, date_c, metric):
    try:
        tmp = df[[date_c] + ([metric] if metric else [])].copy()
        tmp[date_c] = pd.to_datetime(tmp[date_c], errors="coerce")
        tmp = tmp.dropna(subset=[date_c])
        if tmp.empty:
            return None
        span = (tmp[date_c].max() - tmp[date_c].min()).days
        if span <= 62:
            tmp["_p"] = tmp[date_c].dt.floor("D")
        elif span <= 730:
            tmp["_p"] = tmp[date_c].dt.to_period("W").dt.start_time
        else:
            tmp["_p"] = tmp[date_c].dt.to_period("M").dt.to_timestamp()
        if metric:
            tmp[metric] = pd.to_numeric(tmp[metric], errors="coerce")
            s = tmp.groupby("_p")[metric].sum()
            y = metric
        else:
            s = tmp.groupby("_p").size()
            y = "Records"
        agg = s.reset_index()
        agg.columns = ["Period", y]
        fig = px.area(agg, x="Period", y=y, title=f"{y} Area Trend")
        return _apply_theme(fig)
    except Exception:
        return None


def chart_scatter(df, num_cols):
    try:
        if len(num_cols) < 2:
            return None
        x, y = num_cols[0], num_cols[1]
        s = df[[x, y]].dropna()
        if s.empty:
            return None
        fig = px.scatter(s, x=x, y=y, opacity=0.65,
                         color_discrete_sequence=["#3B82F6"],
                         title=f"{x} vs {y} — Scatter")
        return _apply_theme(fig)
    except Exception:
        return None


def chart_box(df, num_cols):
    try:
        if not num_cols:
            return None
        s = df[num_cols[:5]].apply(pd.to_numeric, errors="coerce")
        fig = px.box(s, title="Distribution of Numeric Columns",
                     color_discrete_sequence=PALETTE)
        return _apply_theme(fig)
    except Exception:
        return None


def chart_heatmap_corr(df, num_cols):
    try:
        if len(num_cols) < 2:
            return None
        s = df[num_cols[:10]].apply(pd.to_numeric, errors="coerce")
        corr = s.corr()
        fig = px.imshow(corr, text_auto=".2f", color_continuous_scale="RdBu_r",
                        zmin=-1, zmax=1, title="Correlation Heatmap")
        return _apply_theme(fig)
    except Exception:
        return None


def chart_treemap(df, cat_cols, metric):
    try:
        if not cat_cols:
            return None
        if len(cat_cols) >= 2:
            c1, c2 = cat_cols[0], cat_cols[1]
            if metric:
                g = (df[[c1, c2, metric]].dropna()
                     .assign(**{metric: lambda d: pd.to_numeric(d[metric], errors="coerce")})
                     .dropna().groupby([c1, c2])[metric].sum().reset_index())
                if g.empty:
                    return None
                fig = px.treemap(g, path=[c1, c2], values=metric,
                                 color=metric, color_continuous_scale="Viridis",
                                 title=f"{c1} → {c2} Treemap")
            else:
                g = df.groupby([c1, c2]).size().reset_index(name="Count")
                fig = px.treemap(g, path=[c1, c2], values="Count",
                                 color="Count", color_continuous_scale="Viridis",
                                 title=f"{c1} → {c2} Treemap")
        else:
            c1 = cat_cols[0]
            vc = df[c1].value_counts().reset_index()
            vc.columns = [c1, "Count"]
            fig = px.treemap(vc, path=[c1], values="Count",
                             title=f"{c1} Treemap")
        return _apply_theme(fig)
    except Exception:
        return None


def chart_sunburst(df, cat_cols):
    try:
        if len(cat_cols) < 2:
            return None
        c1, c2 = cat_cols[0], cat_cols[1]
        g = df.groupby([c1, c2]).size().reset_index(name="Count")
        fig = px.sunburst(g, path=[c1, c2], values="Count",
                          title=f"{c1} → {c2} Sunburst",
                          color="Count", color_continuous_scale="Plasma")
        return _apply_theme(fig)
    except Exception:
        return None


def chart_top10(df, cat_c, metric):
    try:
        if metric:
            g = (df[[cat_c, metric]].dropna()
                 .assign(**{metric: lambda d: pd.to_numeric(d[metric], errors="coerce")})
                 .dropna().groupby(cat_c)[metric].sum()
                 .sort_values(ascending=False).head(10).reset_index())
            if g.empty:
                return None
            fig = px.bar(g, x=metric, y=cat_c, orientation="h",
                         text=metric, color=metric,
                         color_continuous_scale="Purples",
                         title=f"Top 10 {cat_c} by {metric}")
            fig.update_traces(texttemplate="%{text:,.0f}",
                              textposition="outside")
            fig.update_layout(coloraxis_showscale=False,
                              yaxis=dict(autorange="reversed"))
        else:
            vc = df[cat_c].value_counts().head(10).reset_index()
            vc.columns = [cat_c, "Count"]
            fig = px.bar(vc, x="Count", y=cat_c, orientation="h",
                         text="Count", color="Count",
                         color_continuous_scale="Purples",
                         title=f"Top 10 {cat_c}")
            fig.update_layout(coloraxis_showscale=False,
                              yaxis=dict(autorange="reversed"))
        return _apply_theme(fig)
    except Exception:
        return None


def build_auto_charts(df, cat_cols, num_cols, date_cols, metric) -> list:
    charts = []
    cat_c = cat_cols[0] if cat_cols else None
    date_c = date_cols[0] if date_cols else None

    if cat_c:
        f = chart_bar(df, cat_c, metric)
        if f is not None:
            charts.append(("Bar — Top Categories", f))
        f = chart_pie(df, cat_c)
        if f is not None:
            charts.append(("Pie — Category Share", f))
        f = chart_donut(df, cat_c)
        if f is not None:
            charts.append(("Donut — Category Split", f))
        f = chart_top10(df, cat_c, metric)
        if f is not None:
            charts.append(("Top 10 Ranking", f))

    if date_c:
        f = chart_line(df, date_c, metric)
        if f is not None:
            charts.append(("Line — Trend", f))
        f = chart_area(df, date_c, metric)
        if f is not None:
            charts.append(("Area — Trend", f))

    if len(num_cols) >= 2:
        f = chart_scatter(df, num_cols)
        if f is not None:
            charts.append(("Scatter — Relationship", f))
        f = chart_heatmap_corr(df, num_cols)
        if f is not None:
            charts.append(("Heatmap — Correlation", f))

    if num_cols:
        f = chart_box(df, num_cols)
        if f is not None:
            charts.append(("Box Plot — Distribution", f))

    if len(cat_cols) >= 2:
        f = chart_treemap(df, cat_cols, metric)
        if f is not None:
            charts.append(("Treemap", f))
        f = chart_sunburst(df, cat_cols)
        if f is not None:
            charts.append(("Sunburst", f))

    return charts


# =============================================================================
# 10. AI CHAT (Ollama / OpenAI)
# =============================================================================
def build_chat_context(df, dataset_type, kpis, summary, insights) -> str:
    cat_cols = detect_categorical_columns(df, exclude=detect_date_columns(df))
    metric = kpis.get("_metric")
    ctx = f"""You are a senior business analyst working inside Business Copilot AI V3.
Answer the user's questions about the uploaded dataset concisely, professionally, and data-driven.

DATASET:
- Type: {dataset_type}
- Rows: {len(df):,}
- Columns: {df.shape[1]}
- Column names: {', '.join(map(str, df.columns[:40].tolist()))}

KEY KPIs:
- Primary metric: {metric}
- Primary value: {kpis['Primary Value']}
- Top category: {kpis['Top Category']}
- Date range: {kpis['Date Range']}
- Growth: {kpis['Growth %']}
- Data quality: {kpis['Data Quality']}
- Missing %: {kpis['Missing %']}

EXECUTIVE SUMMARY:
{summary}

WHAT HAPPENED:
{chr(10).join('- ' + x for x in insights['happened'])}

RISKS:
{chr(10).join('- ' + x for x in insights['risks'])}

OPPORTUNITIES:
{chr(10).join('- ' + x for x in insights['opportunities'])}

RECOMMENDATIONS:
{chr(10).join('- ' + x for x in insights['recommendations'])}

INSTRUCTIONS:
- Answer in business language.
- Be specific with numbers where relevant.
- If the user asks about something not in the data, say so clearly.
- Keep answers concise (3-8 sentences).
"""
    return ctx


def call_ollama(prompt: str, model: str, base_url: str) -> str:
    if not HAS_REQUESTS:
        return "❌ `requests` library is not installed — cannot connect to Ollama."
    try:
        url = base_url.rstrip("/") + "/api/generate"
        resp = requests.post(
            url,
            json={"model": model, "prompt": prompt, "stream": False},
            timeout=120,
        )
        if resp.status_code == 200:
            data = resp.json()
            return data.get("response", "").strip() or "⚠️ Empty response from Ollama."
        return f"❌ Ollama error {resp.status_code}: {resp.text[:300]}"
    except Exception as e:
        return f"❌ Ollama connection failed: {e}. Is `ollama serve` running?"


def call_openai(prompt: str, api_key: str, model: str = "gpt-4o-mini") -> str:
    if not HAS_REQUESTS:
        return "❌ `requests` library is not installed — cannot call OpenAI."
    try:
        resp = requests.post(
            "https://api.openai.com/v1/chat/completions",
            headers={
                "Authorization": f"Bearer {api_key}",
                "Content-Type": "application/json",
            },
            json={
                "model": model,
                "messages": [{"role": "user", "content": prompt}],
                "temperature": 0.3,
            },
            timeout=90,
        )
        if resp.status_code == 200:
            return resp.json()["choices"][0]["message"]["content"].strip()
        return f"❌ OpenAI error {resp.status_code}: {resp.text[:300]}"
    except Exception as e:
        return f"❌ OpenAI connection failed: {e}"


# =============================================================================
# 11. PDF REPORT
# =============================================================================
def build_pdf_report(kpis, summary, insights, rca, dataset_type,
                     df_columns, charts: list) -> bytes:
    if not HAS_FPDF:
        raise RuntimeError("fpdf2 not installed — cannot build PDF.")

    pdf = FPDF(format="A4")
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()

    # Title band
    pdf.set_fill_color(59, 130, 246)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("helvetica", "B", 20)
    pdf.cell(0, 14, " Business Copilot AI V3 ", ln=True, align="C", fill=True)
    pdf.ln(2)
    pdf.set_font("helvetica", "", 11)
    pdf.cell(0, 8,
             f"Management Report — {datetime.now():%d %b %Y %H:%M}",
             ln=True, align="C")
    pdf.set_text_color(30, 30, 30)
    pdf.ln(6)

    def section(title):
        pdf.set_font("helvetica", "B", 14)
        pdf.set_text_color(30, 60, 120)
        pdf.cell(0, 10, sanitize_text(title), ln=True)
        pdf.set_text_color(40, 40, 40)
        pdf.set_font("helvetica", "", 10)
        pdf.ln(1)

    def para(text):
        try:
            pdf.multi_cell(0, 6, sanitize_text(text))
        except Exception:
            pass
        pdf.ln(1)

    section("1. Dataset Overview")
    para(f"Detected type: {dataset_type}")
    para(f"Columns: {', '.join(map(str, df_columns[:40]))}")

    section("2. KPI Summary")
    for k, v in kpis.items():
        if str(k).startswith("_"):
            continue
        para(f"- {k}: {v}")

    section("3. Executive Summary")
    para(summary.replace("**", ""))

    section("4. What Happened")
    for x in insights["happened"]:
        para(f"- {x.replace('**', '')}")

    section("5. Risks")
    for x in insights["risks"]:
        para(f"- {x.replace('**', '')}")

    section("6. Opportunities")
    for x in insights["opportunities"]:
        para(f"- {x.replace('**', '')}")

    section("7. Recommendations")
    for x in insights["recommendations"]:
        para(f"- {x.replace('**', '')}")

    if rca.get("ok"):
        section("8. Root Cause Analysis")
        for line in rca.get("explanation", []):
            para(f"- {line.replace('**', '')}")

    # Try to embed up to 3 chart images
    if charts:
        section("9. Charts")
        embedded = 0
        for title, fig in charts[:3]:
            try:
                img = fig.to_image(format="png", width=1000, height=560, scale=1)
                pdf.image(io.BytesIO(img), w=180)
                pdf.ln(2)
                embedded += 1
            except Exception:
                continue
        if embedded == 0:
            para("(Chart images could not be embedded — install `kaleido` "
                 "for chart rendering in PDF.)")

    try:
        return bytes(pdf.output(dest="S"))
    except Exception:
        return bytes(pdf.output())


# =============================================================================
# 12. POWERPOINT REPORT
# =============================================================================
def build_pptx_report(kpis, summary, insights, rca, dataset_type,
                      charts: list) -> bytes:
    if not HAS_PPTX:
        raise RuntimeError("python-pptx not installed — cannot build PPTX.")

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    def add_title_slide(title, subtitle):
        s = prs.slides.add_slide(prs.slide_layouts[0])
        s.shapes.title.text = title
        if len(s.placeholders) > 1:
            s.placeholders[1].text = subtitle

    def add_bullet_slide(title, bullets):
        s = prs.slides.add_slide(prs.slide_layouts[1])
        s.shapes.title.text = title
        tf = s.placeholders[1].text_frame
        tf.clear()
        for i, b in enumerate(bullets[:8]):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = b.replace("**", "")[:300]
            p.font.size = Pt(16)

    def add_image_slide(title, fig):
        try:
            img = fig.to_image(format="png", width=1200, height=650, scale=1)
            s = prs.slides.add_slide(prs.slide_layouts[5])
            s.shapes.title.text = title
            s.shapes.add_picture(
                io.BytesIO(img),
                Inches(0.6), Inches(1.4),
                width=Inches(12.1),
            )
        except Exception:
            add_bullet_slide(title, ["(Chart image unavailable — install kaleido)"])

    add_title_slide(
        "🚀 Business Copilot AI V3",
        f"Management Report — {datetime.now():%d %b %Y}\nDataset type: {dataset_type}",
    )

    add_bullet_slide(
        "Executive Summary",
        [summary.replace("**", "")],
    )

    add_bullet_slide(
        "KPI Summary",
        [f"{k}: {v}" for k, v in kpis.items() if not str(k).startswith("_")],
    )

    add_bullet_slide("What Happened", insights["happened"])
    add_bullet_slide("Risks", insights["risks"])
    add_bullet_slide("Opportunities", insights["opportunities"])
    add_bullet_slide("Recommendations", insights["recommendations"])

    if rca.get("ok"):
        add_bullet_slide("Root Cause Analysis",
                         [x.replace("**", "") for x in rca.get("explanation", [])])

    for title, fig in charts[:4]:
        add_image_slide(title, fig)

    add_bullet_slide("Next Steps", [
        "Validate AI insights with business owners.",
        "Address data quality gaps identified.",
        "Set weekly KPI review cadence.",
        "Re-run the analysis after the next data refresh.",
    ])

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


# =============================================================================
# 13. EMAIL SUMMARY
# =============================================================================
def build_email_summary(kpis, summary, insights) -> str:
    lines = []
    lines.append("Subject: Executive Brief — Business Copilot AI V3\n")
    lines.append("=" * 60 + "\n")
    lines.append("EXECUTIVE OVERVIEW")
    lines.append("-" * 60)
    lines.append(summary.replace("**", "") + "\n")
    lines.append("KEY KPIs")
    lines.append("-" * 60)
    for k, v in kpis.items():
        if str(k).startswith("_"):
            continue
        lines.append(f"• {k}: {v}")
    lines.append("\nWHAT HAPPENED")
    lines.append("-" * 60)
    for x in insights["happened"]:
        lines.append(f"• {x.replace('**', '')}")
    lines.append("\nRISKS")
    lines.append("-" * 60)
    for x in insights["risks"]:
        lines.append(f"• {x.replace('**', '')}")
    lines.append("\nOPPORTUNITIES")
    lines.append("-" * 60)
    for x in insights["opportunities"]:
        lines.append(f"• {x.replace('**', '')}")
    lines.append("\nRECOMMENDATIONS")
    lines.append("-" * 60)
    for x in insights["recommendations"]:
        lines.append(f"• {x.replace('**', '')}")
    lines.append("\n— Generated by Business Copilot AI V3")
    return "\n".join(lines)


# =============================================================================
# 14. HERO + FILE UPLOAD + SESSION STATE
# =============================================================================
st.markdown(
    """
<div class="hero">
  <h1>🚀 Business Copilot AI V3</h1>
  <h3>AI-Powered Business Intelligence Platform</h3>
  <ul>
    <li>✅ Auto Dataset Detection</li>
    <li>✅ Executive KPIs & Summary</li>
    <li>✅ AI Insights (What Happened · Risks · Opportunities · Recommendations)</li>
    <li>✅ Root Cause Analysis</li>
    <li>✅ Chat with your data (Ollama / OpenAI)</li>
    <li>✅ PDF · PowerPoint · Email Reports</li>
  </ul>
</div>
""",
    unsafe_allow_html=True,
)

if "chat_messages" not in st.session_state:
    st.session_state.chat_messages = []
if "current_file_hash" not in st.session_state:
    st.session_state.current_file_hash = None

uploaded = st.file_uploader(
    "📁 Upload your Excel or CSV file",
    type=["csv", "xlsx", "xls"],
    help="Any CSV / Excel — Sales, Survey, Finance, HR, Operations, Inventory, Marketing…",
)

if uploaded is None:
    st.info("👆 Upload a file to activate your AI business dashboard.")
    st.stop()

# ---- Load ----
try:
    file_bytes = uploaded.getvalue()
    df = load_data(file_bytes, uploaded.name)
except Exception as e:
    st.error(f"❌ Could not read file: {e}")
    st.stop()

if df is None or df.empty:
    st.warning("⚠️ File contains no data rows.")
    st.stop()

# ---- Reset chat on new file ----
file_hash = hashlib.md5(file_bytes).hexdigest()
if st.session_state.current_file_hash != file_hash:
    st.session_state.current_file_hash = file_hash
    st.session_state.chat_messages = []

# ---- Analysis ----
date_cols = detect_date_columns(df)
num_cols = detect_numeric_columns(df)
cat_cols = detect_categorical_columns(df, exclude=tuple(date_cols))
dataset_type = detect_dataset_type(df)
kpis = compute_kpis(df, dataset_type, num_cols, cat_cols, date_cols)
summary = generate_executive_summary(df, dataset_type, kpis, cat_cols, date_cols)
insights = generate_ai_insights(df, num_cols, cat_cols, date_cols, kpis)

# ---- Sidebar filters ----
with st.sidebar:
    st.markdown("## 🎛️ Global Filters")
    st.caption("Filters apply to all tabs below.")

filtered_df = df.copy()

with st.sidebar:
    # Date
    if date_cols:
        dc = date_cols[0]
        try:
            dates = pd.to_datetime(df[dc], errors="coerce").dropna()
            if len(dates) > 1:
                min_d = dates.min().date()
                max_d = dates.max().date()
                if min_d < max_d:
                    dr = st.date_input("📅 Date Range",
                                       value=(min_d, max_d),
                                       min_value=min_d, max_value=max_d)
                    if isinstance(dr, tuple) and len(dr) == 2:
                        s, e = dr
                        col_dates = pd.to_datetime(filtered_df[dc], errors="coerce").dt.date
                        mask = col_dates.between(s, e) | filtered_df[dc].isna()
                        filtered_df = filtered_df[mask]
        except Exception:
            pass

    # Categorical
    for c in cat_cols[:5]:
        try:
            uniques = filtered_df[c].dropna().astype(str).unique().tolist()
            if 1 < len(uniques) <= 100:
                uniques = sorted(uniques)
                chosen = st.multiselect(f"{c}", uniques, default=[])
                if chosen:
                    filtered_df = filtered_df[
                        filtered_df[c].astype(str).isin(chosen)
                    ]
        except Exception:
            continue

    # Numeric range
    for c in num_cols[:2]:
        try:
            s = pd.to_numeric(filtered_df[c], errors="coerce").dropna()
            if len(s) > 0 and s.min() < s.max():
                lo, hi = float(s.min()), float(s.max())
                rng = st.slider(f"📊 {c}", lo, hi, (lo, hi))
                numeric_col = pd.to_numeric(filtered_df[c], errors="coerce")
                filtered_df = filtered_df[
                    numeric_col.between(rng[0], rng[1]) | filtered_df[c].isna()
                ]
        except Exception:
            continue

    st.markdown("---")
    st.caption(f"Showing **{len(filtered_df):,}** / **{len(df):,}** rows")

    # ---- AI chat settings ----
    st.markdown("## 🔌 AI Chat Settings")
    provider = st.selectbox(
        "Provider (auto-fallback: OpenAI if key set)",
        ["Ollama (local)", "OpenAI"],
    )
    ollama_url = st.text_input("Ollama URL", value="http://localhost:11434")
    ollama_model = st.selectbox(
        "Ollama model",
        ["qwen2.5-coder", "llama3", "mistral", "deepseek-coder", "qwen2.5"],
    )
    openai_key = st.text_input("OpenAI API Key (optional)", type="password")
    openai_model = st.selectbox(
        "OpenAI model", ["gpt-4o-mini", "gpt-4o", "gpt-3.5-turbo"]
    )

# Recompute derived insights against filtered_df for tabs (but keep base for chat context)
filtered_summary = generate_executive_summary(
    filtered_df, dataset_type, kpis, cat_cols, date_cols
) if len(filtered_df) > 0 else summary
filtered_insights = generate_ai_insights(
    filtered_df, num_cols, cat_cols, date_cols, kpis
) if len(filtered_df) > 0 else insights

# =============================================================================
# 15. TABS
# =============================================================================
tab_dash, tab_ai, tab_rca, tab_chat, tab_reports, tab_explore = st.tabs([
    "🏠 Executive Dashboard",
    "🤖 AI Insights",
    "🔎 Root Cause Analysis",
    "💬 Ask Your Data",
    "📄 Reports",
    "🔍 Data Explorer",
])

# -----------------------------------------------------------------------------
# TAB 1 — EXECUTIVE DASHBOARD
# -----------------------------------------------------------------------------
with tab_dash:
    st.markdown(f"### 📌 Key Performance Indicators "
                f"— *{dataset_type} Dataset*")

    k1, k2, k3, k4 = st.columns(4)
    k1.markdown(f"""<div class="kpi-card">
        <div class="kpi-label">Records</div>
        <div class="kpi-value">{kpis['Records']:,}</div>
        <div class="kpi-hint">Total rows in dataset</div></div>""",
        unsafe_allow_html=True)
    k2.markdown(f"""<div class="kpi-card">
        <div class="kpi-label">Columns</div>
        <div class="kpi-value">{kpis['Columns']}</div>
        <div class="kpi-hint">Fields detected</div></div>""",
        unsafe_allow_html=True)
    k3.markdown(f"""<div class="kpi-card">
        <div class="kpi-label">Missing Data</div>
        <div class="kpi-value">{kpis['Missing %']}</div>
        <div class="kpi-hint">Overall missingness</div></div>""",
        unsafe_allow_html=True)
    k4.markdown(f"""<div class="kpi-card">
        <div class="kpi-label">Duplicate Rows</div>
        <div class="kpi-value">{kpis['Duplicate Rows']:,}</div>
        <div class="kpi-hint">Repeated records</div></div>""",
        unsafe_allow_html=True)

    k5, k6, k7, k8 = st.columns(4)
    k5.markdown(f"""<div class="kpi-card">
        <div class="kpi-label">Date Range</div>
        <div class="kpi-value">{kpis['Date Range']}</div>
        <div class="kpi-hint">Timeline span</div></div>""",
        unsafe_allow_html=True)
    k6.markdown(f"""<div class="kpi-card">
        <div class="kpi-label">Top Category</div>
        <div class="kpi-value">{kpis['Top Category']}</div>
        <div class="kpi-hint">Leading segment</div></div>""",
        unsafe_allow_html=True)
    k7.markdown(f"""<div class="kpi-card">
        <div class="kpi-label">Primary Metric</div>
        <div class="kpi-value">{kpis['Primary Metric']}</div>
        <div class="kpi-hint">Total: {kpis['Primary Value']}</div></div>""",
        unsafe_allow_html=True)
    k8.markdown(f"""<div class="kpi-card">
        <div class="kpi-label">Data Quality</div>
        <div class="kpi-value">{kpis['Data Quality']}</div>
        <div class="kpi-hint">Growth: {kpis['Growth %']}</div></div>""",
        unsafe_allow_html=True)

    st.markdown("---")

    # ---- Executive Summary ----
    st.markdown("### 📝 Executive Summary")
    st.markdown(
        f'<div class="section-card"><p>{filtered_summary}</p></div>',
        unsafe_allow_html=True,
    )

    st.markdown("---")

    # ---- Auto Charts ----
    st.markdown("### 📊 Auto-Generated Charts")
    charts = build_auto_charts(filtered_df, cat_cols, num_cols, date_cols,
                                kpis.get("_metric"))

    if not charts:
        st.info("ℹ️ Not enough data structure to auto-generate charts.")
    else:
        # First row: 2 charts
        for i in range(0, len(charts), 2):
            row = charts[i:i+2]
            cols = st.columns(len(row))
            for col, (title, fig) in zip(cols, row):
                with col:
                    st.plotly_chart(fig, use_container_width=True)

    st.markdown("---")

    # ---- Quick insight preview ----
    st.markdown("### 💡 Top Insights Preview")
    for x in filtered_insights["happened"][:3]:
        st.info(x)
    for x in filtered_insights["risks"][:2]:
        st.warning(x)

# -----------------------------------------------------------------------------
# TAB 2 — AI INSIGHTS
# -----------------------------------------------------------------------------
with tab_ai:
    st.markdown("### 🤖 AI-Generated Business Insights")
    st.caption("Automatically derived from the uploaded dataset.")

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("#### 📈 What Happened")
        for x in filtered_insights["happened"]:
            st.markdown(f'<div class="chip-info">{x}</div>',
                        unsafe_allow_html=True)

        st.markdown("#### ⚠️ Risks")
        for x in filtered_insights["risks"]:
            st.markdown(f'<div class="chip-warn">{x}</div>',
                        unsafe_allow_html=True)

    with col2:
        st.markdown("#### 🚀 Opportunities")
        for x in filtered_insights["opportunities"]:
            st.markdown(f'<div class="chip-success">{x}</div>',
                        unsafe_allow_html=True)

        st.markdown("#### 🎯 Recommendations")
        for x in filtered_insights["recommendations"]:
            st.markdown(f'<div class="chip-danger">{x}</div>',
                        unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("### 🩺 Data Health Snapshot")
    h1, h2, h3, h4 = st.columns(4)
    h1.metric("Numeric Columns", len(num_cols))
    h2.metric("Categorical Columns", len(cat_cols))
    h3.metric("Date Columns", len(date_cols))
    h4.metric("Missing %", kpis["Missing %"])

# -----------------------------------------------------------------------------
# TAB 3 — ROOT CAUSE ANALYSIS
# -----------------------------------------------------------------------------
with tab_rca:
    st.markdown("### 🔎 Root Cause Analysis")
    st.caption("Identifies which segments drove the change in the primary metric.")

    if st.button("🔬 Analyze Root Cause", type="primary"):
        st.session_state.rca_result = run_root_cause_analysis(
            filtered_df, num_cols, cat_cols, date_cols
        )

    rca = st.session_state.get("rca_result",
                               {"ok": False, "message": "Click the button above to run analysis."})

    if not rca.get("ok"):
        st.warning(rca.get("message", "Root cause analysis could not be run."))
    else:
        st.success(
            f"**Primary metric:** `{rca['metric']}` · "
            f"**Earlier:** {format_number(rca['early_total'])} · "
            f"**Later:** {format_number(rca['late_total'])} · "
            f"**Change:** {rca['change_pct']:+.1f}%"
        )

        st.markdown("#### 📝 Explanation")
        for line in rca["explanation"]:
            st.markdown(f'<div class="chip-info">{line}</div>',
                        unsafe_allow_html=True)

        st.markdown("#### 📊 Driver Breakdown by Dimension")
        for d in rca["drivers"]:
            with st.expander(f"Dimension: **{d['dimension']}**"):
                try:
                    diff = d["all"].reset_index()
                    diff.columns = [d["dimension"], "Change"]
                    diff["Change"] = diff["Change"].round(2)
                    fig = px.bar(
                        diff.sort_values("Change"),
                        x="Change", y=d["dimension"],
                        orientation="h",
                        color="Change",
                        color_continuous_scale="RdYlGn",
                        title=f"Change contribution by {d['dimension']}",
                    )
                    fig.update_layout(**PLOTLY_LAYOUT,
                                      coloraxis_showscale=False,
                                      height=360)
                    st.plotly_chart(fig, use_container_width=True)
                except Exception as e:
                    st.caption(f"Could not render chart: {e}")

# -----------------------------------------------------------------------------
# TAB 4 — AI CHAT
# -----------------------------------------------------------------------------
with tab_chat:
    st.markdown("### 💬 Ask Your Data")
    st.caption(
        "The assistant understands your uploaded dataset, KPIs, and insights. "
        "Uses **OpenAI** if a key is provided, otherwise **Ollama** (local)."
    )

    # Chat history
    for msg in st.session_state.chat_messages:
        cls = "chat-user" if msg["role"] == "user" else "chat-bot"
        st.markdown(f'<div class="{cls}">{msg["content"]}</div>',
                    unsafe_allow_html=True)

    # Suggested queries
    st.markdown("**Try one of these:**")
    sugg_cols = st.columns(4)
    suggestions = [
        "Explain the dataset.",
        "What is the biggest risk?",
        "Who is the top contributor?",
        "Generate an executive summary.",
    ]
    for col, s in zip(sugg_cols, suggestions):
        if col.button(s, key=f"sugg_{s}"):
            st.session_state._pending_question = s

    user_q = st.text_area("Ask a question about the data:", height=80,
                          key="chat_input_box",
                          value=st.session_state.pop("_pending_question", ""))

    c1, c2 = st.columns([1, 5])
    send = c1.button("🚀 Send", type="primary", use_container_width=True)
    clear = c2.button("🧹 Clear Chat", use_container_width=False)

    if clear:
        st.session_state.chat_messages = []
        st.rerun()

    if send and user_q.strip():
        st.session_state.chat_messages.append(
            {"role": "user", "content": user_q.strip()}
        )

        context = build_chat_context(
            df, dataset_type, kpis, summary, insights
        )
        full_prompt = (
            context
            + "\n\nUSER QUESTION:\n"
            + user_q.strip()
            + "\n\nAnswer clearly and concisely."
        )

        with st.spinner("Thinking…"):
            if provider == "OpenAI" and openai_key.strip():
                answer = call_openai(full_prompt, openai_key.strip(),
                                     openai_model)
            elif provider == "OpenAI" and not openai_key.strip():
                answer = ("⚠️ OpenAI selected but no API key provided — "
                          "falling back to Ollama.\n\n"
                          + call_ollama(full_prompt, ollama_model, ollama_url))
            else:
                answer = call_ollama(full_prompt, ollama_model, ollama_url)

        st.session_state.chat_messages.append(
            {"role": "assistant", "content": answer}
        )
        st.rerun()

# -----------------------------------------------------------------------------
# TAB 5 — REPORTS
# -----------------------------------------------------------------------------
with tab_reports:
    st.markdown("### 📄 Management Report Generator")
    st.caption("Generate PDF, PowerPoint, and Email versions of the full report.")

    rca = st.session_state.get(
        "rca_result", {"ok": False, "message": "Not run"}
    )

    # Regenerate charts specifically for reports (uses full df for stability)
    report_charts = build_auto_charts(df, cat_cols, num_cols, date_cols,
                                       kpis.get("_metric"))

    # ---- PDF ----
    st.markdown("#### 📕 PDF Report")
    if HAS_FPDF:
        try:
            pdf_bytes = build_pdf_report(
                kpis, summary, insights, rca, dataset_type,
                df.columns, report_charts,
            )
            st.download_button(
                "⬇️ Download PDF Report",
                data=pdf_bytes,
                file_name=f"business_report_{datetime.now():%Y%m%d_%H%M}.pdf",
                mime="application/pdf",
                use_container_width=True,
            )
        except Exception as e:
            st.warning(f"PDF generation failed: {e}")
    else:
        st.warning("Install `fpdf2` to enable PDF reports.")

    st.markdown("---")

    # ---- PPTX ----
    st.markdown("#### 📙 PowerPoint Report")
    if HAS_PPTX:
        try:
            pptx_bytes = build_pptx_report(
                kpis, summary, insights, rca, dataset_type, report_charts,
            )
            st.download_button(
                "⬇️ Download PowerPoint",
                data=pptx_bytes,
                file_name=f"business_report_{datetime.now():%Y%m%d_%H%M}.pptx",
                mime=("application/vnd.openxmlformats-officedocument"
                      ".presentationml.presentation"),
                use_container_width=True,
            )
        except Exception as e:
            st.warning(f"PowerPoint generation failed: {e}")
    else:
        st.warning("Install `python-pptx` to enable PowerPoint reports.")

    st.markdown("---")

    # ---- Email ----
    st.markdown("#### ✉️ Email Summary")
    email_txt = build_email_summary(kpis, summary, insights)
    st.text_area("Copy this into Outlook / Gmail:", value=email_txt,
                 height=340)
    st.download_button(
        "⬇️ Download Email (.txt)",
        data=email_txt.encode("utf-8"),
        file_name=f"executive_email_{datetime.now():%Y%m%d}.txt",
        mime="text/plain",
        use_container_width=True,
    )

    st.markdown("---")

    # ---- Data exports ----
    st.markdown("#### 📦 Data Exports")
    ec1, ec2 = st.columns(2)

    csv_bytes = filtered_df.to_csv(index=False).encode("utf-8")
    ec1.download_button(
        "⬇️ Download Filtered CSV",
        data=csv_bytes,
        file_name="business_copilot_filtered.csv",
        mime="text/csv",
        use_container_width=True,
    )

    try:
        excel_buf = io.BytesIO()
        with pd.ExcelWriter(excel_buf, engine="openpyxl") as writer:
            filtered_df.to_excel(writer, index=False, sheet_name="Data")
        ec2.download_button(
            "⬇️ Download Filtered Excel",
            data=excel_buf.getvalue(),
            file_name="business_copilot_filtered.xlsx",
            mime=("application/vnd.openxmlformats-officedocument"
                  ".spreadsheetml.sheet"),
            use_container_width=True,
        )
    except Exception as e:
        ec2.warning(f"Excel export failed: {e}")

# -----------------------------------------------------------------------------
# TAB 6 — DATA EXPLORER
# -----------------------------------------------------------------------------
with tab_explore:
    st.markdown("### 🔍 Explore Your Data")

    search_q = st.text_input("🔎 Search across all text columns",
                             placeholder="Type to filter rows…")

    text_cols = [
        c for c in filtered_df.columns
        if not pd.api.types.is_numeric_dtype(filtered_df[c])
        and not pd.api.types.is_datetime64_any_dtype(filtered_df[c])
    ]

    view = filtered_df.copy()
    if search_q.strip() and text_cols:
        mask = pd.Series(False, index=view.index)
        q = search_q
        for c in text_cols:
            mask |= view[c].astype(str).str.contains(
                q, case=False, na=False, regex=False
            )
        view = view[mask]

    st.caption(f"Showing **{len(view):,}** / **{len(filtered_df):,}** rows")
    st.dataframe(view, use_container_width=True, height=460)

    st.markdown("---")
    st.markdown("### 🧬 Column Analysis")

    rows = []
    for c in df.columns:
        s = df[c]
        try:
            sample = s.dropna().astype(str).unique()[:2]
            sample_str = ", ".join(sample) if len(sample) else "—"
        except Exception:
            sample_str = "—"
        rows.append({
            "Column": c,
            "DType": str(s.dtype),
            "Non-Null": int(s.notna().sum()),
            "Missing": int(s.isna().sum()),
            "Missing %": round(float(s.isna().mean()) * 100, 2),
            "Unique": int(s.nunique(dropna=True)),
            "Sample": sample_str[:60],
        })

    st.dataframe(pd.DataFrame(rows), use_container_width=True, height=420)