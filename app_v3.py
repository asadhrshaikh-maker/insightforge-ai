"""
🚀 Business Copilot AI V3 — Enterprise AI-Powered Business Intelligence
=======================================================================
Single-file, production-ready Streamlit application.

Run:
    python -m streamlit run app_v3.py

Stack: Streamlit · Pandas · NumPy · Plotly · OpenPyXL · python-pptx ·
       Matplotlib (report chart images) · Requests (Ollama / OpenAI)

PDF engine: built-in pure-Python writer (zero extra dependencies).
"""

from __future__ import annotations

import html
import io
import json
import re
import struct
import warnings
import zlib
from datetime import datetime

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import plotly.io as pio
import requests
import streamlit as st

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

# =====================================================================
# BRAND / THEME CONSTANTS
# =====================================================================
APP_NAME = "Business Copilot AI"
APP_VERSION = "V3"
BLUE, PURPLE, GREEN, ORANGE, RED = "#3B82F6", "#8B5CF6", "#10B981", "#F59E0B", "#EF4444"
CYAN, PINK, LIME = "#06B6D4", "#EC4899", "#84CC16"
BG0, BG1, BG2 = "#050B18", "#071326", "#0B2344"
TEXT, MUTED = "#E5EAF5", "#94A3B8"
PALETTE = [BLUE, PURPLE, GREEN, ORANGE, RED, CYAN, PINK, LIME]
MAX_EXCEL_ROWS = 1_048_575

st.set_page_config(page_title=f"{APP_NAME} {APP_VERSION}", page_icon="🚀",
                   layout="wide", initial_sidebar_state="expanded")

CUSTOM_CSS = """
<style>
#MainMenu {visibility: hidden;} footer {visibility: hidden;}
[data-testid="stAppViewContainer"] {
  background: radial-gradient(1200px 600px at 15% -10%, #0B2344 0%, #071326 45%, #050B18 100%) fixed;
}
[data-testid="stSidebar"] {background: linear-gradient(180deg, #071326 0%, #050B18 100%);
  border-right: 1px solid rgba(59,130,246,0.20);}
[data-testid="stHeader"] {background: rgba(5,11,24,0.6); backdrop-filter: blur(8px);}
.block-container {padding-top: 1.2rem; padding-bottom: 3rem; max-width: 1500px;}
h1, h2, h3 {color: #E5EAF5; letter-spacing: -0.01em;}
.hero-title {font-size: 2.6rem; font-weight: 900; line-height: 1.1;
  background: linear-gradient(90deg, #3B82F6, #8B5CF6, #06B6D4);
  -webkit-background-clip: text; background-clip: text; color: transparent;}
.hero-sub {color: #94A3B8; font-size: 1.05rem; margin-top: 6px;}
.chip {display: inline-block; padding: 4px 12px; margin: 10px 8px 0 0; border-radius: 999px;
  font-size: 0.78rem; font-weight: 600; color: #CFE0FF;
  background: rgba(59,130,246,0.14); border: 1px solid rgba(59,130,246,0.35);}
.badge {display: inline-block; padding: 6px 16px; border-radius: 999px; font-weight: 700;
  font-size: 0.85rem; color: #fff; background: linear-gradient(90deg, #3B82F6, #8B5CF6);
  box-shadow: 0 4px 18px rgba(59,130,246,0.35);}
.glass {background: rgba(11,35,68,0.45); backdrop-filter: blur(14px);
  border: 1px solid rgba(59,130,246,0.22); border-radius: 18px;
  box-shadow: 0 8px 30px rgba(0,0,0,0.35); padding: 20px 24px; margin: 8px 0 16px 0;}
.kpi {background: rgba(11,35,68,0.55); backdrop-filter: blur(14px);
  border: 1px solid rgba(59,130,246,0.22); border-radius: 16px;
  box-shadow: 0 8px 24px rgba(0,0,0,0.35); padding: 14px 14px; text-align: center; height: 100%;}
.kpi-icon {font-size: 1.25rem;}
.kpi-label {font-size: 0.68rem; font-weight: 700; letter-spacing: .09em;
  text-transform: uppercase; color: #94A3B8; margin-top: 4px;}
.kpi-value {font-size: 1.32rem; font-weight: 800; margin-top: 2px; color: #E5EAF5;
  word-break: break-word;}
.kpi-value.pos {color: #10B981;} .kpi-value.neg {color: #EF4444;}
.kpi-sub {font-size: 0.68rem; color: #7C8DB0; margin-top: 2px;}
.section {display: flex; align-items: center; gap: 10px; margin: 18px 0 10px 0;}
.section-bar {width: 5px; height: 26px; border-radius: 3px;}
.section-title {font-size: 1.25rem; font-weight: 800; color: #E5EAF5;}
.insight {background: rgba(11,35,68,0.5); border: 1px solid rgba(148,163,184,0.18);
  border-left: 5px solid #3B82F6; border-radius: 14px; padding: 14px 16px; margin-bottom: 12px;
  box-shadow: 0 6px 18px rgba(0,0,0,0.3);}
.insight-title {font-weight: 800; font-size: 0.95rem; color: #E5EAF5; margin-bottom: 4px;}
.insight-body {font-size: 0.88rem; color: #B9C6DE; line-height: 1.5;}
.sev {float: right; font-size: 0.66rem; font-weight: 800; padding: 2px 10px; border-radius: 999px;
  text-transform: uppercase; letter-spacing: .06em;}
button[data-baseweb="tab"] {color: #94A3B8; font-weight: 700;}
button[data-baseweb="tab"][aria-selected="true"] {color: #E5EAF5;}
button[data-baseweb="tab-highlight"] {background: linear-gradient(90deg,#3B82F6,#8B5CF6);}
.stButton > button, .stDownloadButton > button {
  background: linear-gradient(90deg, #3B82F6, #8B5CF6); color: white; border: none;
  border-radius: 12px; font-weight: 700; padding: 0.5rem 1.2rem;
  box-shadow: 0 6px 18px rgba(59,130,246,0.35);}
.stButton > button:hover, .stDownloadButton > button:hover {filter: brightness(1.12); color: white;}
[data-testid="stChatMessage"] {background: rgba(11,35,68,0.45);
  border: 1px solid rgba(59,130,246,0.18); border-radius: 14px; padding: 10px 14px;}
hr {border-color: rgba(148,163,184,0.15);}
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

_TPL = go.layout.Template()
_TPL.layout = go.Layout(
    paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
    font=dict(color=TEXT, family="Segoe UI, Inter, Arial, sans-serif", size=13),
    colorway=PALETTE,
    title=dict(font=dict(size=17, color=TEXT), x=0.01, xanchor="left"),
    xaxis=dict(gridcolor="rgba(148,163,184,0.14)", zerolinecolor="rgba(148,163,184,0.28)"),
    yaxis=dict(gridcolor="rgba(148,163,184,0.14)", zerolinecolor="rgba(148,163,184,0.28)"),
    legend=dict(bgcolor="rgba(0,0,0,0)"),
    margin=dict(l=10, r=10, t=56, b=10),
    hoverlabel=dict(bgcolor=BG2, font_color=TEXT),
)
pio.templates["copilot_dark"] = _TPL
warnings.filterwarnings("ignore", category=FutureWarning, module="plotly")


def style_fig(fig: go.Figure, title: str | None = None, height: int = 400) -> go.Figure:
    fig.update_layout(template="copilot_dark", height=height)
    if title:
        fig.update_layout(title=title)
    return fig


# =====================================================================
# FORMATTING UTILITIES
# =====================================================================
def fmt_int(v) -> str:
    try:
        return f"{int(v):,}"
    except (TypeError, ValueError):
        return "—"


def fmt_compact(value, symbol: str = "₹") -> str:
    try:
        v = float(value)
    except (TypeError, ValueError):
        return "N/A"
    if not np.isfinite(v):
        return "N/A"
    sign = "-" if v < 0 else ""
    a = abs(v)
    for div, suf in ((1e9, "B"), (1e6, "M"), (1e3, "K")):
        if a >= div:
            return f"{sign}{symbol}{a / div:,.1f}{suf}"
    if a >= 1 or a == 0:
        return f"{sign}{symbol}{a:,.0f}"
    return f"{sign}{symbol}{a:,.2f}"


def fmt_pct(v, digits: int = 1) -> str:
    if v is None:
        return "—"
    try:
        f = float(v)
    except (TypeError, ValueError):
        return "—"
    if not np.isfinite(f):
        return "—"
    return f"{f:+.{digits}f}%" if f else f"{f:.{digits}f}%"


def esc(v) -> str:
    return html.escape(str(v))


def clip(text: str, n: int = 200) -> str:
    text = str(text)
    return text if len(text) <= n else text[: n - 1] + "…"


# =====================================================================
# DATA LOADING & PREPARATION
# =====================================================================
@st.cache_data(show_spinner="📂 Loading dataset…")
def load_data(file_bytes: bytes, file_name: str, sheet_name: str | None = None) -> pd.DataFrame:
    if file_name.lower().endswith(".csv"):
        last_error: Exception | None = None
        for encoding in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
            try:
                return pd.read_csv(io.BytesIO(file_bytes), encoding=encoding)
            except UnicodeDecodeError as err:
                last_error = err
        raise ValueError(f"Could not decode CSV file: {last_error}")
    return pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name or 0)


@st.cache_data(show_spinner=False)
def get_sheet_names(file_bytes: bytes, file_name: str) -> list[str]:
    if file_name.lower().endswith(".csv"):
        return []
    try:
        with pd.ExcelFile(io.BytesIO(file_bytes)) as xls:
            return list(xls.sheet_names)
    except Exception:
        return []


STRONG_ID_TOKENS = {"id", "ids", "uuid", "guid", "index", "unnamed", "phone", "mobile",
                    "ssn", "zip", "zipcode", "pin", "pincode", "postal", "account",
                    "pan", "aadhaar", "email", "sku"}
WEAK_ID_TOKENS = {"code", "number", "no", "year", "yr"}
NON_DATE_TOKENS = {"version", "ver", "ip", "mac", "url", "email", "address"}
CATEGORY_PRIORITY = ("region", "category", "product", "segment", "status", "department",
                     "channel", "campaign", "city", "country", "state", "location",
                     "project", "warehouse", "plant", "shift", "type", "brand", "team",
                     "group", "gender", "interviewer")
NUMERIC_PRIORITY = ("revenue", "sales", "amount", "profit", "income", "total", "actual",
                    "spend", "stock", "price", "expense", "cost", "budget", "salary", "quantity",
                    "qty", "units", "value", "conversions", "clicks", "impressions",
                    "score", "satisfaction")
DATE_PRIORITY = ("date", "month", "time", "day", "created", "timestamp", "snapshot")


def _name_tokens(name) -> list[str]:
    return [t for t in re.split(r"[^a-z0-9]+", str(name).lower()) if t]


def _is_identifier(name) -> bool:
    tokens = _name_tokens(name)
    if "of" in tokens:
        return False
    return any(t in STRONG_ID_TOKENS for t in tokens)


def _is_metric_name(name) -> bool:
    tokens = _name_tokens(name)
    if "of" in tokens:
        return True
    return not any(t in STRONG_ID_TOKENS | WEAK_ID_TOKENS for t in tokens)


def _priority(name, keywords) -> int:
    lname = str(name).lower()
    for i, kw in enumerate(keywords):
        if kw in lname:
            return i
    return len(keywords)


def _looks_like_dates(text: pd.Series) -> bool:
    sample = text.head(200)
    if sample.empty:
        return False
    sep = sample.str.contains(r"[-/\.]|jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec",
                              case=False, regex=True)
    if float(sep.mean()) < 0.8:
        return False
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        parsed = pd.to_datetime(sample, errors="coerce")
    ok = parsed.notna()
    if float(ok.mean()) < 0.8:
        return False
    return bool(parsed.dt.year[ok].between(1900, 2100).mean() >= 0.9)


def prepare_df(df: pd.DataFrame) -> pd.DataFrame:
    """Clean headers, drop empty rows/cols, infer numeric & date types."""
    df = df.copy()
    new_cols, seen = [], {}
    for i, col in enumerate(df.columns):
        name = str(col).strip() or f"Column_{i + 1}"
        if name in seen:
            seen[name] += 1
            name = f"{name}_{seen[name]}"
        else:
            seen[name] = 0
        new_cols.append(name)
    df.columns = new_cols
    unnamed = [c for c in df.columns
               if c.lower().startswith("unnamed") and df[c].isna().all()]
    if unnamed:
        df = df.drop(columns=unnamed)
    df = df.dropna(axis=0, how="all").dropna(axis=1, how="all").reset_index(drop=True)
    if df.empty:
        return df
    for col in df.columns:
        s = df[col]
        if not (pd.api.types.is_object_dtype(s) or pd.api.types.is_string_dtype(s)):
            continue
        if _is_identifier(col):
            continue
        non_null = s.dropna()
        if non_null.empty:
            continue
        text = non_null.astype(str).str.strip()
        stripped = text.str.replace(r"[₹$€£,%\s]", "", regex=True)
        numeric = pd.to_numeric(stripped, errors="coerce")
        if (float(numeric.notna().mean()) >= 0.9
                and bool(stripped.str.match(r"^[+-]?[\d.]").any())):
            df[col] = pd.to_numeric(
                s.astype(str).str.strip().str.replace(r"[₹$€£,%\s]", "", regex=True),
                errors="coerce")
            continue
        if any(t in NON_DATE_TOKENS for t in _name_tokens(col)):
            continue
        if _looks_like_dates(text):
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                df[col] = pd.to_datetime(s, errors="coerce")
    for col in df.columns:
        s = df[col]
        if (pd.api.types.is_numeric_dtype(s) and not pd.api.types.is_bool_dtype(s)
                and "year" in _name_tokens(col)):
            vals = s.dropna()
            if (not vals.empty and bool(vals.between(1900, 2100).all())
                    and bool((vals % 1 == 0).all())):
                df[col] = pd.to_datetime(s.astype("Int64").astype(str),
                                         format="%Y", errors="coerce")
    return df


# =====================================================================
# SAMPLE DATA GENERATOR (7 business domains, deterministic)
# =====================================================================
SAMPLE_KINDS = ["Sales", "HR", "Finance", "Survey", "Operations", "Inventory", "Marketing"]


@st.cache_data(show_spinner=False)
def make_sample(kind: str = "Sales") -> pd.DataFrame:
    rng = np.random.default_rng(42)
    n = 1500
    months = pd.date_range("2025-01-01", "2026-10-01", freq="MS")
    dates = pd.to_datetime(rng.choice(months, n)) + pd.to_timedelta(
        rng.integers(0, 27, n), unit="D")
    month_idx = (dates.year - 2025) * 12 + dates.month
    trend = 1 + 0.015 * month_idx
    base = rng.lognormal(mean=9.6, sigma=0.55, size=n)

    if kind == "Sales":
        region = rng.choice(["North", "South", "East", "West"], n, p=[.38, .27, .21, .14])
        product = rng.choice(["Alpha CRM", "Beta Analytics", "Cloud Suite", "Support Plan"],
                             n, p=[.36, .27, .22, .15])
        channel = rng.choice(["Online", "Partner", "Direct"], n, p=[.5, .3, .2])
        rm = pd.Series(region).map({"North": 1.3, "South": 1.0, "East": .9, "West": .75}).to_numpy()
        last = month_idx >= 22
        shock = np.where(last & (pd.Series(region) == "East").to_numpy(), .30, 1.0)
        shock = shock * np.where(last & (pd.Series(region) == "West").to_numpy(), .45, 1.0)
        shock = shock * np.where(last & (pd.Series(region) == "South").to_numpy(), .72, 1.0)
        revenue = np.round(base * rm * trend * shock, 2)
        df = pd.DataFrame({"Order Date": dates, "Region": region, "Product": product,
                           "Channel": channel, "Revenue": revenue,
                           "Profit": np.round(revenue * rng.uniform(.08, .3, n), 2),
                           "Units": rng.integers(1, 40, n),
                           "Customer ID": [f"CUST-{i:05d}" for i in range(n)]})
    elif kind == "HR":
        dept = rng.choice(["Engineering", "Sales", "Marketing", "Finance", "Operations"],
                          n, p=[.32, .24, .16, .14, .14])
        dm = pd.Series(dept).map({"Engineering": 1.5, "Sales": 1.1, "Marketing": 1.0,
                                  "Finance": 1.15, "Operations": .85}).to_numpy()
        df = pd.DataFrame({"Hire Date": dates, "Department": dept,
                           "Gender": rng.choice(["Female", "Male"], n, p=[.46, .54]),
                           "Location": rng.choice(["Bengaluru", "Pune", "Delhi", "Remote"], n),
                           "Salary": np.round(base * dm * 28, 0),
                           "Performance Score": np.round(rng.uniform(2.2, 5.0, n), 2),
                           "Attrition": rng.choice(["Yes", "No"], n, p=[.14, .86]),
                           "Employee ID": [f"EMP-{i:05d}" for i in range(n)]})
    elif kind == "Finance":
        dept = rng.choice(["Engineering", "Sales", "Marketing", "Finance", "Operations"], n)
        cat = rng.choice(["Travel", "Software", "Payroll", "Infrastructure", "Events"], n)
        budget = np.round(base * 20, 2)
        actual = np.round(budget * rng.uniform(.7, 1.35, n) * trend, 2)
        df = pd.DataFrame({"Month": dates, "Department": dept, "Expense Category": cat,
                           "Budget": budget, "Actual": actual,
                           "Variance": np.round(actual - budget, 2)})
    elif kind == "Survey":
        df = pd.DataFrame({"Survey Date": dates,
                           "Project": rng.choice(["Alpha Rollout", "Beta Pilot", "Gamma Study"], n),
                           "Interviewer": rng.choice(["Team A", "Team B", "Team C", "Team D"], n),
                           "Region": rng.choice(["North", "South", "East", "West"], n),
                           "Status": rng.choice(["Completed", "Pending", "Refused"], n, p=[.78, .16, .06]),
                           "Satisfaction": np.round(rng.uniform(1, 10, n), 1),
                           "Responses": rng.integers(5, 60, n)})
    elif kind == "Operations":
        df = pd.DataFrame({"Production Date": dates,
                           "Plant": rng.choice(["Plant 1", "Plant 2", "Plant 3"], n),
                           "Shift": rng.choice(["Morning", "Evening", "Night"], n),
                           "Location": rng.choice(["Chennai", "Pune", "Ahmedabad"], n),
                           "Status": rng.choice(["On Time", "Delayed", "Halted"], n, p=[.8, .15, .05]),
                           "Units Produced": np.round(base * .8 * trend, 0),
                           "Downtime Minutes": np.round(rng.gamma(2.0, 18, n), 1)})
    elif kind == "Inventory":
        df = pd.DataFrame({"Snapshot Date": dates,
                           "Warehouse": rng.choice(["WH-North", "WH-South", "WH-East", "WH-West"], n),
                           "Product Category": rng.choice(["Electronics", "Apparel", "Home", "Grocery"], n),
                           "SKU": [f"SKU-{rng.integers(1000, 9999)}" for _ in range(n)],
                           "Stock Qty": rng.integers(0, 900, n),
                           "Unit Cost": np.round(rng.uniform(40, 2200, n), 2),
                           "Reorder Level": rng.integers(50, 250, n)})
    else:  # Marketing
        channel = rng.choice(["Search Ads", "Social", "Email", "Display", "Affiliate"],
                             n, p=[.3, .26, .18, .14, .12])
        spend = np.round(base * 2.2 * trend, 2)
        df = pd.DataFrame({"Date": dates, "Channel": channel,
                           "Campaign": rng.choice(["Spring Launch", "Festive Push", "Always-On",
                                                   "Retargeting"], n),
                           "Impressions": (spend * rng.uniform(80, 160, n)).astype(int),
                           "Clicks": (spend * rng.uniform(2, 9, n)).astype(int),
                           "Spend": spend,
                           "Conversions": (spend * rng.uniform(.05, .4, n)).astype(int)})

    for col in df.columns:
        if df[col].dtype == object and col not in ("Customer ID", "Employee ID", "SKU"):
            idx = rng.choice(n, size=int(n * .012), replace=False)
            df.loc[idx, col] = np.nan
    num_cols = [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]
    if num_cols:
        idx = rng.choice(n, size=int(n * .008), replace=False)
        df[num_cols[0]] = df[num_cols[0]].astype(float)
        df.loc[idx, num_cols[0]] = df[num_cols[0]].max() * rng.uniform(4, 7, len(idx))
    df = pd.concat([df, df.iloc[:12]], ignore_index=True)  # deliberate duplicates
    return df.sort_values(df.columns[0]).reset_index(drop=True)


# =====================================================================
# DATASET TYPE DETECTION & PROFILING
# =====================================================================
TYPE_KEYWORDS = {
    "Sales": {"revenue": 2, "sales": 2, "product": 1, "customer": 1, "order": 1,
              "quantity": 1, "price": 1, "discount": 1, "units": 1},
    "Survey": {"survey": 2, "interviewer": 2, "project": 1, "status": 1, "response": 1,
               "satisfaction": 2, "respondent": 2, "questionnaire": 2},
    "Finance": {"budget": 2, "expense": 2, "profit": 1, "variance": 2, "ledger": 2,
                "invoice": 1, "payment": 1, "cost": 1, "actual": 1},
    "HR": {"employee": 2, "salary": 2, "attrition": 2, "department": 1, "hire": 2,
           "performance": 1, "gender": 1, "designation": 2, "leave": 1},
    "Operations": {"plant": 2, "shift": 2, "downtime": 2, "production": 2, "operations": 1,
                   "throughput": 2, "machine": 1, "utilization": 2, "units produced": 2},
    "Inventory": {"stock": 2, "inventory": 2, "warehouse": 2, "sku": 2, "reorder": 2,
                  "supply": 1, "on hand": 2},
    "Marketing": {"campaign": 2, "channel": 1, "impressions": 2, "clicks": 2, "ctr": 2,
                  "conversion": 2, "spend": 1, "ads": 1, "leads": 1},
}


def detect_dataset_type(df: pd.DataFrame) -> str:
    cols = " ".join(df.columns.astype(str).str.lower())
    scores = {kind: sum(w for kw, w in kws.items() if kw in cols)
              for kind, kws in TYPE_KEYWORDS.items()}
    best = max(scores, key=lambda k: (scores[k], -list(TYPE_KEYWORDS).index(k)))
    return f"{best} Dataset" if scores[best] > 0 else "General Dataset"


def profile_dataset(df: pd.DataFrame) -> dict:
    n = len(df)
    dates, nums, cats, texts, ids = [], [], [], [], []
    for col in df.columns:
        s = df[col]
        if pd.api.types.is_datetime64_any_dtype(s):
            dates.append(col)
            continue
        if _is_identifier(col):
            ids.append(col)
            continue
        nun = int(s.nunique(dropna=True))
        if pd.api.types.is_numeric_dtype(s) and not pd.api.types.is_bool_dtype(s):
            if not _is_metric_name(col):
                ids.append(col)
                continue
            if (n > 10 and pd.api.types.is_integer_dtype(s) and nun >= 0.95 * n
                    and _priority(col, NUMERIC_PRIORITY) == len(NUMERIC_PRIORITY)):
                ids.append(col)
                continue
            nums.append(col)
            if 2 <= nun <= 12:
                cats.append(col)
            continue
        if nun == 0:
            continue
        if n > 10 and nun >= 0.9 * n:
            texts.append(col)
            continue
        if nun <= 50:
            cats.append(col)
        else:
            texts.append(col)
    dims = sorted(set(cats), key=lambda c: (_priority(c, CATEGORY_PRIORITY),
                                            int(df[c].nunique(dropna=True))))
    metrics = sorted(nums, key=lambda c: (_priority(c, NUMERIC_PRIORITY),
                                          -float(df[c].abs().sum(skipna=True))))
    primary_metric = None
    if metrics:
        named = [c for c in metrics if _priority(c, NUMERIC_PRIORITY) < len(NUMERIC_PRIORITY)]
        primary_metric = named[0] if named else metrics[0]
    return {
        "dates": dates, "nums": nums, "cats": list(dims), "texts": texts, "ids": ids,
        "dimensions": dims[:10], "metrics": metrics,
        "primary_date": sorted(dates, key=lambda c: _priority(c, DATE_PRIORITY))[0] if dates else None,
        "primary_metric": primary_metric,
        "primary_cat": dims[0] if dims else None,
        "type": detect_dataset_type(df),
    }


# =====================================================================
# TREND / PERIOD / GROWTH ENGINE
# =====================================================================
def missing_pct(df: pd.DataFrame) -> float:
    return float(df.isna().sum().sum()) / float(df.size) * 100.0 if df.size else 0.0


def get_trend(df: pd.DataFrame, date_col, metric=None):
    """Adaptive daily / weekly / monthly series. Returns (series, label, rule)."""
    if date_col is None:
        return None, "", ""
    data = df[[date_col] + ([metric] if metric else [])].dropna(subset=[date_col])
    if data.empty:
        return None, "", ""
    data = data.sort_values(date_col)
    span = (data[date_col].max() - data[date_col].min()).days
    if span >= 90:
        rule, label = "MS", "Monthly"
    elif span >= 21:
        rule, label = "W-MON", "Weekly"
    else:
        rule, label = "D", "Daily"
    ts = data.set_index(date_col)
    series = ts[metric].resample(rule).sum() if metric else ts.resample(rule).size()
    series.name = str(metric) if metric else "Records"
    series = series.loc[series.index.notna()]
    if len(series) < 2:
        return None, "", ""
    return series, label, rule


def _bucket_start(s: pd.Series, rule: str) -> pd.Series:
    if rule == "MS":
        return s.dt.to_period("M").dt.to_timestamp()
    if rule == "W-MON":
        return s.dt.normalize() - pd.to_timedelta(s.dt.weekday, unit="D")
    return s.dt.normalize()


def _bucket_label(start: pd.Timestamp, rule: str) -> str:
    if rule == "MS":
        return start.strftime("%b %Y")
    if rule == "W-MON":
        return "Week of " + start.strftime("%d %b %Y")
    return start.strftime("%d %b %Y")


def period_growth(df: pd.DataFrame, prof: dict) -> dict:
    """Compare the latest period with the previous one (adaptive fallback)."""
    out = {"available": False, "reason": "No usable date column with at least two periods.",
           "metric": prof["primary_metric"] or "Records", "pct": None,
           "prev_total": None, "curr_total": None, "delta": None,
           "prev_label": "", "curr_label": "", "partial": False,
           "prev_mask": None, "curr_mask": None, "rule": ""}
    date_col, metric = prof["primary_date"], prof["primary_metric"]
    if date_col is None:
        out["reason"] = "The dataset has no date column, so period comparison is unavailable."
        return out
    s = df[date_col]
    valid = s.dropna()
    if len(valid) < 2 or valid.nunique() < 2:
        out["reason"] = "Not enough date history to compare two periods."
        return out
    span = (valid.max() - valid.min()).days
    rule = "MS" if span >= 62 else ("W-MON" if span >= 21 else "D")
    buckets = _bucket_start(s, rule)
    uniq = sorted(buckets.dropna().unique())
    if len(uniq) >= 2:
        curr_start, prev_start = pd.Timestamp(uniq[-1]), pd.Timestamp(uniq[-2])
        curr_mask, prev_mask = buckets == curr_start, buckets == prev_start
        prev_label, curr_label = _bucket_label(prev_start, rule), _bucket_label(curr_start, rule)
        end = (curr_start + pd.offsets.MonthEnd(1) if rule == "MS"
               else curr_start + pd.Timedelta(days=6 if rule == "W-MON" else 0))
        partial = bool(valid.max().normalize() < end.normalize())
    else:
        mid = valid.median()
        prev_mask, curr_mask = s < mid, s >= mid
        if not prev_mask.any() or not curr_mask.any():
            return out
        prev_label, curr_label = "Earlier period", "Recent period"
        partial = False
    prev_vals = df.loc[prev_mask, metric] if metric else df.loc[prev_mask]
    curr_vals = df.loc[curr_mask, metric] if metric else df.loc[curr_mask]
    prev_total = float(prev_vals.sum()) if metric else float(len(prev_vals))
    curr_total = float(curr_vals.sum()) if metric else float(len(curr_vals))
    delta = curr_total - prev_total
    pct = (delta / abs(prev_total) * 100.0) if prev_total != 0 else None
    out.update(available=True, reason="", prev_total=prev_total, curr_total=curr_total,
               delta=delta, pct=pct, prev_label=prev_label, curr_label=curr_label,
               partial=partial, prev_mask=prev_mask, curr_mask=curr_mask, rule=rule)
    return out


def top_categories(df: pd.DataFrame, cat, num=None, n: int = 10,
                   ascending: bool = False) -> pd.DataFrame | None:
    if cat is None:
        return None
    data = df[[cat] + ([num] if num else [])].dropna(subset=[cat])
    if data.empty:
        return None
    grouped = data.groupby(cat)[num].sum() if num else data.groupby(cat).size()
    value_name = str(num) if num else "Records"
    grouped = grouped.sort_values(ascending=ascending).head(n)
    out = grouped.reset_index()
    out.columns = [str(cat), value_name]
    out[str(cat)] = out[str(cat)].astype(str)
    return out


def outlier_stats(df: pd.DataFrame, col) -> dict:
    out = {"count": 0, "pct": 0.0, "low": None, "high": None, "example": ""}
    if col is None:
        return out
    s = df[col].dropna()
    if len(s) < 8:
        return out
    q1, q3 = float(s.quantile(.25)), float(s.quantile(.75))
    iqr = q3 - q1
    lo, hi = q1 - 1.5 * iqr, q3 + 1.5 * iqr
    mask = (df[col] < lo) | (df[col] > hi)
    count = int(mask.sum())
    out.update(count=count, pct=count / len(df) * 100.0 if len(df) else 0.0, low=lo, high=hi)
    if count:
        extreme_idx = df.loc[mask, col].abs().idxmax()
        out["example"] = f"{col} = {df.loc[extreme_idx, col]:,.2f} (row {int(extreme_idx) + 1})"
    return out


def trend_anomalies(series) -> list[dict]:
    if series is None or len(series) < 6:
        return []
    mean = series.rolling(3, min_periods=3).mean().shift(1)
    std = series.rolling(3, min_periods=3).std().shift(1).replace(0, np.nan)
    z = (series - mean) / std
    flags = z.abs() > 2.5
    return [{"period": str(idx.date()), "value": float(series.loc[idx]), "z": float(z.loc[idx])}
            for idx in series.index[flags.fillna(False)]]


def quality_score(df: pd.DataFrame, miss: float, dup_pct: float, out_pct: float) -> dict:
    constants = sum(int(df[c].nunique(dropna=True) <= 1) for c in df.columns)
    score = 100.0 - 1.5 * miss - 1.0 * dup_pct - 0.5 * out_pct - 2.0 * constants
    score = round(float(np.clip(score, 0, 100)), 1)
    grade = ("Excellent" if score >= 90 else "Good" if score >= 75
             else "Fair" if score >= 60 else "Poor")
    return {"score": score, "grade": grade, "constants": constants}


def analyze(df: pd.DataFrame, prof: dict) -> dict:
    """Single analysis pass consumed by every tab, report and the chat engine."""
    metric, cat, date_col = prof["primary_metric"], prof["primary_cat"], prof["primary_date"]
    records = len(df)
    dup = int(df.duplicated().sum())
    miss = missing_pct(df)
    growth = period_growth(df, prof)
    trend, trend_label, _ = get_trend(df, date_col, metric)
    top = top_categories(df, cat, metric, n=10)
    bottom = top_categories(df, cat, metric, n=5, ascending=True)
    share, hhi = None, None
    if top is not None:
        vals = top.iloc[:, 1].astype(float)
        total = float(vals.sum())
        if total > 0:
            shares = vals / total * 100.0
            share = float(shares.iloc[0])
            hhi = float((shares ** 2).sum())
    outliers = outlier_stats(df, metric)
    anomalies = trend_anomalies(trend)
    quality = quality_score(df, miss, dup / records * 100.0 if records else 0.0, outliers["pct"])
    corr = None
    usable = [c for c in prof["nums"] if df[c].nunique(dropna=True) > 1]
    if len(usable) >= 2:
        corr = df[usable].corr(numeric_only=True)
    if date_col is not None and df[date_col].notna().any():
        lo, hi = df[date_col].min(), df[date_col].max()
        date_text = (lo.strftime("%d %b %Y") if lo.date() == hi.date()
                     else f"{lo.strftime('%d %b %y')} → {hi.strftime('%d %b %y')}")
    else:
        date_text = "—"
    top_cat_text = "—"
    if cat is not None:
        modes = df[cat].dropna().mode()
        if not modes.empty:
            top_cat_text = clip(str(modes.iloc[0]), 18)
    kpis = {
        "records": records, "columns": df.shape[1], "missing_pct": miss,
        "duplicates": dup, "date_range": date_text, "top_category": top_cat_text,
        "metric_name": metric or "Records",
        "metric_total": float(df[metric].sum()) if metric else float(records),
        "quality": quality, "growth_pct": growth["pct"],
    }
    return {"kpis": kpis, "growth": growth, "trend": trend, "trend_label": trend_label,
            "top": top, "bottom": bottom, "share": share, "hhi": hhi,
            "outliers": outliers, "anomalies": anomalies, "corr": corr}


# =====================================================================
# SEGMENT GROWTH & ROOT CAUSE ANALYSIS
# =====================================================================
def segment_table(df: pd.DataFrame, dim: str, growth: dict, metric) -> pd.DataFrame:
    prev = df.loc[growth["prev_mask"]]
    curr = df.loc[growth["curr_mask"]]
    if metric:
        a = prev.groupby(dim)[metric].sum()
        b = curr.groupby(dim)[metric].sum()
    else:
        a = prev.groupby(dim).size()
        b = curr.groupby(dim).size()
    tab = pd.concat([a.rename("Previous"), b.rename("Current")], axis=1).fillna(0.0)
    tab["Change"] = tab["Current"] - tab["Previous"]
    total_change = float(tab["Change"].sum())
    denom = abs(total_change) if abs(total_change) > 1e-9 else float(tab["Change"].abs().sum())
    tab["Contribution %"] = (tab["Change"] / denom * 100.0) if denom > 1e-9 else 0.0
    tab["Growth %"] = np.where(tab["Previous"] != 0,
                               tab["Change"] / tab["Previous"].abs() * 100.0, np.nan)
    tab = tab.reindex(tab["Change"].abs().sort_values(ascending=False).index)
    tab.index = tab.index.astype(str)
    return tab.reset_index(names=str(dim))


def run_root_cause(df: pd.DataFrame, prof: dict, ana: dict) -> dict:
    g = ana["growth"]
    if not g["available"]:
        return {"available": False, "reason": g["reason"]}
    metric = prof["primary_metric"]
    metric_label = g["metric"]
    symbol_dims = prof["dimensions"][:8]
    dims = []
    for dim in symbol_dims:
        tab = segment_table(df, dim, g, metric)
        if tab.empty:
            continue
        dims.append({"dim": dim, "table": tab,
                     "score": float(tab["Change"].abs().sum())})
    dims.sort(key=lambda d: d["score"], reverse=True)
    delta, pct = g["delta"], g["pct"]
    direction = "increased" if delta > 0 else "decreased" if delta < 0 else "was unchanged"
    narrative = []
    head = (f"{metric_label} {direction} from {fmt_compact(g['prev_total'])} "
            f"({g['prev_label']}) to {fmt_compact(g['curr_total'])} ({g['curr_label']})")
    head += f" — a change of {fmt_compact(delta)}"
    head += f" ({fmt_pct(pct)})." if pct is not None else "."
    if g["partial"]:
        head += " Note: the current period is partial, which amplifies the apparent change."
    narrative.append(head)
    drivers, offsets = [], []
    for d in dims:
        for _, row in d["table"].head(5).iterrows():
            item = {"dim": d["dim"], "segment": str(row.iloc[0]),
                    "prev": float(row["Previous"]), "curr": float(row["Current"]),
                    "delta": float(row["Change"]), "contrib": float(row["Contribution %"])}
            if delta != 0 and item["delta"] * delta > 0:
                drivers.append(item)
            elif delta != 0 and item["delta"] * delta < 0:
                offsets.append(item)
    drivers.sort(key=lambda x: abs(x["delta"]), reverse=True)
    offsets.sort(key=lambda x: abs(x["delta"]), reverse=True)
    for i, drv in enumerate(drivers[:3]):
        lead = ("The change was driven primarily by" if i == 0
                else "The next largest driver was" if i == 1 else "A further driver was")
        narrative.append(
            f"{lead} **{drv['segment']}** ({drv['dim']}): "
            f"{fmt_compact(drv['prev'])} → {fmt_compact(drv['curr'])} "
            f"({fmt_compact(drv['delta'])}), accounting for {abs(drv['contrib']):.0f}% "
            f"of the total change.")
    for off in offsets[:2]:
        narrative.append(
            f"Partially offsetting the move, **{off['segment']}** ({off['dim']}) "
            f"{'grew' if off['delta'] > 0 else 'fell'} by {fmt_compact(abs(off['delta']))}.")
    n_prev = int(g["prev_mask"].sum())
    n_curr = int(g["curr_mask"].sum())
    vol_pct = (n_curr - n_prev) / n_prev * 100.0 if n_prev else None
    avg_prev = g["prev_total"] / n_prev if n_prev else 0.0
    avg_curr = g["curr_total"] / n_curr if n_curr else 0.0
    avg_pct = (avg_curr - avg_prev) / abs(avg_prev) * 100.0 if avg_prev else None
    if vol_pct is not None and avg_pct is not None:
        driver_kind = ("volume (fewer/more transactions)" if abs(vol_pct) >= abs(avg_pct)
                       else "average transaction value")
        narrative.append(
            f"Transaction volume changed {fmt_pct(vol_pct)} while average value changed "
            f"{fmt_pct(avg_pct)} — the move is mainly **{driver_kind}** driven.")
    for m in prof["metrics"][1:3]:
        mp = float(df.loc[g["prev_mask"], m].sum())
        mc = float(df.loc[g["curr_mask"], m].sum())
        if mp != 0:
            narrative.append(f"Related metric **{m}** moved {fmt_pct((mc - mp) / abs(mp) * 100.0)} "
                             f"over the same periods ({fmt_compact(mp)} → {fmt_compact(mc)}).")
    if ana["kpis"]["missing_pct"] > 5:
        narrative.append("Data quality caveat: missing values exceed 5%, so driver "
                         "estimates carry additional uncertainty.")
    waterfall = None
    if dims:
        best = dims[0]["table"].head(6)
        segs = best.iloc[:, 0].astype(str).tolist()
        deltas = best["Change"].astype(float).tolist()
        other = delta - float(sum(deltas))
        if abs(other) > 1e-9:
            segs.append("All other segments")
            deltas.append(other)
        waterfall = {"x": [g["prev_label"]] + segs + [g["curr_label"]],
                     "y": [g["prev_total"]] + deltas + [0],
                     "measure": ["absolute"] + ["relative"] * len(segs) + ["total"],
                     "title": f"{metric_label} bridge by {dims[0]['dim']}"}
    return {"available": True, "narrative": narrative, "dimensions": dims,
            "drivers": drivers[:8], "offsets": offsets[:5], "waterfall": waterfall,
            "growth": g, "volume_pct": vol_pct, "avg_pct": avg_pct}


# =====================================================================
# EXECUTIVE SUMMARY & AI INSIGHTS ENGINE
# =====================================================================
def _strip_md(text: str) -> str:
    return text.replace("**", "").replace("*", "")


def executive_summary(df: pd.DataFrame, prof: dict, ana: dict) -> str:
    k, g = ana["kpis"], ana["growth"]
    metric = prof["primary_metric"]
    parts = []
    if metric:
        parts.append(f"{metric} reached {fmt_compact(k['metric_total'])} across "
                     f"{k['records']:,} records in this {prof['type'].lower()}"
                     + (f" spanning {k['date_range']}." if k["date_range"] != "—" else "."))
    else:
        parts.append(f"The {prof['type'].lower()} contains {k['records']:,} records "
                     f"across {k['columns']} columns.")
    if ana["top"] is not None and ana["share"] is not None:
        name = str(ana["top"].iloc[0, 0])
        parts.append(f"{name} contributes {ana['share']:.0f}% of total "
                     f"{k['metric_name'].lower()} and remains the strongest performer.")
    if g["available"] and g["pct"] is not None:
        verb = "increased" if g["pct"] >= 0 else "decreased"
        caveat = " (current period is partial)" if g["partial"] else ""
        parts.append(f"{g['metric']} {verb} {abs(g['pct']):.0f}% compared with the "
                     f"previous period ({g['prev_label']} → {g['curr_label']}){caveat}.")
    elif g["available"]:
        parts.append(f"{g['metric']} moved from {fmt_compact(g['prev_total'])} to "
                     f"{fmt_compact(g['curr_total'])} versus the previous period.")
    else:
        parts.append("Period-over-period comparison is unavailable because the data "
                     "lacks sufficient date history.")
    q = k["quality"]
    parts.append(f"Data quality is {q['grade'].lower()} (score {q['score']:.0f}/100) with "
                 f"{k['missing_pct']:.1f}% missing values and {k['duplicates']:,} duplicate rows.")
    ins = build_insights(df, prof, ana)
    if ins["recommendations"]:
        first = _strip_md(ins["recommendations"][0]["detail"])
        parts.append("Management should " + first[0].lower() + first[1:])
    return " ".join(parts)


def build_insights(df: pd.DataFrame, prof: dict, ana: dict) -> dict:
    k, g = ana["kpis"], ana["growth"]
    metric, cat = prof["primary_metric"], prof["primary_cat"]
    what, risks, opps, recs = [], [], [], []

    # ---------------- WHAT HAPPENED ----------------
    what.append({"icon": "🗂️", "sev": "info", "title": "Dataset Overview",
                 "detail": f"{prof['type']} with **{k['records']:,} records** and "
                           f"**{k['columns']} columns**"
                           + (f", covering **{k['date_range']}**." if k["date_range"] != "—" else ".")})
    if ana["top"] is not None and not ana["top"].empty:
        name, val = str(ana["top"].iloc[0, 0]), float(ana["top"].iloc[0, 1])
        share_txt = f" ({ana['share']:.1f}% of total)" if ana["share"] is not None else ""
        what.append({"icon": "🏆", "sev": "success", "title": "Top Contributor",
                     "detail": f"**{name}** leads {k['metric_name']} with "
                               f"**{fmt_compact(val)}**{share_txt}."})
    if cat is not None:
        counts = df[cat].dropna().value_counts()
        if not counts.empty:
            what.append({"icon": "📦", "sev": "info", "title": "Largest Category",
                         "detail": f"**{counts.index[0]}** leads by volume with "
                                   f"**{int(counts.iloc[0]):,} records** across "
                                   f"{len(counts)} {cat} groups."})
    if metric and df[metric].notna().any():
        hi_idx, lo_idx = df[metric].idxmax(), df[metric].idxmin()
        hi_tag = f" ({df.loc[hi_idx, cat]})" if cat and pd.notna(df.loc[hi_idx, cat]) else ""
        lo_tag = f" ({df.loc[lo_idx, cat]})" if cat and pd.notna(df.loc[lo_idx, cat]) else ""
        what.append({"icon": "⬆️", "sev": "info", "title": "Highest Value",
                     "detail": f"Largest single {metric} value: "
                               f"**{fmt_compact(df.loc[hi_idx, metric])}**{hi_tag}."})
        what.append({"icon": "⬇️", "sev": "info", "title": "Lowest Value",
                     "detail": f"Smallest {metric} value: "
                               f"**{fmt_compact(df.loc[lo_idx, metric])}**{lo_tag}."})
    if g["available"]:
        if g["pct"] is not None:
            verb = "grew" if g["pct"] >= 0 else "declined"
            sev = "success" if g["pct"] >= 0 else "warning"
            what.append({"icon": "📈" if g["pct"] >= 0 else "📉", "sev": sev,
                         "title": "Trend & Period Change",
                         "detail": f"{g['metric']} **{verb} {abs(g['pct']):.1f}%** "
                                   f"({g['prev_label']} → {g['curr_label']}): "
                                   f"{fmt_compact(g['prev_total'])} → {fmt_compact(g['curr_total'])}."
                                   + (" Current period is partial." if g["partial"] else "")})
        else:
            what.append({"icon": "📈", "sev": "info", "title": "Trend & Period Change",
                         "detail": f"{g['metric']} moved {fmt_compact(g['prev_total'])} → "
                                   f"{fmt_compact(g['curr_total'])} (previous period was zero)."})
    elif ana["trend"] is not None:
        what.append({"icon": "📈", "sev": "info", "title": "Trend",
                     "detail": f"{ana['trend_label']} history spans "
                               f"{len(ana['trend'])} periods."})

    # ---------------- RISKS ----------------
    miss = k["missing_pct"]
    if miss > 0:
        worst = df.isna().mean().mul(100).idxmax()
        sev = "critical" if miss >= 15 else "warning" if miss >= 5 else "info"
        risks.append({"icon": "🧹", "sev": sev, "title": "Missing Values",
                      "detail": f"**{miss:.1f}%** of cells are missing. Worst column: "
                                f"**{worst}** ({df[worst].isna().mean() * 100:.1f}%). "
                                "Impute or re-collect before board-level reporting."})
    else:
        risks.append({"icon": "✅", "sev": "success", "title": "Completeness",
                      "detail": "No missing values detected — completeness risk is nil."})
    if k["duplicates"] > 0:
        risks.append({"icon": "👯", "sev": "warning" if k["duplicates"] / max(k["records"], 1) > .02 else "info",
                      "title": "Duplicate Rows",
                      "detail": f"**{k['duplicates']:,} duplicate rows** "
                                f"({k['duplicates'] / max(k['records'], 1) * 100:.1f}%) may double-count KPIs."})
    if ana["share"] is not None and (ana["share"] > 40 or (ana["hhi"] or 0) > 2500):
        risks.append({"icon": "🎯", "sev": "warning", "title": "Concentration Risk",
                      "detail": f"**{ana['top'].iloc[0, 0]}** drives **{ana['share']:.1f}%** of "
                                f"{k['metric_name']} (HHI {ana['hhi']:.0f}). A single segment "
                                "shock would materially hit results."})
    if ana["outliers"]["count"] > 0:
        risks.append({"icon": "🚨", "sev": "warning" if ana["outliers"]["pct"] > 2 else "info",
                      "title": "Outliers",
                      "detail": f"**{ana['outliers']['count']:,} outliers** "
                                f"({ana['outliers']['pct']:.1f}%) on {k['metric_name']} "
                                f"outside IQR fences. Extreme case: {ana['outliers']['example']}."})
    if ana["anomalies"]:
        a = ana["anomalies"][0]
        risks.append({"icon": "📡", "sev": "warning", "title": "Statistical Anomalies",
                      "detail": f"**{len(ana['anomalies'])} unusual period(s)** vs rolling history, "
                                f"e.g. {a['period']} (z-score {a['z']:.1f}). Investigate spikes/dips."})
    if g["available"] and g["pct"] is not None and g["pct"] < -5:
        risks.append({"icon": "📉", "sev": "critical", "title": "Performance Decline",
                      "detail": f"{g['metric']} fell **{abs(g['pct']):.1f}%** period-over-period. "
                                "Run Root Cause Analysis to isolate the driving segments."})

    # ---------------- OPPORTUNITIES ----------------
    if g["available"] and cat is not None:
        tab = segment_table(df, cat, g, metric)
        growers = tab[(tab["Change"] > 0) & (tab["Previous"] > 0)].copy()
        growers = growers.sort_values("Change", ascending=False)
        for _, row in growers.head(2).iterrows():
            opps.append({"icon": "🌱", "sev": "opportunity", "title": "Growth Segment",
                         "detail": f"**{row.iloc[0]}** added {fmt_compact(row['Change'])} "
                                   f"({fmt_pct(row['Growth %'])}) — strongest growing {cat}. "
                                   "Increase investment and capacity here."})
        emerging = tab[(tab["Previous"] == 0) & (tab["Current"] > 0)]
        for _, row in emerging.head(2).iterrows():
            opps.append({"icon": "✨", "sev": "opportunity", "title": "Emerging Category",
                         "detail": f"**{row.iloc[0]}** appeared this period with "
                                   f"{fmt_compact(row['Current'])} — nurture early momentum."})
        if ana["share"] is not None and metric:
            avg = df.groupby(cat)[metric].mean().sort_values(ascending=False)
            totals = df.groupby(cat)[metric].sum()
            total_all = float(totals.sum())
            if total_all > 0:
                for name in avg.index[:5]:
                    sh = float(totals[name]) / total_all * 100.0
                    if sh < 15 and float(avg[name]) > float(avg.median()):
                        opps.append({"icon": "💎", "sev": "opportunity",
                                     "title": "Underutilized Segment",
                                     "detail": f"**{name}** has above-median average {metric} "
                                               f"but only {sh:.1f}% share — scale this "
                                               "high-value, low-penetration segment."})
                        break
    if g["available"] and g["pct"] is not None and g["pct"] > 5:
        opps.append({"icon": "🚀", "sev": "opportunity", "title": "Positive Momentum",
                     "detail": f"{g['metric']} is up **{g['pct']:.1f}%**. Lock in the drivers "
                               "and replicate the playbook in lagging segments."})
    if not opps:
        opps.append({"icon": "🔭", "sev": "opportunity", "title": "Baseline Opportunity",
                     "detail": "No dominant growth pocket detected yet — expand date history "
                               "or segment granularity to surface emerging winners."})

    # ---------------- RECOMMENDATIONS ----------------
    if miss >= 5:
        worst = df.isna().mean().mul(100).idxmax()
        recs.append({"icon": "🛠️", "sev": "action", "title": "P1 · Stabilize Data Quality",
                     "detail": f"Fix the **{worst}** column first ({df[worst].isna().mean() * 100:.1f}% "
                               "missing) and add validation at the source before trusting KPIs."})
    if g["available"] and g["pct"] is not None and g["pct"] < -5:
        top_drv = ""
        rc = run_root_cause(df, prof, ana)
        if rc.get("available") and rc["drivers"]:
            d0 = rc["drivers"][0]
            top_drv = f" Start with **{d0['segment']}** ({d0['dim']})."
        recs.append({"icon": "🧭", "sev": "action", "title": "P1 · Arrest the Decline",
                     "detail": f"Launch a recovery review on the {abs(g['pct']):.1f}% drop.{top_drv} "
                               "Assign owners and a 2-week corrective plan."})
    if ana["share"] is not None and ana["share"] > 40:
        recs.append({"icon": "⚖️", "sev": "action", "title": "P2 · Reduce Concentration",
                     "detail": f"Diversify beyond **{ana['top'].iloc[0, 0]}** "
                               f"({ana['share']:.0f}% share) — set a target share ceiling and "
                               "grow the next two segments deliberately."})
    if opps and opps[0]["title"] in ("Growth Segment", "Emerging Category", "Underutilized Segment"):
        recs.append({"icon": "📈", "sev": "action", "title": "P2 · Scale What Works",
                     "detail": _strip_md(opps[0]["detail"])})
    if ana["outliers"]["count"] > 0:
        recs.append({"icon": "🔍", "sev": "action", "title": "P3 · Review Outliers",
                     "detail": f"Audit the {ana['outliers']['count']:,} extreme "
                               f"{k['metric_name']} records — separate genuine wins from "
                               "data-entry errors before forecasting."})
    if k["duplicates"] > 0:
        recs.append({"icon": "🧬", "sev": "action", "title": "P3 · De-duplicate Pipeline",
                     "detail": f"Remove {k['duplicates']:,} duplicate rows and add a uniqueness "
                               "check to the ingestion job."})
    recs.append({"icon": "📅", "sev": "action", "title": "P3 · Operating Cadence",
                 "detail": f"Review {k['metric_name']} weekly with segment owners; alert on "
                           "moves beyond ±10% and re-run Root Cause Analysis monthly."})
    return {"what_happened": what, "risks": risks, "opportunities": opps,
            "recommendations": recs[:6]}


# =====================================================================
# PLOTLY CHART FACTORY (auto-selected, dark-themed)
# =====================================================================
def _metric_or_records(prof):
    return prof["primary_metric"]


def chart_top_bar(df, prof, horizontal=False):
    cat, metric = prof["primary_cat"], _metric_or_records(prof)
    agg = top_categories(df, cat, metric, n=10)
    if agg is None or agg.empty:
        return None
    vcol = agg.columns[1]
    if horizontal:
        agg = agg.iloc[::-1]
        fig = px.bar(agg, x=vcol, y=cat, orientation="h", color=vcol,
                     color_continuous_scale="Blues", text_auto=".3s")
    else:
        fig = px.bar(agg, x=cat, y=vcol, color=vcol,
                     color_continuous_scale="Blues", text_auto=".3s")
        fig.update_xaxes(categoryorder="total descending")
    fig.update_layout(coloraxis_showscale=False)
    return style_fig(fig, f"🏆 Top {cat} by {vcol}")


def chart_share(df, prof, hole=0.55, by_metric=True):
    cat = prof["primary_cat"]
    metric = _metric_or_records(prof) if by_metric else None
    agg = top_categories(df, cat, metric, n=8)
    if agg is None or agg.empty:
        return None
    vcol = agg.columns[1]
    fig = px.pie(agg, names=cat, values=vcol, hole=hole)
    fig.update_traces(textinfo="percent+label", textfont_size=12)
    kind = "Donut" if hole else "Pie"
    return style_fig(fig, f"🥧 {kind} · {vcol} share by {cat}")


def chart_trend(df, prof, area=False):
    series, label, _ = get_trend(df, prof["primary_date"], _metric_or_records(prof))
    if series is None:
        return None
    if area:
        series = series.cumsum()
    tdf = series.rename_axis(prof["primary_date"]).reset_index()
    builder = px.area if area else px.line
    fig = builder(tdf, x=prof["primary_date"], y=series.name,
                  markers=not area)
    name = f"Cumulative {series.name}" if area else series.name
    return style_fig(fig, f"{'🌊' if area else '📅'} {label} {name} {'Area' if area else 'Trend'}")


def chart_scatter(df, prof, ana):
    nums = [c for c in prof["nums"] if df[c].nunique(dropna=True) > 1]
    if len(nums) < 2:
        return None
    x, y = nums[0], nums[1]
    if ana["corr"] is not None and len(nums) > 2:
        c = ana["corr"].abs().where(~np.eye(len(nums), dtype=bool))
        pair = c.stack().idxmax()
        x, y = pair[0], pair[1]
    sample = df.sample(min(len(df), 2000), random_state=42)
    color = prof["primary_cat"] if prof["primary_cat"] else None
    fig = px.scatter(sample, x=x, y=y, color=color, opacity=.7)
    return style_fig(fig, f"🔬 {x} vs {y}")


def chart_box(df, prof):
    cat, metric = prof["primary_cat"], prof["primary_metric"]
    if cat is None or metric is None:
        return None
    keep = df[cat].dropna().value_counts().head(12).index
    data = df[df[cat].isin(keep)]
    if data.empty:
        return None
    fig = px.box(data, x=cat, y=metric, points="outliers", color=cat)
    fig.update_layout(showlegend=False)
    return style_fig(fig, f"📦 {metric} distribution by {cat}")


def chart_heatmap(df, prof):
    dims = prof["dimensions"]
    if len(dims) < 2:
        return None
    d1, d2 = dims[0], dims[1]
    k1 = df[d1].dropna().value_counts().head(12).index
    k2 = df[d2].dropna().value_counts().head(12).index
    data = df[df[d1].isin(k1) & df[d2].isin(k2)]
    if data.empty:
        return None
    ct = pd.crosstab(data[d1], data[d2])
    fig = px.imshow(ct, text_auto=True, color_continuous_scale="Blues", aspect="auto")
    fig.update_layout(coloraxis_showscale=False)
    return style_fig(fig, f"🗺️ Record density · {d1} × {d2}")


def chart_corr(ana):
    if ana["corr"] is None or ana["corr"].shape[0] < 2:
        return None
    fig = px.imshow(ana["corr"], text_auto=".2f", zmin=-1, zmax=1,
                    color_continuous_scale="RdBu_r", aspect="auto")
    return style_fig(fig, "🧮 Correlation Heatmap")


def chart_hierarchy(df, prof, kind="treemap"):
    dims = prof["dimensions"][:2]
    metric = _metric_or_records(prof)
    if not dims or (kind == "sunburst" and len(dims) < 2):
        return None
    cols = dims + ([metric] if metric else [])
    data = df[cols].dropna(subset=dims)
    if data.empty:
        return None
    agg = data.groupby(dims)[metric].sum().reset_index() if metric \
        else data.groupby(dims).size().reset_index(name="Records")
    vcol = agg.columns[-1]
    if (agg[vcol] <= 0).all():
        return None
    agg = agg[agg[vcol] > 0]
    builder = px.treemap if kind == "treemap" else px.sunburst
    fig = builder(agg, path=dims, values=vcol, color=vcol,
                  color_continuous_scale="Blues")
    fig.update_layout(coloraxis_showscale=False)
    icon = "🌳" if kind == "treemap" else "☀️"
    return style_fig(fig, f"{icon} {kind.title()} · {' → '.join(dims)} by {vcol}", height=460)


def chart_dim_trend(df, prof, dim):
    metric, date_col = _metric_or_records(prof), prof["primary_date"]
    if date_col is None or dim is None:
        return None
    keep = df[dim].dropna().value_counts().head(5).index
    data = df[df[dim].isin(keep)]
    if data.empty:
        return None
    span = (data[date_col].max() - data[date_col].min()).days
    rule = "MS" if span >= 90 else ("W-MON" if span >= 21 else "D")
    tmp = data[[date_col, dim] + ([metric] if metric else [])].dropna(subset=[date_col])
    grp = tmp.set_index(date_col).groupby(dim)
    parts = []
    for name, g in grp:
        s = g[metric].resample(rule).sum() if metric else g.resample(rule).size()
        s.name = str(name)
        parts.append(s)
    if not parts:
        return None
    wide = pd.concat(parts, axis=1).fillna(0)
    if wide.shape[0] < 2:
        return None
    long = wide.rename_axis(date_col).reset_index().melt(
        id_vars=date_col, var_name=dim, value_name=metric or "Records")
    fig = px.line(long, x=date_col, y=metric or "Records", color=dim, markers=True)
    return style_fig(fig, f"📅 {metric or 'Records'} trend by {dim} (top 5)")


def build_waterfall_fig(rca):
    wf = rca.get("waterfall")
    if not wf:
        return None
    fig = go.Figure(go.Waterfall(
        x=wf["x"], y=wf["y"], measure=wf["measure"],
        text=[fmt_compact(v) if m != "total" else "" for v, m in zip(wf["y"], wf["measure"])],
        textposition="outside",
        increasing={"marker": {"color": GREEN}},
        decreasing={"marker": {"color": RED}},
        totals={"marker": {"color": BLUE}},
        connector={"line": {"color": "rgba(148,163,184,0.4)"}},
    ))
    return style_fig(fig, "🌉 " + wf["title"], height=440)


def auto_charts(df, prof, ana) -> list[dict]:
    """Auto-select every chart the data can support, with rationale."""
    specs = [
        ("Top 10 Ranking", lambda: chart_top_bar(df, prof, horizontal=True),
         "Ranking of the leading categories by the primary metric."),
        ("Category Bar", lambda: chart_top_bar(df, prof),
         "Compare category totals side by side."),
        ("Metric Donut", lambda: chart_share(df, prof, hole=.55, by_metric=True),
         "Share of the primary metric held by each category."),
        ("Volume Pie", lambda: chart_share(df, prof, hole=0, by_metric=False),
         "Share of records contributed by each category."),
        ("Trend Line", lambda: chart_trend(df, prof),
         "How the primary metric evolves over time."),
        ("Cumulative Area", lambda: chart_trend(df, prof, area=True),
         "Running total highlights accumulation pace."),
        ("Segment Trends", lambda: chart_dim_trend(df, prof, prof["primary_cat"]),
         "Top segments plotted over time to expose diverging trajectories."),
        ("Scatter", lambda: chart_scatter(df, prof, ana),
         "Relationship between the two most correlated numeric fields."),
        ("Box Plot", lambda: chart_box(df, prof),
         "Spread and outliers of the metric within each category."),
        ("Density Heatmap", lambda: chart_heatmap(df, prof),
         "Where records concentrate across two dimensions."),
        ("Correlation Heatmap", lambda: chart_corr(ana),
         "Linear relationships between all numeric metrics."),
        ("Treemap", lambda: chart_hierarchy(df, prof, "treemap"),
         "Hierarchical composition of the metric."),
        ("Sunburst", lambda: chart_hierarchy(df, prof, "sunburst"),
         "Two-level composition ring for drill-down thinking."),
    ]
    out = []
    for title, builder, why in specs:
        try:
            fig = builder()
        except Exception:
            fig = None
        if fig is not None:
            out.append({"title": title, "fig": fig, "why": why})
    return out


# =====================================================================
# STATIC CHART IMAGES (Matplotlib) — embedded into PDF & PPTX
# =====================================================================
def _compact_axis(v):
    a = abs(v)
    for div, suf in ((1e9, "B"), (1e6, "M"), (1e3, "K")):
        if a >= div:
            return f"{v / div:.1f}{suf}"
    return f"{v:,.0f}"


def _mpl_axes(w=8.0, h=4.2):
    fig, ax = plt.subplots(figsize=(w, h), dpi=150)
    fig.patch.set_facecolor("#0B1220")
    ax.set_facecolor("#0B1220")
    ax.tick_params(colors="#B9C6DE", labelsize=9)
    for spine in ax.spines.values():
        spine.set_color("#24344F")
    ax.grid(True, color="#1B2A44", linewidth=.7, alpha=.9)
    ax.set_axisbelow(True)
    ax.yaxis.set_major_formatter(FuncFormatter(lambda v, p: _compact_axis(v)))
    return fig, ax


def _fig_bytes(fig, fmt="png") -> bytes:
    buf = io.BytesIO()
    fig.savefig(buf, format=fmt, facecolor=fig.get_facecolor(), bbox_inches="tight")
    plt.close(fig)
    return buf.getvalue()


def png_bar(labels, values, title) -> bytes:
    fig, ax = plt.subplots(figsize=(8, 4.2), dpi=150)
    fig.patch.set_facecolor("#0B1220")
    ax.set_facecolor("#0B1220")
    colors = [PALETTE[i % len(PALETTE)] for i in range(len(labels))]
    ypos = np.arange(len(labels))[::-1]
    ax.barh(ypos, values, color=colors, height=.62)
    ax.set_yticks(ypos, [clip(str(l), 22) for l in labels])
    ax.tick_params(colors="#B9C6DE", labelsize=9)
    for spine in ax.spines.values():
        spine.set_color("#24344F")
    ax.grid(True, axis="x", color="#1B2A44", linewidth=.7)
    ax.set_axisbelow(True)
    ax.xaxis.set_major_formatter(FuncFormatter(lambda v, p: _compact_axis(v)))
    for y, v in zip(ypos, values):
        ax.text(v, y, " " + _compact_axis(v), va="center", color="#E5EAF5", fontsize=8)
    ax.set_title(title, color="#E5EAF5", fontsize=12, fontweight="bold", loc="left")
    ax.margins(x=.18)
    return _fig_bytes(fig)


def png_line(x, y, title) -> bytes:
    fig, ax = _mpl_axes()
    ax.plot(x, y, color=BLUE, linewidth=2.2, marker="o", markersize=4)
    ax.fill_between(x, y, color=BLUE, alpha=.15)
    ax.set_title(title, color="#E5EAF5", fontsize=12, fontweight="bold", loc="left")
    fig.autofmt_xdate()
    return _fig_bytes(fig)


def png_donut(labels, values, title) -> bytes:
    fig, ax = plt.subplots(figsize=(6.4, 4.2), dpi=150)
    fig.patch.set_facecolor("#0B1220")
    ax.set_facecolor("#0B1220")
    colors = [PALETTE[i % len(PALETTE)] for i in range(len(labels))]
    wedges, _, autotexts = ax.pie(
        values, colors=colors, autopct="%1.0f%%", pctdistance=.78,
        wedgeprops=dict(width=.42, edgecolor="#0B1220"),
        textprops=dict(color="#E5EAF5", fontsize=8))
    ax.legend(wedges, [clip(str(l), 20) for l in labels], loc="center left",
              bbox_to_anchor=(1.0, .5), frameon=False,
              labelcolor="#B9C6DE", fontsize=8)
    ax.set_title(title, color="#E5EAF5", fontsize=12, fontweight="bold", loc="left")
    return _fig_bytes(fig)


def png_drivers(drivers, title) -> bytes:
    labels = [f"{d['segment']} ({d['dim']})" for d in drivers][:8]
    values = [d["delta"] for d in drivers][:8]
    fig, ax = plt.subplots(figsize=(8, 4.2), dpi=150)
    fig.patch.set_facecolor("#0B1220")
    ax.set_facecolor("#0B1220")
    ypos = np.arange(len(labels))[::-1]
    colors = [GREEN if v >= 0 else RED for v in values]
    ax.barh(ypos, values, color=colors, height=.62)
    ax.set_yticks(ypos, [clip(l, 28) for l in labels])
    ax.tick_params(colors="#B9C6DE", labelsize=9)
    for spine in ax.spines.values():
        spine.set_color("#24344F")
    ax.grid(True, axis="x", color="#1B2A44", linewidth=.7)
    ax.set_axisbelow(True)
    ax.xaxis.set_major_formatter(FuncFormatter(lambda v, p: _compact_axis(v)))
    ax.axvline(0, color="#475569", linewidth=1)
    ax.set_title(title, color="#E5EAF5", fontsize=12, fontweight="bold", loc="left")
    ax.margins(x=.15)
    return _fig_bytes(fig)


# =====================================================================
# LLM CLIENTS — OpenAI → Ollama → Built-in Analyst (automatic fallback)
# =====================================================================
DEFAULT_OLLAMA_MODELS = ["qwen2.5-coder", "llama3", "qwen", "mistral", "deepseek-r1"]
OPENAI_MODELS = ["gpt-4o-mini", "gpt-4o", "gpt-4.1-mini"]
SYSTEM_PROMPT = (
    "You are Business Copilot AI, a senior BI analyst presenting to executives. "
    "Answer strictly from the DATA CONTEXT provided; never invent numbers. "
    "Use concrete figures, business language and short paragraphs or bullets. "
    "Keep answers under 250 words. If the context cannot answer the question, "
    "say what data is missing.")


def fetch_ollama_models(base_url: str) -> list[str]:
    r = requests.get(base_url.rstrip("/") + "/api/tags", timeout=3)
    r.raise_for_status()
    return sorted(m["name"] for m in r.json().get("models", []))


def ask_openai(api_key: str, model: str, messages: list[dict]) -> str:
    r = requests.post(
        "https://api.openai.com/v1/chat/completions",
        headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
        json={"model": model, "messages": messages, "temperature": 0.3},
        timeout=(5, 90))
    if r.status_code != 200:
        raise RuntimeError(f"OpenAI HTTP {r.status_code}: {clip(r.text, 180)}")
    return r.json()["choices"][0]["message"]["content"].strip()


def ask_ollama(base_url: str, model: str, messages: list[dict]) -> str:
    r = requests.post(
        base_url.rstrip("/") + "/api/chat",
        json={"model": model, "messages": messages, "stream": False,
              "options": {"temperature": 0.3}},
        timeout=(5, 180))
    if r.status_code != 200:
        raise RuntimeError(f"Ollama HTTP {r.status_code}: {clip(r.text, 180)}")
    return r.json()["message"]["content"].strip()


def build_chat_context(df, prof, ana, insights, rca=None) -> str:
    k = ana["kpis"]
    lines = [
        f"Dataset type: {prof['type']} | Rows: {k['records']:,} | Columns: {k['columns']}",
        f"Date range: {k['date_range']} | Primary metric: {k['metric_name']} = {fmt_compact(k['metric_total'])}",
        f"Missing: {k['missing_pct']:.2f}% | Duplicates: {k['duplicates']:,} | "
        f"Quality: {k['quality']['score']}/100 ({k['quality']['grade']})",
        f"Dimensions: {', '.join(prof['dimensions']) or 'none'} | "
        f"Metrics: {', '.join(prof['metrics']) or 'none'}",
    ]
    g = ana["growth"]
    if g["available"]:
        lines.append(f"Period comparison: {g['prev_label']} {fmt_compact(g['prev_total'])} → "
                     f"{g['curr_label']} {fmt_compact(g['curr_total'])} "
                     f"({fmt_pct(g['pct'])}){' [current period partial]' if g['partial'] else ''}")
    if ana["top"] is not None:
        lines.append("Top categories: " + "; ".join(
            f"{r.iloc[0]}={fmt_compact(r.iloc[1])}" for _, r in ana["top"].head(5).iterrows()))
    for section in ("what_happened", "risks", "opportunities", "recommendations"):
        lines.append(section.upper() + ": " + " | ".join(
            f"{i['title']}: {clip(_strip_md(i['detail']), 140)}" for i in insights[section]))
    if rca and rca.get("available"):
        lines.append("ROOT CAUSE: " + " ".join(clip(_strip_md(n), 220) for n in rca["narrative"][:5]))
    return clip("\n".join(lines), 6500)


def answer_question(question, df, prof, ana, insights, settings, rca=None) -> dict:
    context = build_chat_context(df, prof, ana, insights, rca)
    messages = [{"role": "system", "content": SYSTEM_PROMPT + "\n\nDATA CONTEXT:\n" + context},
                {"role": "user", "content": question}]
    errors = []
    if settings.get("openai_key"):
        try:
            text = ask_openai(settings["openai_key"], settings["openai_model"], messages)
            return {"text": text, "fig": None, "engine": f"OpenAI · {settings['openai_model']}"}
        except Exception as err:
            errors.append(f"OpenAI: {clip(err, 120)}")
    try:
        text = ask_ollama(settings["ollama_url"], settings["ollama_model"], messages)
        return {"text": text, "fig": None, "engine": f"Ollama · {settings['ollama_model']}"}
    except Exception as err:
        errors.append(f"Ollama: {clip(err, 120)}")
    result = rule_based_answer(question, df, prof, ana, insights, rca)
    result["engine"] = "Built-in Analyst (offline fallback)"
    if errors:
        result["fallback_note"] = " · ".join(errors)
    return result


# =====================================================================
# BUILT-IN ANALYST — deterministic NLQ engine (works with zero LLMs)
# =====================================================================
def _find_dimension(q: str, prof: dict):
    for dim in prof["dimensions"]:
        if str(dim).lower() in q:
            return dim
    return None


def _find_values(q: str, df: pd.DataFrame, prof: dict) -> list[tuple]:
    matches = []
    for dim in prof["dimensions"][:6]:
        for val in df[dim].dropna().unique()[:300]:
            sval = str(val)
            if len(sval) >= 3 and sval.lower() in q:
                matches.append((dim, sval))
    return matches


def _bullet(items) -> str:
    return "\n".join(f"- **{i['title']}** — {_strip_md(i['detail'])}" for i in items)


def rule_based_answer(question, df, prof, ana, insights, rca=None) -> dict:
    q = question.lower().strip()
    k = ana["kpis"]
    metric = prof["primary_metric"]
    fig = None

    def ans(text, figure=None):
        return {"text": text, "fig": figure}

    if any(w in q for w in ("root cause", "driver", "why")):
        rc = rca if rca and rca.get("available") else run_root_cause(df, prof, ana)
        if rc.get("available"):
            return ans("\n\n".join(rc["narrative"]), build_waterfall_fig(rc))
        return ans(f"Root cause analysis is unavailable: {rc.get('reason', 'insufficient history')}")
    if "risk" in q or "concern" in q or "worried" in q:
        return ans("**Key risks right now:**\n\n" + _bullet(insights["risks"]))
    if "opportunit" in q or "growth segment" in q or "upside" in q:
        return ans("**Where the upside is:**\n\n" + _bullet(insights["opportunities"]))
    if any(w in q for w in ("recommend", "action", "suggest", "next step", "should")):
        return ans("**Prioritized recommendations:**\n\n" + _bullet(insights["recommendations"]))
    if "compare" in q:
        matches = _find_values(q, df, prof)
        if len(matches) >= 2:
            rows = []
            for dim, val in matches[:4]:
                sub = df[df[dim].astype(str) == val]
                total = float(sub[metric].sum()) if metric else len(sub)
                avg = float(sub[metric].mean()) if metric and len(sub) else 0.0
                rows.append(f"- **{val}** ({dim}): {fmt_compact(total)} total · "
                            f"{len(sub):,} records · avg {fmt_compact(avg)}")
            return ans("**Comparison:**\n\n" + "\n".join(rows))
        if ana["top"] is not None and len(ana["top"]) >= 2:
            t, b = ana["top"].iloc[0], ana["top"].iloc[-1]
            return ans(f"**{t.iloc[0]}** leads with {fmt_compact(t.iloc[1])} versus "
                       f"**{b.iloc[0]}** at {fmt_compact(b.iloc[1])} within the top 10. "
                       "Name two specific segments (e.g. “compare North and South”) for a deeper cut.")
    if "trend" in q or "over time" in q or "timeline" in q:
        dim = _find_dimension(q, prof)
        if dim:
            fig = chart_dim_trend(df, prof, dim)
            g = ana["growth"]
            extra = ""
            if g["available"]:
                tab = segment_table(df, dim, g, metric).head(5)
                extra = "\n\nLatest period movers:\n" + "\n".join(
                    f"- **{r.iloc[0]}**: {fmt_compact(r['Previous'])} → {fmt_compact(r['Current'])} "
                    f"({fmt_pct(r['Growth %'])})" for _, r in tab.iterrows())
            return ans(f"**{k['metric_name']} trend by {dim}** (chart below).{extra}", fig)
        fig = chart_trend(df, prof)
        g = ana["growth"]
        txt = (f"{g['metric']} moved {fmt_compact(g['prev_total'])} → {fmt_compact(g['curr_total'])} "
               f"({fmt_pct(g['pct'])}) between {g['prev_label']} and {g['curr_label']}."
               if g["available"] else "No usable date history for a trend.")
        return ans(txt, fig)
    if any(w in q for w in ("top", "best", "largest", "highest", "contributor", "leading", "strongest")):
        if ana["top"] is not None and not ana["top"].empty:
            t = ana["top"].iloc[0]
            share = f" ({ana['share']:.1f}% of total)" if ana["share"] is not None else ""
            table = "\n".join(f"- {r.iloc[0]}: {fmt_compact(r.iloc[1])}"
                              for _, r in ana["top"].head(5).iterrows())
            return ans(f"**{t.iloc[0]}** is the top contributor with "
                       f"**{fmt_compact(t.iloc[1])}**{share}.\n\nLeaderboard:\n{table}",
                       chart_top_bar(df, prof))
        return ans("No categorical dimension is available to rank contributors.")
    if any(w in q for w in ("lowest", "bottom", "smallest", "weakest", "worst perform")):
        if ana["bottom"] is not None and not ana["bottom"].empty:
            table = "\n".join(f"- {r.iloc[0]}: {fmt_compact(r.iloc[1])}"
                              for _, r in ana["bottom"].iterrows())
            return ans(f"**{ana['bottom'].iloc[0, 0]}** is the weakest segment.\n\n{table}")
        return ans("No categorical dimension is available to rank the bottom performers.")
    if any(w in q for w in ("summary", "summarize", "summarise", "executive", "overview", "brief")):
        return ans("**Executive summary**\n\n" + executive_summary(df, prof, ana))
    if any(w in q for w in ("explain", "describe", "what is this", "about", "dataset")):
        return ans(
            f"This is a **{prof['type']}** with **{k['records']:,} rows × {k['columns']} columns**"
            + (f", covering **{k['date_range']}**" if k["date_range"] != "—" else "") + ".\n\n"
            f"- **Dimensions:** {', '.join(prof['dimensions']) or 'none detected'}\n"
            f"- **Metrics:** {', '.join(prof['metrics']) or 'none detected'}\n"
            f"- **Primary metric:** {k['metric_name']} = {fmt_compact(k['metric_total'])}\n"
            f"- **Quality:** {k['quality']['score']}/100 ({k['quality']['grade']}), "
            f"{k['missing_pct']:.1f}% missing, {k['duplicates']:,} duplicates")
    if any(w in q for w in ("missing", "quality", "duplicate", "clean")):
        qd = k["quality"]
        worst = df.isna().mean().mul(100).sort_values(ascending=False).head(3)
        return ans(f"Quality score **{qd['score']}/100 ({qd['grade']})**. "
                   f"Missing **{k['missing_pct']:.2f}%**, duplicates **{k['duplicates']:,}**, "
                   f"outliers **{ana['outliers']['count']:,}**.\n\nWorst columns:\n"
                   + "\n".join(f"- {c}: {v:.1f}% missing" for c, v in worst.items()))
    if "outlier" in q or "anomal" in q:
        o = ana["outliers"]
        txt = (f"**{o['count']:,} outliers** ({o['pct']:.1f}%) on {k['metric_name']}. "
               f"Extreme case: {o['example'] or 'n/a'}.")
        if ana["anomalies"]:
            txt += "\n\nAnomalous periods:\n" + "\n".join(
                f"- {a['period']}: {fmt_compact(a['value'])} (z {a['z']:.1f})"
                for a in ana["anomalies"][:5])
        return ans(txt)
    if "correlat" in q or "relationship" in q:
        if ana["corr"] is not None:
            c = ana["corr"].abs().where(~np.eye(len(ana["corr"]), dtype=bool))
            pairs = c.stack().sort_values(ascending=False)
            seen, rows = set(), []
            for (a, b), v in pairs.items():
                key = tuple(sorted((a, b)))
                if key in seen:
                    continue
                seen.add(key)
                rows.append(f"- **{a} ↔ {b}**: r = {ana['corr'].loc[a, b]:.2f}")
                if len(rows) >= 5:
                    break
            return ans("**Strongest relationships:**\n\n" + "\n".join(rows), chart_corr(ana))
        return ans("At least two varying numeric columns are needed for correlation analysis.")
    if any(w in q for w in ("how many", "count", "records", "rows")):
        return ans(f"The current view contains **{k['records']:,} records** across "
                   f"**{k['columns']} columns** (global filters applied).")
    if any(w in q for w in ("average", "mean", "median")) and metric:
        s = df[metric].dropna()
        return ans(f"**{metric}** — mean {fmt_compact(s.mean())}, median {fmt_compact(s.median())}, "
                   f"min {fmt_compact(s.min())}, max {fmt_compact(s.max())}.")
    return ans(
        "I can answer from your uploaded data. Try:\n"
        "- *Why did revenue decline?* · *What is the biggest risk?*\n"
        "- *Who is the top contributor?* · *Show trends by region*\n"
        "- *Compare North and South* · *Generate executive summary*\n"
        "- *Explain the dataset* · *Show outliers* · *Correlations*\n\n"
        "**Quick snapshot:** " + executive_summary(df, prof, ana))


# =====================================================================
# COLUMN PROFILE & REPORT DATA
# =====================================================================
def column_profile(df: pd.DataFrame) -> pd.DataFrame:
    rows, n = [], len(df)
    for col in df.columns:
        s = df[col]
        missing = int(s.isna().sum())
        example = str(s.dropna().iloc[0])[:40] if s.notna().any() else ""
        rows.append({"Column": str(col), "Type": str(s.dtype), "Missing": missing,
                     "Missing %": round(missing / n * 100.0, 2) if n else 0.0,
                     "Unique": int(s.nunique(dropna=True)), "Example": example})
    return pd.DataFrame(rows)


def _md_table(headers: list[str], rows: list[list]) -> str:
    out = ["| " + " | ".join(headers) + " |",
           "| " + " | ".join("---" for _ in headers) + " |"]
    for r in rows:
        out.append("| " + " | ".join(str(c) for c in r) + " |")
    return "\n".join(out)


def _ins_lines(items) -> str:
    return "\n".join(f"- **{i['title']}** — {_strip_md(i['detail'])}" for i in items)


def build_text_reports(df, prof, ana, ins, rca, source_name) -> dict:
    k, g = ana["kpis"], ana["growth"]
    stamp = datetime.now().strftime("%d %b %Y %H:%M")
    head = f"# {APP_NAME} {APP_VERSION}\n\n**Source:** {source_name} · **Type:** {prof['type']} · **Generated:** {stamp}\n\n"
    summary = executive_summary(df, prof, ana)
    kpi_rows = [
        ["Total Records", f"{k['records']:,}"], ["Total Columns", k["columns"]],
        ["Missing Data %", f"{k['missing_pct']:.2f}%"], ["Duplicate Rows", f"{k['duplicates']:,}"],
        ["Date Range", k["date_range"]], ["Top Category", k["top_category"]],
        ["Primary Metric", f"{k['metric_name']} = {fmt_compact(k['metric_total'])}"],
        ["Data Quality Score", f"{k['quality']['score']}/100 ({k['quality']['grade']})"],
        ["Growth % (PoP)", fmt_pct(k["growth_pct"])],
    ]
    kpi_md = _md_table(["KPI", "Value"], kpi_rows)
    reports = {}
    reports["Management Report"] = (
        head + "## Executive Summary\n\n" + summary + "\n\n## KPI Summary\n\n" + kpi_md
        + "\n\n## What Happened\n\n" + _ins_lines(ins["what_happened"])
        + "\n\n## Risks\n\n" + _ins_lines(ins["risks"])
        + "\n\n## Opportunities\n\n" + _ins_lines(ins["opportunities"])
        + "\n\n## Recommendations\n\n" + _ins_lines(ins["recommendations"])
        + ("\n\n## Root Cause Analysis\n\n" + "\n\n".join(_strip_md(n) for n in rca["narrative"])
           if rca and rca.get("available") else ""))
    reports["Executive Summary"] = (head + "## Executive Summary\n\n" + summary
                                    + "\n\n## Headline Numbers\n\n" + kpi_md)
    reports["KPI Summary"] = (head + "## KPI Summary\n\n" + kpi_md
                              + "\n\n*Quality score = 100 − 1.5×missing% − 1.0×duplicate% − "
                                "0.5×outlier% − 2×constant columns (clamped 0–100).*")
    risk_stats = (f"\n\n**Outliers:** {ana['outliers']['count']:,} "
                  f"({ana['outliers']['pct']:.1f}%) · **Anomalous periods:** "
                  f"{len(ana['anomalies'])}")
    if ana["hhi"] is not None:
        risk_stats += f" · **HHI:** {ana['hhi']:.0f}"
    reports["Risk Report"] = head + "## Risk Report\n\n" + _ins_lines(ins["risks"]) + risk_stats
    opp_extra = ""
    if g["available"] and prof["primary_cat"]:
        tab = segment_table(df, prof["primary_cat"], g, prof["primary_metric"]).head(8)
        opp_extra = "\n\n## Segment Movement\n\n" + _md_table(
            list(tab.columns),
            [[r.iloc[0], fmt_compact(r["Previous"]), fmt_compact(r["Current"]),
              fmt_compact(r["Change"]), fmt_pct(r["Growth %"])] for _, r in tab.iterrows()])
    reports["Opportunity Report"] = (head + "## Opportunity Report\n\n"
                                     + _ins_lines(ins["opportunities"]) + opp_extra)
    if rca and rca.get("available"):
        drv_md = _md_table(["Segment", "Dimension", "Previous", "Current", "Change", "Contribution"],
                           [[d["segment"], d["dim"], fmt_compact(d["prev"]), fmt_compact(d["curr"]),
                             fmt_compact(d["delta"]), f"{d['contrib']:.0f}%"] for d in rca["drivers"]])
        reports["Root Cause Report"] = (head + "## Root Cause Report\n\n"
                                        + "\n\n".join(_strip_md(n) for n in rca["narrative"])
                                        + "\n\n## Driver Table\n\n" + drv_md)
    else:
        reports["Root Cause Report"] = (head + "## Root Cause Report\n\n"
                                        + (rca.get("reason", "Root cause analysis unavailable.")
                                           if rca else "Run Root Cause Analysis first."))
    return reports


# =====================================================================
# PURE-PYTHON PDF ENGINE (no external PDF dependency)
# =====================================================================
PDF_W, PDF_H = 595.27, 841.89
PDF_MARGIN = 46.0
PDF_HEADER_H = 64.0
PDF_TOP = PDF_H - PDF_HEADER_H - 26.0
PDF_BOTTOM = 56.0
_HELV_W = list(map(int, ("278 278 355 556 556 889 667 191 333 333 389 584 278 333 278 278 "
                         "556 556 556 556 556 556 556 556 556 556 278 278 584 584 584 556 "
                         "1015 667 667 722 722 667 611 778 722 278 500 667 556 833 722 778 "
                         "667 778 722 667 611 722 667 944 667 667 611 278 278 278 469 556 "
                         "333 556 556 500 556 556 278 556 556 222 222 500 222 833 556 556 "
                         "556 556 333 500 278 556 500 722 500 500 500 334 260 334 584").split()))
_HELV_B_W = list(map(int, ("278 333 474 556 556 889 722 238 333 333 389 584 278 333 278 278 "
                           "556 556 556 556 556 556 556 556 556 556 333 333 584 584 584 611 "
                           "975 722 722 722 722 667 611 778 722 278 556 722 611 833 722 778 "
                           "667 778 722 667 611 722 667 944 667 667 611 333 278 333 584 556 "
                           "333 556 611 556 611 556 333 611 611 278 278 556 278 889 611 611 "
                           "611 611 389 556 333 611 556 778 556 556 500 389 280 389 584").split()))
_PDF_REPL = {"₹": "Rs.", "→": "->", "—": "-", "–": "-", "’": "'", "‘": "'",
             "“": '"', "”": '"', "…": "...", "•": "-", "≥": ">=", "≤": "<=",
             "≈": "~", "×": "x", "Δ": "Delta", "±": "+/-", "↔": "<->", "·": "-"}


def _pdf_safe(text) -> str:
    text = _strip_md(str(text))
    for a, b in _PDF_REPL.items():
        text = text.replace(a, b)
    return text.encode("cp1252", errors="ignore").decode("cp1252")


def _pdf_escape(text: str) -> bytes:
    out = []
    for b in _pdf_safe(text).encode("cp1252"):
        if b in (40, 41, 92):
            out.append("\\" + chr(b))
        elif 32 <= b <= 126:
            out.append(chr(b))
        else:
            out.append(f"\\{b:03o}")
    return "".join(out).encode("latin-1")


def _hex_rgb(h: str) -> tuple:
    h = h.lstrip("#")
    return tuple(int(h[i:i + 2], 16) / 255.0 for i in (0, 2, 4))


def _text_w(s: str, size: float, bold: bool) -> float:
    widths = _HELV_B_W if bold else _HELV_W
    return sum(widths[ord(c) - 32] if 32 <= ord(c) <= 126 else 556
               for c in _pdf_safe(s)) / 1000.0 * size


class PDFDoc:
    """Minimal, dependency-free PDF writer: text, tables, KPI cards, JPEG images."""

    def __init__(self, report_title: str, source: str):
        self.report_title = report_title
        self.source = source
        self.pages: list[list[bytes]] = []
        self.images: list[tuple[bytes, int, int]] = []
        self.ops: list[bytes] = []
        self.y = PDF_TOP
        self._new_page()

    # ---------- low level ----------
    def _new_page(self):
        self.ops = []
        self.pages.append(self.ops)
        self.y = PDF_TOP

    def ensure(self, h: float):
        if self.y - h < PDF_BOTTOM:
            self._new_page()

    def _rect(self, x, y, w, h, fill, stroke=None, lw=0.6):
        r, g, b = _hex_rgb(fill)
        self.ops.append(f"{r:.3f} {g:.3f} {b:.3f} rg {x:.2f} {y:.2f} {w:.2f} {h:.2f} re f".encode())
        if stroke:
            r, g, b = _hex_rgb(stroke)
            self.ops.append(f"{r:.3f} {g:.3f} {b:.3f} RG {lw:.2f} w {x:.2f} {y:.2f} {w:.2f} {h:.2f} re S".encode())

    def _line(self, x1, y1, x2, y2, color, lw=0.6):
        r, g, b = _hex_rgb(color)
        self.ops.append(f"{r:.3f} {g:.3f} {b:.3f} RG {lw:.2f} w {x1:.2f} {y1:.2f} m {x2:.2f} {y2:.2f} l S".encode())

    def _circle(self, cx, cy, r, fill):
        k = 0.5523 * r
        rr, gg, bb = _hex_rgb(fill)
        p = (f"{rr:.3f} {gg:.3f} {bb:.3f} rg {cx + r:.2f} {cy:.2f} m "
             f"{cx + r:.2f} {cy + k:.2f} {cx + k:.2f} {cy + r:.2f} {cx:.2f} {cy + r:.2f} c "
             f"{cx - k:.2f} {cy + r:.2f} {cx - r:.2f} {cy + k:.2f} {cx - r:.2f} {cy:.2f} c "
             f"{cx - k:.2f} {cy - r:.2f} {cx - r:.2f} {cy - k:.2f} {cx:.2f} {cy - r:.2f} c "
             f"{cx + k:.2f} {cy - r:.2f} {cx + r:.2f} {cy - k:.2f} {cx + r:.2f} {cy:.2f} c f")
        self.ops.append(p.encode())

    def _text(self, x, y, s, size, bold=False, color="#1F2937", italic=False):
        font = "/F2" if bold else ("/F3" if italic else "/F1")
        r, g, b = _hex_rgb(color)
        self.ops.append(f"BT {r:.3f} {g:.3f} {b:.3f} rg {font} {size:.2f} Tf "
                        f"1 0 0 1 {x:.2f} {y:.2f} Tm (".encode()
                        + _pdf_escape(s) + b") Tj ET")

    def _wrap(self, s, size, bold, maxw) -> list[str]:
        s = _pdf_safe(s)
        lines, cur = [], ""
        for word in s.split():
            trial = (cur + " " + word).strip()
            if _text_w(trial, size, bold) <= maxw:
                cur = trial
            else:
                if cur:
                    lines.append(cur)
                while _text_w(word, size, bold) > maxw and len(word) > 1:
                    cut = len(word)
                    while cut > 1 and _text_w(word[:cut], size, bold) > maxw:
                        cut -= 1
                    lines.append(word[:cut])
                    word = word[cut:]
                cur = word
        if cur:
            lines.append(cur)
        return lines or [""]

    # ---------- high level ----------
    def spacer(self, h=8.0):
        self.y -= h

    def heading(self, text, level=1):
        size = 16.0 if level == 1 else 12.5
        self.ensure(size + 16)
        self.spacer(6 if level == 1 else 3)
        if level == 1:
            self._rect(PDF_MARGIN, self.y - size + 1, 4.5, size + 3, BLUE)
            self._text(PDF_MARGIN + 12, self.y - size + 4, text, size, True, "#0B2344")
        else:
            self._text(PDF_MARGIN, self.y - size + 4, text, size, True, "#1D4ED8")
        self.y -= size + 10

    def paragraph(self, text, size=10.0, color="#1F2937", bold=False, indent=0.0):
        maxw = PDF_W - 2 * PDF_MARGIN - indent
        leading = size * 1.45
        for line in self._wrap(text, size, bold, maxw):
            self.ensure(leading)
            self._text(PDF_MARGIN + indent, self.y - size, line, size, bold, color)
            self.y -= leading
        self.spacer(3)

    def bullets(self, items, size=10.0):
        for item in items:
            maxw = PDF_W - 2 * PDF_MARGIN - 14
            lines = self._wrap(item, size, False, maxw)
            self.ensure(size * 1.45)
            self._text(PDF_MARGIN + 2, self.y - size, "-", size, True, BLUE)
            for j, line in enumerate(lines):
                self.ensure(size * 1.45)
                self._text(PDF_MARGIN + 14, self.y - size, line, size, False, "#1F2937")
                self.y -= size * 1.45
            self.spacer(2)

    def kpi_grid(self, pairs):
        gap, cols = 10.0, 3
        cw = (PDF_W - 2 * PDF_MARGIN - (cols - 1) * gap) / cols
        ch = 52.0
        accents = [BLUE, PURPLE, GREEN]
        for i in range(0, len(pairs), cols):
            self.ensure(ch + gap)
            row = pairs[i:i + cols]
            for j, (label, value) in enumerate(row):
                x = PDF_MARGIN + j * (cw + gap)
                y = self.y - ch
                self._rect(x, y, cw, ch, "#EEF4FF", "#C7D7F5")
                self._rect(x, y, 3.5, ch, accents[j % 3])
                self._text(x + 10, y + ch - 16, str(label).upper(), 7.2, True, "#64748B")
                vlines = self._wrap(str(value), 12.5, True, cw - 18)
                self._text(x + 10, y + 12, vlines[0], 12.5, True, "#0B2344")
            self.y -= ch + gap

    def table(self, headers, rows, fractions=None, size=8.5):
        n = len(headers)
        total_w = PDF_W - 2 * PDF_MARGIN
        fractions = fractions or [1.0 / n] * n
        widths = [f * total_w for f in fractions]
        pad, leading = 5.0, size * 1.35

        def draw_header():
            self.ensure(size + 2 * pad + 4)
            h = size + 2 * pad
            self._rect(PDF_MARGIN, self.y - h, total_w, h, "#0B2344")
            x = PDF_MARGIN
            for hdr, w in zip(headers, widths):
                self._text(x + pad, self.y - pad - size, clip(str(hdr), 30), size, True, "#FFFFFF")
                x += w
            self.y -= h

        draw_header()
        for ri, row in enumerate(rows):
            cells = [self._wrap(str(c), size, False, w - 2 * pad) for c, w in zip(row, widths)]
            rh = max(len(c) for c in cells) * leading + 2 * pad
            if self.y - rh < PDF_BOTTOM:
                self._new_page()
                draw_header()
            if ri % 2 == 0:
                self._rect(PDF_MARGIN, self.y - rh, total_w, rh, "#F4F7FC")
            x = PDF_MARGIN
            for cell, w in zip(cells, widths):
                ty = self.y - pad - size
                for line in cell:
                    self._text(x + pad, ty, line, size, False, "#1F2937")
                    ty -= leading
                x += w
            self._line(PDF_MARGIN, self.y - rh, PDF_MARGIN + total_w, self.y - rh, "#D7E0EE", 0.5)
            self.y -= rh
        self.spacer(8)

    def image_png(self, png_bytes: bytes, caption: str = ""):
        from PIL import Image as PILImage
        pil = PILImage.open(io.BytesIO(png_bytes))
        pw, ph = pil.size
        if pil.mode != "RGB":
            pil = pil.convert("RGB")
        buf = io.BytesIO()
        pil.save(buf, format="JPEG", quality=90)
        jpeg = buf.getvalue()
        maxw = PDF_W - 2 * PDF_MARGIN
        w = maxw
        h = w * ph / pw
        if h > 330:
            h = 330.0
            w = h * pw / ph
        self.ensure(h + 24)
        idx = len(self.images)
        self.images.append((jpeg, pw, ph))
        x = PDF_MARGIN + (maxw - w) / 2
        y = self.y - h
        self.ops.append(f"q {w:.2f} 0 0 {h:.2f} {x:.2f} {y:.2f} cm /Im{idx} Do Q".encode())
        self.y = y - 6
        if caption:
            self.paragraph(caption, size=8.5, color="#64748B")
        self.spacer(6)

    # ---------- assembly ----------
    def _chrome(self, ops, page_no, total):
        head = []
        r, g, b = _hex_rgb("#0B2344")
        head.append(f"{r:.3f} {g:.3f} {b:.3f} rg 0 {PDF_H - PDF_HEADER_H:.2f} {PDF_W:.2f} {PDF_HEADER_H:.2f} re f".encode())
        r, g, b = _hex_rgb(BLUE)
        head.append(f"{r:.3f} {g:.3f} {b:.3f} rg 0 {PDF_H - PDF_HEADER_H - 3:.2f} {PDF_W / 2:.2f} 3 re f".encode())
        r, g, b = _hex_rgb(PURPLE)
        head.append(f"{r:.3f} {g:.3f} {b:.3f} rg {PDF_W / 2:.2f} {PDF_H - PDF_HEADER_H - 3:.2f} {PDF_W / 2:.2f} 3 re f".encode())
        saved_ops, saved_y = self.ops, self.y
        self.ops = head
        self._circle(40, PDF_H - 32, 16, BLUE)
        self._text(31.5, PDF_H - 36.5, "BC", 12, True, "#FFFFFF")
        self._text(66, PDF_H - 30, APP_NAME, 15, True, "#FFFFFF")
        self._text(66, PDF_H - 46, self.report_title, 8.5, False, "#9FB3D9")
        stamp = datetime.now().strftime("%d %b %Y")
        right = f"{APP_VERSION} · {stamp}"
        self._text(PDF_W - PDF_MARGIN - _text_w(right, 9, True), PDF_H - 30, right, 9, True, "#9FB3D9")
        self._line(PDF_MARGIN, 44, PDF_W - PDF_MARGIN, 44, "#D7E0EE", 0.7)
        foot = f"Generated {datetime.now().strftime('%d %b %Y %H:%M')} · Source: {clip(self.source, 40)}"
        self._text(PDF_MARGIN, 32, foot, 7.5, False, "#64748B")
        self._text(PDF_W / 2 - 30, 32, "Confidential", 7.5, True, "#94A3B8")
        pn = f"Page {page_no} of {total}"
        self._text(PDF_W - PDF_MARGIN - _text_w(pn, 8, True), 32, pn, 8, True, "#64748B")
        self.ops = saved_ops
        self.y = saved_y
        return head + ops

    def build(self) -> bytes:
        n = len(self.pages)
        contents = [zlib.compress(b"\n".join(self._chrome(p, i + 1, n)))
                    for i, p in enumerate(self.pages)]
        objs: dict[int, bytes] = {}
        kids = " ".join(f"{6 + 2 * i} 0 R" for i in range(n))
        objs[1] = b"<< /Type /Catalog /Pages 2 0 R >>"
        objs[2] = f"<< /Type /Pages /Kids [{kids}] /Count {n} >>".encode()
        objs[3] = b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>"
        objs[4] = b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding >>"
        objs[5] = b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Oblique /Encoding /WinAnsiEncoding >>"
        img_start = 6 + 2 * n
        xobjs = " ".join(f"/Im{i} {img_start + i} 0 R" for i in range(len(self.images)))
        for i, comp in enumerate(contents):
            pnum, cnum = 6 + 2 * i, 7 + 2 * i
            objs[pnum] = (f"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {PDF_W:.2f} {PDF_H:.2f}] "
                          f"/Resources << /Font << /F1 3 0 R /F2 4 0 R /F3 5 0 R >> "
                          f"/XObject << {xobjs} >> >> /Contents {cnum} 0 R >>").encode()
            objs[cnum] = (f"<< /Length {len(comp)} /Filter /FlateDecode >>\nstream\n".encode()
                          + comp + b"\nendstream")
        for i, (jpeg, pw, ph) in enumerate(self.images):
            objs[img_start + i] = (
                f"<< /Type /XObject /Subtype /Image /Width {pw} /Height {ph} "
                f"/ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode "
                f"/Length {len(jpeg)} >>\nstream\n".encode() + jpeg + b"\nendstream")
        out = bytearray(b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n")
        offsets = {}
        for num in sorted(objs):
            offsets[num] = len(out)
            out += f"{num} 0 obj\n".encode() + objs[num] + b"\nendobj\n"
        xref_pos = len(out)
        maxnum = max(objs)
        out += f"xref\n0 {maxnum + 1}\n0000000000 65535 f \n".encode()
        for num in range(1, maxnum + 1):
            out += f"{offsets[num]:010d} 00000 n \n".encode()
        out += (f"trailer\n<< /Size {maxnum + 1} /Root 1 0 R >>\nstartxref\n{xref_pos}\n%%EOF").encode()
        return bytes(out)


def build_pdf_report(df, prof, ana, ins, rca, source_name) -> bytes:
    k = ana["kpis"]
    doc = PDFDoc("Executive Intelligence Report", source_name)
    doc.heading("Executive Intelligence Report")
    doc.paragraph(f"Source: {source_name}   ·   Detected type: {prof['type']}   ·   "
                  f"Records: {k['records']:,}   ·   Period: {k['date_range']}", size=9, color="#64748B")
    doc.heading("Executive Summary", 2)
    doc.paragraph(executive_summary(df, prof, ana), size=10.5)
    doc.heading("Key Performance Indicators", 2)
    doc.kpi_grid([
        ("Total Records", f"{k['records']:,}"), ("Total Columns", k["columns"]),
        ("Missing Data", f"{k['missing_pct']:.2f}%"), ("Duplicate Rows", f"{k['duplicates']:,}"),
        ("Date Range", k["date_range"]), ("Top Category", k["top_category"]),
        ("Primary Metric", f"{k['metric_name']}: {fmt_compact(k['metric_total'])}"),
        ("Quality Score", f"{k['quality']['score']}/100 {k['quality']['grade']}"),
        ("Growth (PoP)", fmt_pct(k["growth_pct"])),
    ])
    doc.heading("Performance Charts", 2)
    if ana["top"] is not None and not ana["top"].empty:
        t = ana["top"]
        doc.image_png(png_bar(t.iloc[:, 0].tolist(), t.iloc[:, 1].tolist(),
                              f"Top {t.columns[0]} by {t.columns[1]}"),
                      f"Figure 1 · Leading {t.columns[0]} segments by {t.columns[1]}.")
    if ana["trend"] is not None:
        doc.image_png(png_line(ana["trend"].index, ana["trend"].values,
                               f"{ana['trend_label']} {ana['trend'].name}"),
                      f"Figure 2 · {ana['trend_label']} trajectory of {ana['trend'].name}.")
    if ana["top"] is not None and len(ana["top"]) >= 2:
        t = ana["top"].head(6)
        doc.image_png(png_donut(t.iloc[:, 0].tolist(), t.iloc[:, 1].tolist(),
                                f"{t.columns[1]} share"),
                      "Figure 3 · Composition of the primary metric (top 6 segments).")
    if ana["top"] is not None and not ana["top"].empty:
        doc.heading("Top Contributors", 2)
        t = ana["top"]
        total = float(t.iloc[:, 1].sum())
        rows = [[r.iloc[0], fmt_compact(r.iloc[1]),
                 f"{r.iloc[1] / total * 100:.1f}%" if total > 0 else "—"]
                for _, r in t.iterrows()]
        doc.table([str(t.columns[0]), str(t.columns[1]), "Share"], rows, [.46, .3, .24])
    for title, key in (("What Happened", "what_happened"), ("Risk Analysis", "risks"),
                       ("Opportunities", "opportunities")):
        doc.heading(title, 2)
        doc.bullets([f"{i['title']}: {_strip_md(i['detail'])}" for i in ins[key]])
    doc.heading("Recommendations", 2)
    doc.bullets([f"{i['title']}: {_strip_md(i['detail'])}" for i in ins["recommendations"]])
    if rca and rca.get("available"):
        doc.heading("Root Cause Analysis", 2)
        for nline in rca["narrative"]:
            doc.paragraph(_strip_md(nline), size=9.5)
        if rca["drivers"]:
            doc.image_png(png_drivers(rca["drivers"], "Change drivers (period over period)"),
                          "Figure 4 · Largest segment-level contributors to the change.")
            doc.table(["Segment", "Dimension", "Previous", "Current", "Change", "Contrib."],
                      [[d["segment"], d["dim"], fmt_compact(d["prev"]), fmt_compact(d["curr"]),
                        fmt_compact(d["delta"]), f"{d['contrib']:.0f}%"] for d in rca["drivers"]],
                      [.24, .16, .15, .15, .15, .15])
    doc.heading("Data Quality · Column Profile", 2)
    prof_tab = column_profile(df)
    doc.table(list(prof_tab.columns),
              prof_tab.head(18).values.tolist(), [.24, .16, .12, .12, .12, .24])
    return doc.build()


# =====================================================================
# POWERPOINT GENERATION (python-pptx)
# =====================================================================
def _pptx_slide(prs, title, subtitle=""):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = RGBColor.from_string(BG0.lstrip("#"))
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(0.14), prs.slide_height)
    bar.fill.solid()
    bar.fill.fore_color.rgb = RGBColor.from_string(BLUE.lstrip("#"))
    bar.line.fill.background()
    tb = slide.shapes.add_textbox(Inches(0.55), Inches(0.32), Inches(12.3), Inches(0.85))
    p = tb.text_frame.paragraphs[0]
    p.text = title
    p.font.size, p.font.bold = Pt(30), True
    p.font.color.rgb = RGBColor.from_string(TEXT.lstrip("#"))
    if subtitle:
        sb = slide.shapes.add_textbox(Inches(0.57), Inches(1.12), Inches(12.2), Inches(0.5))
        sp = sb.text_frame.paragraphs[0]
        sp.text = subtitle
        sp.font.size = Pt(13)
        sp.font.color.rgb = RGBColor.from_string(MUTED.lstrip("#"))
    return slide


def _pptx_bullets(slide, items, top=1.75, size=15):
    tb = slide.shapes.add_textbox(Inches(0.6), Inches(top), Inches(12.1), Inches(5.3))
    tf = tb.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = "•  " + clip(_strip_md(item), 210)
        p.font.size = Pt(size)
        p.font.color.rgb = RGBColor.from_string("E2E8F0")
        p.space_after = Pt(10)


def build_pptx_report(df, prof, ana, ins, rca, source_name) -> bytes:
    k = ana["kpis"]
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    stamp = datetime.now().strftime("%d %b %Y")

    s = _pptx_slide(prs, "", "")
    band = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(2.0), prs.slide_width, Inches(2.6))
    band.fill.solid()
    band.fill.fore_color.rgb = RGBColor.from_string(BG2.lstrip("#"))
    band.line.fill.background()
    logo = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.7), Inches(2.5), Inches(1.5), Inches(1.5))
    logo.fill.solid()
    logo.fill.fore_color.rgb = RGBColor.from_string(BLUE.lstrip("#"))
    logo.line.fill.background()
    lp = logo.text_frame.paragraphs[0]
    lp.text, lp.font.size, lp.font.bold = "BC", Pt(40), True
    lp.font.color.rgb = RGBColor.from_string("FFFFFF")
    lp.alignment = PP_ALIGN.CENTER
    tb = s.shapes.add_textbox(Inches(2.5), Inches(2.45), Inches(10.3), Inches(1.1))
    p = tb.text_frame.paragraphs[0]
    p.text = f"{APP_NAME} {APP_VERSION}"
    p.font.size, p.font.bold = Pt(44), True
    p.font.color.rgb = RGBColor.from_string(TEXT.lstrip("#"))
    tb2 = s.shapes.add_textbox(Inches(2.55), Inches(3.55), Inches(10.2), Inches(0.9))
    p2 = tb2.text_frame.paragraphs[0]
    p2.text = f"Executive Intelligence Deck · {prof['type']}"
    p2.font.size = Pt(18)
    p2.font.color.rgb = RGBColor.from_string(MUTED.lstrip("#"))
    meta = s.shapes.add_textbox(Inches(0.7), Inches(5.2), Inches(12), Inches(1))
    mp = meta.text_frame.paragraphs[0]
    mp.text = f"Source: {source_name}   ·   {k['records']:,} records   ·   {k['date_range']}   ·   {stamp}"
    mp.font.size = Pt(12)
    mp.font.color.rgb = RGBColor.from_string(MUTED.lstrip("#"))

    s = _pptx_slide(prs, "Executive Summary", f"{prof['type']} · {source_name}")
    summary = executive_summary(df, prof, ana)
    sentences = [x.strip() for x in re.split(r"(?<=\.)\s+", summary) if x.strip()]
    _pptx_bullets(s, sentences, top=1.8, size=16)

    s = _pptx_slide(prs, "Key Performance Indicators", "Snapshot of the current view")
    cards = [("Total Records", f"{k['records']:,}"), ("Total Columns", str(k["columns"])),
             ("Missing Data", f"{k['missing_pct']:.2f}%"), ("Duplicates", f"{k['duplicates']:,}"),
             ("Date Range", k["date_range"]), ("Top Category", k["top_category"]),
             (k["metric_name"], fmt_compact(k["metric_total"])),
             ("Quality Score", f"{k['quality']['score']}/100"), ("Growth (PoP)", fmt_pct(k["growth_pct"]))]
    accents = [BLUE, PURPLE, GREEN]
    for i, (label, value) in enumerate(cards):
        r, c = divmod(i, 3)
        shape = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                   Inches(0.6 + c * 4.25), Inches(1.75 + r * 1.8),
                                   Inches(3.95), Inches(1.55))
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor.from_string(BG2.lstrip("#"))
        shape.line.color.rgb = RGBColor.from_string(accents[c].lstrip("#"))
        shape.line.width = Pt(1.2)
        tf = shape.text_frame
        tf.word_wrap = True
        p1 = tf.paragraphs[0]
        p1.text = label.upper()
        p1.font.size, p1.font.bold = Pt(11), True
        p1.font.color.rgb = RGBColor.from_string(MUTED.lstrip("#"))
        p2 = tf.add_paragraph()
        p2.text = clip(value, 26)
        p2.font.size, p2.font.bold = Pt(21), True
        p2.font.color.rgb = RGBColor.from_string(TEXT.lstrip("#"))

    s = _pptx_slide(prs, "Performance Charts", "Auto-generated from the uploaded data")
    if ana["top"] is not None and not ana["top"].empty:
        t = ana["top"]
        s.shapes.add_picture(io.BytesIO(png_bar(t.iloc[:, 0].tolist(), t.iloc[:, 1].tolist(),
                                                f"Top {t.columns[0]} by {t.columns[1]}")),
                             Inches(0.55), Inches(1.8), width=Inches(6.1))
    if ana["trend"] is not None:
        s.shapes.add_picture(io.BytesIO(png_line(ana["trend"].index, ana["trend"].values,
                                                 f"{ana['trend_label']} {ana['trend'].name}")),
                             Inches(6.85), Inches(1.8), width=Inches(6.1))
    elif ana["top"] is not None and len(ana["top"]) >= 2:
        t = ana["top"].head(6)
        s.shapes.add_picture(io.BytesIO(png_donut(t.iloc[:, 0].tolist(), t.iloc[:, 1].tolist(),
                                                  f"{t.columns[1]} share")),
                             Inches(6.85), Inches(1.8), width=Inches(6.1))

    s = _pptx_slide(prs, "Insights · What Happened", "Deterministic AI insights engine")
    _pptx_bullets(s, [f"{i['title']}: {i['detail']}" for i in ins["what_happened"][:6]], size=14)

    s = _pptx_slide(prs, "Recommendations", "Prioritized management actions")
    _pptx_bullets(s, [f"{i['title']}: {i['detail']}" for i in ins["recommendations"][:6]], size=14)

    s = _pptx_slide(prs, "Risk Analysis", "Quality, concentration, outliers and decline")
    _pptx_bullets(s, [f"{i['title']}: {i['detail']}" for i in ins["risks"][:6]], size=14)

    s = _pptx_slide(prs, "Next Steps", "From insight to execution")
    steps = [f"{i['title'].split('·')[-1].strip()}: {clip(_strip_md(i['detail']), 150)}"
             for i in ins["recommendations"][:4]]
    steps.append("Assign an owner and deadline to every P1 action this week.")
    steps.append(f"Review {k['metric_name']} weekly; re-run Root Cause Analysis monthly.")
    _pptx_bullets(s, steps, size=14)

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


# =====================================================================
# EMAIL SUMMARY & EXCEL EXPORT
# =====================================================================
def build_email_summary(df, prof, ana, ins, source_name) -> str:
    k, g = ana["kpis"], ana["growth"]
    growth_txt = fmt_pct(k["growth_pct"]) if k["growth_pct"] is not None else "n/a"
    subject = (f"{prof['type']} Brief — {k['metric_name']} {fmt_compact(k['metric_total'])} "
               f"| {growth_txt} PoP | {APP_NAME}")
    lines = [f"Subject: {subject}", "", "Hi team,", "",
             "EXECUTIVE OVERVIEW", executive_summary(df, prof, ana), "", "KEY INSIGHTS"]
    lines += [f"- {i['title']}: {_strip_md(i['detail'])}" for i in ins["what_happened"][:3]]
    lines += ["", "RISKS"]
    lines += [f"- {i['title']}: {_strip_md(i['detail'])}" for i in ins["risks"][:3]]
    lines += ["", "RECOMMENDATIONS"]
    lines += [f"- {i['title']}: {_strip_md(i['detail'])}" for i in ins["recommendations"][:4]]
    lines += ["", f"Source: {source_name} · {k['records']:,} records · {k['date_range']}",
              "", "Best regards,", f"{APP_NAME} {APP_VERSION} · Automated Executive Brief"]
    return "\n".join(lines)


def build_excel_export(df, prof, ana, ins, rca, source_name) -> bytes:
    k = ana["kpis"]
    export_df = df.iloc[:MAX_EXCEL_ROWS] if len(df) > MAX_EXCEL_ROWS else df
    kpi_df = pd.DataFrame({
        "KPI": ["Total Records", "Total Columns", "Missing Data %", "Duplicate Rows",
                "Date Range", "Top Category", "Primary Metric", "Primary Metric Total",
                "Data Quality Score", "Quality Grade", "Growth % (PoP)", "Detected Type", "Source"],
        "Value": [k["records"], k["columns"], round(k["missing_pct"], 2), k["duplicates"],
                  k["date_range"], k["top_category"], k["metric_name"],
                  round(k["metric_total"], 2), k["quality"]["score"], k["quality"]["grade"],
                  "n/a" if k["growth_pct"] is None else round(k["growth_pct"], 2),
                  prof["type"], source_name]})
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        export_df.to_excel(writer, index=False, sheet_name="Data")
        kpi_df.to_excel(writer, index=False, sheet_name="KPIs")
        pd.DataFrame({"Executive Summary": [executive_summary(df, prof, ana)]}).to_excel(
            writer, index=False, sheet_name="Executive Summary")
        for sheet, key in (("What Happened", "what_happened"), ("Risks", "risks"),
                           ("Opportunities", "opportunities"), ("Recommendations", "recommendations")):
            pd.DataFrame([{"Title": i["title"], "Severity": i["sev"],
                           "Detail": _strip_md(i["detail"])} for i in ins[key]]).to_excel(
                writer, index=False, sheet_name=sheet)
        if rca and rca.get("available"):
            pd.DataFrame({"Root Cause Narrative": [_strip_md(n) for n in rca["narrative"]]}).to_excel(
                writer, index=False, sheet_name="Root Cause")
            pd.DataFrame(rca["drivers"]).to_excel(writer, index=False, sheet_name="RCA Drivers")
        if ana["top"] is not None:
            ana["top"].to_excel(writer, index=False, sheet_name="Top Contributors")
        column_profile(df).to_excel(writer, index=False, sheet_name="Column Profile")
    buf.seek(0)
    return buf.getvalue()


# =====================================================================
# UI COMPONENTS
# =====================================================================
SEV_COLORS = {"info": BLUE, "success": GREEN, "warning": ORANGE, "critical": RED,
              "opportunity": GREEN, "action": PURPLE}


def _md_to_html(text: str) -> str:
    return re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", esc(text))


def kpi_card(icon, label, value, sub="", tone="") -> str:
    return (f'<div class="kpi"><div class="kpi-icon">{icon}</div>'
            f'<div class="kpi-label">{esc(label)}</div>'
            f'<div class="kpi-value {tone}">{esc(value)}</div>'
            f'<div class="kpi-sub">{esc(sub)}</div></div>')


def section_header(title, color=BLUE):
    st.markdown(f'<div class="section"><div class="section-bar" style="background:{color};">'
                f'</div><div class="section-title">{esc(title)}</div></div>',
                unsafe_allow_html=True)


def insight_card(item) -> str:
    color = SEV_COLORS.get(item["sev"], BLUE)
    return (f'<div class="insight" style="border-left-color:{color};">'
            f'<span class="sev" style="background:{color}22;color:{color};">{esc(item["sev"])}</span>'
            f'<div class="insight-title">{item["icon"]} {esc(item["title"])}</div>'
            f'<div class="insight-body">{_md_to_html(item["detail"])}</div></div>')


def render_insight_grid(items, cols=2):
    grid = st.columns(cols)
    for i, item in enumerate(items):
        with grid[i % cols]:
            st.markdown(insight_card(item), unsafe_allow_html=True)


def render_hero(prof=None, source=None):
    st.markdown(f'<div class="hero-title">🚀 {APP_NAME} {APP_VERSION}</div>',
                unsafe_allow_html=True)
    st.markdown('<div class="hero-sub">Enterprise AI-powered Business Intelligence — upload any '
                'Excel or CSV and instantly get an executive dashboard, AI insights, root-cause '
                'analysis, chat with your data, and board-ready PDF / PowerPoint reports.</div>',
                unsafe_allow_html=True)
    chips = ["Auto Dashboard", "AI Insights", "Root Cause Analysis", "Chat with Data",
             "PDF & PPTX Reports", "Email Brief", "13+ Auto Charts"]
    st.markdown("".join(f'<span class="chip">{esc(c)}</span>' for c in chips),
                unsafe_allow_html=True)
    if prof is not None:
        st.markdown(f'<div style="margin-top:14px;"><span class="badge">🗂️ {esc(prof["type"])}</span>'
                    f'&nbsp;&nbsp;<span class="chip">📄 {esc(source)}</span></div>',
                    unsafe_allow_html=True)
    st.write("")


def render_kpis(ana):
    k = ana["kpis"]
    growth = k["growth_pct"]
    g_tone = "" if growth is None else ("pos" if growth >= 0 else "neg")
    g_icon = "📈" if growth is not None and growth >= 0 else "📉"
    cards = [
        ("📦", "Total Records", f"{k['records']:,}", "rows in view", ""),
        ("🧱", "Total Columns", f"{k['columns']:,}", "fields detected", ""),
        ("🧹", "Missing Data", f"{k['missing_pct']:.1f}%", "of all cells",
         "neg" if k["missing_pct"] >= 5 else ""),
        ("👯", "Duplicate Rows", f"{k['duplicates']:,}", "exact duplicates",
         "neg" if k["duplicates"] > 0 else ""),
        ("📅", "Date Range", k["date_range"], "coverage", ""),
        ("🏆", "Top Category", k["top_category"], "most frequent", ""),
        ("💰", "Primary Metric", fmt_compact(k["metric_total"]), k["metric_name"], ""),
        ("🛡️", "Quality Score", f"{k['quality']['score']:.0f}/100", k["quality"]["grade"],
         "pos" if k["quality"]["score"] >= 75 else "neg"),
        (g_icon, "Growth %", fmt_pct(growth), "period over period", g_tone),
    ]
    for ui_col, card in zip(st.columns(5), cards[:5]):
        with ui_col:
            st.markdown(kpi_card(*card), unsafe_allow_html=True)
    st.write("")
    for ui_col, card in zip(st.columns(4), cards[5:]):
        with ui_col:
            st.markdown(kpi_card(*card), unsafe_allow_html=True)


# =====================================================================
# SIDEBAR — upload, samples, AI engines, global filters
# =====================================================================
def sidebar_controls():
    with st.sidebar:
        st.markdown("## ⚙️ Control Panel")
        uploaded = st.file_uploader("📤 Upload Excel or CSV", type=["csv", "xlsx", "xls"],
                                    help="Processed in memory only — never stored.")
        sample_kind = st.selectbox("🎲 Sample domain", SAMPLE_KINDS, key="sample_kind")
        if st.button("🎲 Load Sample Dataset"):
            st.session_state["use_sample"] = True
            st.session_state["sample_active"] = sample_kind
        st.divider()
        st.markdown("### 🧠 AI Engine")
        openai_key = st.text_input("OpenAI API key (optional)", type="password", key="openai_key")
        openai_model = st.selectbox("OpenAI model", OPENAI_MODELS, key="openai_model")
        ollama_url = st.text_input("Ollama URL", "http://localhost:11434", key="ollama_url")
        if st.button("🔄 Refresh Ollama models"):
            try:
                st.session_state["ollama_models"] = fetch_ollama_models(ollama_url)
                st.success(f"Found {len(st.session_state['ollama_models'])} model(s).")
            except Exception as err:
                st.warning(f"Ollama not reachable — {clip(err, 90)}")
        models = list(dict.fromkeys(DEFAULT_OLLAMA_MODELS + st.session_state.get("ollama_models", [])))
        ollama_model = st.selectbox("Ollama model", models, index=0, key="ollama_model")
        st.caption("Priority: **OpenAI** (if key) → **Ollama** → **Built-in Analyst**.")
        st.divider()
        st.caption("🔒 Privacy: files are processed in memory and never stored on a server.")
    return uploaded, {"openai_key": openai_key.strip(), "openai_model": openai_model,
                      "ollama_url": ollama_url.strip() or "http://localhost:11434",
                      "ollama_model": ollama_model}


def apply_global_filters(df, prof):
    with st.sidebar.expander("🌍 Global Filters", expanded=False):
        fdf = df
        date_col = prof["primary_date"]
        if date_col is not None and df[date_col].notna().any():
            lo, hi = df[date_col].min().date(), df[date_col].max().date()
            if lo < hi:
                rng = st.date_input("Date range", (lo, hi), min_value=lo, max_value=hi,
                                    key="gf_date")
                if isinstance(rng, (tuple, list)) and len(rng) == 2:
                    fdf = fdf[fdf[date_col].isna()
                              | fdf[date_col].dt.date.between(rng[0], rng[1])]
        for dim in prof["dimensions"][:6]:
            opts = sorted(df[dim].dropna().unique().tolist(), key=str)
            chosen = st.multiselect(str(dim), opts, default=[], format_func=str,
                                    key=f"gf_dim_{dim}")
            if chosen:
                fdf = fdf[fdf[dim].isin(chosen)]
        for m in prof["metrics"][:3]:
            vals = df[m].dropna()
            if not vals.empty:
                lo, hi = float(vals.min()), float(vals.max())
                if lo < hi:
                    r = st.slider(f"{m} range", lo, hi, (lo, hi), key=f"gf_num_{m}")
                    fdf = fdf[fdf[m].isna() | fdf[m].between(r[0], r[1])]
        tcol = st.selectbox("Text filter column", ["(none)"] + list(df.columns), key="gf_textcol")
        tval = st.text_input("Contains text", "", key="gf_textval").strip()
        if tcol != "(none)" and tval:
            fdf = fdf[fdf[tcol].astype(str).str.contains(tval, case=False, regex=False, na=False)]
        search = st.text_input("🔎 Global search (all columns)", "", key="gf_search").strip()
        if search:
            mask = fdf.astype(str).apply(
                lambda s: s.str.contains(search, case=False, regex=False, na=False)).any(axis=1)
            fdf = fdf[mask]
        if st.button("♻️ Reset all filters"):
            for key in list(st.session_state):
                if key.startswith("gf_"):
                    del st.session_state[key]
            st.rerun()
    return fdf


# =====================================================================
# TAB RENDERERS
# =====================================================================
def render_dashboard(df, prof, ana, ins):
    section_header("🧠 Executive Summary", GREEN)
    st.markdown(f'<div class="glass">{esc(executive_summary(df, prof, ana))}</div>',
                unsafe_allow_html=True)
    section_header("📈 Performance Charts", BLUE)
    c1, c2 = st.columns(2)
    with c1:
        fig = chart_top_bar(df, prof)
        if fig is not None:
            st.plotly_chart(fig, width="stretch", config={"displaylogo": False})
    with c2:
        fig = chart_share(df, prof, hole=.55, by_metric=True)
        if fig is not None:
            st.plotly_chart(fig, width="stretch", config={"displaylogo": False})
    fig = chart_trend(df, prof)
    if fig is not None:
        st.plotly_chart(fig, width="stretch", config={"displaylogo": False})
    else:
        st.info("ℹ️ No usable date column — trend charts activate automatically when dates exist.")
    section_header("💡 Key Insights (preview)", PURPLE)
    render_insight_grid(ins["what_happened"][:2] + ins["risks"][:2])
    st.caption("Full engine → 🤖 AI Insights tab · Drivers → 🔎 Root Cause Analysis tab.")


def render_insights_tab(ins):
    for title, key, color in (("📈 What Happened", "what_happened", BLUE),
                              ("⚠️ Risks", "risks", RED),
                              ("🌱 Opportunities", "opportunities", GREEN),
                              ("🎯 Recommendations", "recommendations", PURPLE)):
        section_header(title, color)
        render_insight_grid(ins[key])


def _rca_sig(df, ana):
    return (len(df), round(ana["kpis"]["metric_total"], 2),
            ana["growth"]["prev_label"], ana["growth"]["curr_label"])


def render_rca_tab(df, prof, ana):
    st.caption("Decomposes the latest period-over-period change across Date, Category, Region, "
               "Department, Project, Status, Product, Channel, Location and every other detected "
               "dimension — then explains the drivers in plain business language.")
    sig = _rca_sig(df, ana)
    if st.button("🔎 Analyze Root Cause", key="rca_run"):
        with st.spinner("Decomposing segment drivers…"):
            st.session_state["rca_result"] = run_root_cause(df, prof, ana)
            st.session_state["rca_sig"] = sig
    res = st.session_state.get("rca_result")
    if not res or st.session_state.get("rca_sig") != sig:
        st.info("Click **Analyze Root Cause** to explain the latest change.")
        return
    if not res.get("available"):
        st.info(f"Root cause analysis unavailable: {res.get('reason', 'insufficient history')}")
        return
    g = res["growth"]
    m1, m2, m3, m4 = st.columns(4)
    m1.metric(f"Previous · {g['prev_label']}", fmt_compact(g["prev_total"]))
    m2.metric(f"Current · {g['curr_label']}", fmt_compact(g["curr_total"]),
              delta=fmt_pct(g["pct"]))
    m3.metric("Absolute Change", fmt_compact(g["delta"]))
    m4.metric("Volume Change", fmt_pct(res["volume_pct"]))
    wf = build_waterfall_fig(res)
    if wf is not None:
        st.plotly_chart(wf, width="stretch", config={"displaylogo": False})
    section_header("🧾 Explanation", GREEN)
    st.markdown('<div class="glass">' + "<br><br>".join(_md_to_html(n) for n in res["narrative"])
                + "</div>", unsafe_allow_html=True)
    if res["drivers"]:
        section_header("🚗 Driver Table", BLUE)
        drv = pd.DataFrame([{"Segment": d["segment"], "Dimension": d["dim"],
                             "Previous": round(d["prev"], 2), "Current": round(d["curr"], 2),
                             "Change": round(d["delta"], 2),
                             "Contribution %": round(d["contrib"], 1)} for d in res["drivers"]])
        st.dataframe(drv, width="stretch", hide_index=True)
    section_header("🧭 Dimension Deep-Dives", PURPLE)
    for d in res["dimensions"]:
        with st.expander(f"{d['dim']} — explained change {fmt_compact(d['score'])}"):
            tab = d["table"].copy()
            tab["Direction"] = np.where(tab["Change"] >= 0, "Gain", "Loss")
            fig = px.bar(tab.head(12), x=str(d["dim"]), y="Change", color="Direction",
                         color_discrete_map={"Gain": GREEN, "Loss": RED})
            st.plotly_chart(style_fig(fig, f"{g['metric']} change by {d['dim']}"),
                            width="stretch", config={"displaylogo": False})
            show = tab.drop(columns=["Direction"]).copy()
            for col in ("Previous", "Current", "Change"):
                show[col] = show[col].round(2)
            show["Contribution %"] = show["Contribution %"].round(1)
            show["Growth %"] = show["Growth %"].round(1)
            st.dataframe(show, width="stretch", hide_index=True)


def render_chat_tab(df, prof, ana, ins, settings):
    st.caption("Ask in plain English. Engine priority: OpenAI (if key supplied) → Ollama → "
               "Built-in Analyst. The assistant sees your KPIs, insights and root-cause context.")
    if "chat_history" not in st.session_state:
        st.session_state["chat_history"] = []
    examples = ["Why did revenue decline?", "What is the biggest risk?",
                "Who is the top contributor?", "Show trends by region"]
    cols = st.columns(4)
    for i, (col, ex) in enumerate(zip(cols, examples)):
        with col:
            if st.button(ex, key=f"chat_q{i}"):
                st.session_state["chat_pending"] = ex
    for msg in st.session_state["chat_history"]:
        with st.chat_message(msg["role"]):
            st.markdown(msg["text"])
            if msg.get("fig") is not None:
                st.plotly_chart(msg["fig"], width="stretch", config={"displaylogo": False})
            if msg.get("engine"):
                st.caption(f"🤖 {msg['engine']}")
    question = st.chat_input("Ask anything about your data…")
    pending = st.session_state.pop("chat_pending", None)
    question = question or pending
    if question:
        st.session_state["chat_history"].append({"role": "user", "text": question})
        with st.chat_message("user"):
            st.markdown(question)
        with st.chat_message("assistant"):
            with st.spinner("Analyzing…"):
                rca = st.session_state.get("rca_result")
                answer = answer_question(question, df, prof, ana, ins, settings, rca)
            st.markdown(answer["text"])
            if answer.get("fig") is not None:
                st.plotly_chart(answer["fig"], width="stretch", config={"displaylogo": False})
            st.caption(f"🤖 {answer['engine']}")
            if answer.get("fallback_note"):
                st.caption("⚠️ " + answer["fallback_note"])
        st.session_state["chat_history"].append(
            {"role": "assistant", "text": answer["text"], "fig": answer.get("fig"),
             "engine": answer["engine"]})
    if st.session_state["chat_history"] and st.button("🗑️ Clear chat", key="chat_clear"):
        st.session_state["chat_history"] = []
        st.rerun()


def render_gallery_tab(df, prof, ana):
    charts = auto_charts(df, prof, ana)
    st.caption(f"{len(charts)} chart types auto-selected from your data's column profile.")
    for i in range(0, len(charts), 2):
        cols = st.columns(2)
        for col, spec in zip(cols, charts[i:i + 2]):
            with col:
                st.plotly_chart(spec["fig"], width="stretch", config={"displaylogo": False})
                st.caption(f"**{spec['title']}** — {spec['why']}")


def render_explorer_tab(df, prof):
    st.caption(f"{len(df):,} rows × {df.shape[1]} columns (global filters applied).")
    st.dataframe(df, width="stretch", hide_index=True, height=440)
    section_header("🔬 Column Analysis", BLUE)
    col = st.selectbox("Choose a column", list(df.columns), key="analysis_col")
    s = df[col]
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Type", str(s.dtype))
    m2.metric("Non-Null", f"{int(s.notna().sum()):,}")
    m3.metric("Missing %", f"{s.isna().mean() * 100:.1f}%" if len(s) else "0%")
    m4.metric("Unique", f"{int(s.nunique(dropna=True)):,}")
    if len(df) == 0:
        st.info("No rows in the current view.")
    elif pd.api.types.is_datetime64_any_dtype(s):
        series, label, _ = get_trend(df, col, None)
        if series is not None:
            tdf = series.rename_axis(col).reset_index()
            fig = px.line(tdf, x=col, y=series.name, markers=True)
            st.plotly_chart(style_fig(fig, f"{label} record count"), width="stretch",
                            config={"displaylogo": False})
        else:
            st.caption(f"Range: {s.min()} → {s.max()}")
    elif pd.api.types.is_numeric_dtype(s) and not pd.api.types.is_bool_dtype(s):
        fig = px.histogram(df, x=col, nbins=30)
        st.plotly_chart(style_fig(fig, f"Distribution of {col}"), width="stretch",
                        config={"displaylogo": False})
        desc = s.describe().rename("Value").reset_index().rename(columns={"index": "Statistic"})
        st.dataframe(desc, width="stretch", hide_index=True)
    else:
        counts = s.dropna().astype(str).value_counts().head(15)
        if counts.empty:
            st.info("This column has no non-null values to analyse.")
        else:
            vc = counts.rename_axis(col).reset_index(name="Count")
            fig = px.bar(vc, x=col, y="Count", color="Count",
                         color_continuous_scale="Blues", text_auto=True)
            fig.update_layout(coloraxis_showscale=False)
            fig.update_xaxes(categoryorder="total descending")
            st.plotly_chart(style_fig(fig, f"Top values — {col}"), width="stretch",
                            config={"displaylogo": False})
    with st.expander("🧾 Full column profile"):
        st.dataframe(column_profile(df), width="stretch", hide_index=True)


def _report_sig(source, df, ana):
    return (source, len(df), round(ana["kpis"]["metric_total"], 2))


def _pdf_pptx_buttons(df, prof, ana, ins, rca, source, sig, key_prefix):
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("#### 📕 PDF Report")
        if st.button("🛠️ Build PDF Report", key=f"{key_prefix}_pdf_build"):
            with st.spinner("Rendering PDF (charts, tables, insights)…"):
                try:
                    st.session_state["pdf_bytes"] = build_pdf_report(df, prof, ana, ins, rca, source)
                    st.session_state["pdf_sig"] = sig
                except Exception as err:
                    st.error(f"PDF build failed: {err}")
        if st.session_state.get("pdf_bytes") and st.session_state.get("pdf_sig") == sig:
            st.download_button("⬇️ Download PDF", st.session_state["pdf_bytes"],
                               file_name="business_copilot_report.pdf", mime="application/pdf",
                               key=f"{key_prefix}_pdf_dl")
            st.caption("Title & logo area · summary · KPIs · charts · insights · tables.")
        else:
            st.caption("Click **Build PDF Report** to render the branded document.")
    with c2:
        st.markdown("#### 📙 PowerPoint Deck")
        if st.button("🛠️ Build PowerPoint", key=f"{key_prefix}_pptx_build"):
            with st.spinner("Building 8-slide executive deck…"):
                try:
                    st.session_state["pptx_bytes"] = build_pptx_report(df, prof, ana, ins, rca, source)
                    st.session_state["pptx_sig"] = sig
                except Exception as err:
                    st.error(f"PowerPoint build failed: {err}")
        if st.session_state.get("pptx_bytes") and st.session_state.get("pptx_sig") == sig:
            st.download_button("⬇️ Download .pptx", st.session_state["pptx_bytes"],
                               file_name="business_copilot_deck.pptx",
                               mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                               key=f"{key_prefix}_pptx_dl")
            st.caption("Title · Summary · KPIs · Charts · Insights · Recommendations · Risks · Next Steps.")
        else:
            st.caption("Click **Build PowerPoint** to generate the deck.")


def render_reports_tab(df, prof, ana, ins, source):
    rca = st.session_state.get("rca_result") or run_root_cause(df, prof, ana)
    reports = build_text_reports(df, prof, ana, ins, rca, source)
    choice = st.radio("Report type", list(reports), horizontal=True, key="report_choice")
    st.markdown(reports[choice])
    slug = re.sub(r"\W+", "_", choice.lower()).strip("_")
    st.download_button(f"⬇️ Download {choice} (.md)", reports[choice].encode("utf-8"),
                       file_name=f"{slug}.md", mime="text/markdown", key="report_md_dl")
    st.divider()
    _pdf_pptx_buttons(df, prof, ana, ins, rca, source, _report_sig(source, df, ana), "rep")
    st.divider()
    section_header("📧 Executive Email", ORANGE)
    email = build_email_summary(df, prof, ana, ins, source)
    st.code(email, language="text")
    st.download_button("⬇️ Download email (.txt)", email.encode("utf-8"),
                       file_name="executive_email.txt", mime="text/plain", key="email_dl")
    st.caption("Copy straight into Outlook or Gmail — subject line included.")


def render_export_tab(df, prof, ana, ins, source):
    sig = _report_sig(source, df, ana)
    st.caption(f"**{len(df):,} rows × {df.shape[1]} columns** ready for export "
               "(global filters applied).")
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("#### 📄 CSV")
        st.download_button("⬇️ Download CSV", df.to_csv(index=False).encode("utf-8-sig"),
                           file_name="business_copilot_data.csv", mime="text/csv", key="exp_csv")
        st.caption("UTF-8 with BOM — opens cleanly in Excel.")
        st.markdown("#### 🧠 Insights Report")
        rca = st.session_state.get("rca_result") or run_root_cause(df, prof, ana)
        reports = build_text_reports(df, prof, ana, ins, rca, source)
        st.download_button("⬇️ Download Insights Report (.md)",
                           reports["Management Report"].encode("utf-8"),
                           file_name="insights_report.md", mime="text/markdown", key="exp_md")
    with c2:
        st.markdown("#### 📊 Excel Workbook")
        if st.button("🛠️ Build Excel Workbook", key="exp_xlsx_build"):
            with st.spinner("Writing multi-sheet workbook…"):
                try:
                    rca = st.session_state.get("rca_result") or run_root_cause(df, prof, ana)
                    st.session_state["xlsx_bytes"] = build_excel_export(df, prof, ana, ins, rca, source)
                    st.session_state["xlsx_sig"] = sig
                except Exception as err:
                    st.error(f"Excel build failed: {err}")
        if len(df) > MAX_EXCEL_ROWS:
            st.warning(f"Excel limit: data sheet will be truncated to {MAX_EXCEL_ROWS:,} rows.")
        if st.session_state.get("xlsx_bytes") and st.session_state.get("xlsx_sig") == sig:
            st.download_button("⬇️ Download .xlsx", st.session_state["xlsx_bytes"],
                               file_name="business_copilot_workbook.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               key="exp_xlsx_dl")
            st.caption("Data · KPIs · Summary · Insights · Risks · Opportunities · "
                       "Recommendations · Root Cause · Drivers · Top Contributors · Profile.")
        else:
            st.caption("Click **Build Excel Workbook** to prepare all sheets.")
    st.divider()
    _pdf_pptx_buttons(df, prof, ana, ins,
                      st.session_state.get("rca_result") or run_root_cause(df, prof, ana),
                      source, sig, "exp")


# =====================================================================
# MAIN
# =====================================================================
def main():
    uploaded, settings = sidebar_controls()
    df_raw, source = None, "dataset.csv"

    if uploaded is not None:
        st.session_state["use_sample"] = False
        upload_sig = (uploaded.name, uploaded.size)
        if st.session_state.get("last_upload") != upload_sig:
            st.session_state["last_upload"] = upload_sig
            for key in list(st.session_state):
                if key.startswith(("gf_", "rca_", "chat_", "pdf_", "pptx_", "xlsx_")):
                    del st.session_state[key]
        file_bytes = uploaded.getvalue()
        sheets = get_sheet_names(file_bytes, uploaded.name)
        sheet = (st.sidebar.selectbox("📄 Excel sheet", sheets)
                 if len(sheets) > 1 else (sheets[0] if sheets else None))
        try:
            df_raw = load_data(file_bytes, uploaded.name, sheet)
        except Exception as err:
            render_hero()
            st.error(f"❌ Could not load the file: {err}")
            st.info("💡 For Excel support install the engines: `pip install openpyxl xlrd`.")
            return
        source = uploaded.name
    elif st.session_state.get("use_sample"):
        kind = st.session_state.get("sample_active", "Sales")
        df_raw = make_sample(kind)
        source = f"sample_{kind.lower()}_data.csv"

    if df_raw is None:
        render_hero()
        st.info("👈 Upload an **Excel or CSV** file in the sidebar — or pick a domain and click "
                "**🎲 Load Sample Dataset** — to launch the full platform.")
        c1, c2, c3 = st.columns(3)
        for col, (icon, title, body) in zip((c1, c2, c3), [
            ("📊", "Auto Executive Dashboard", "9 business KPIs, executive summary, auto charts "
             "and a data-quality score — zero configuration."),
            ("🤖", "AI Insights & Root Cause", "What happened, risks, opportunities and prioritized "
             "actions — plus one-click driver decomposition."),
            ("📄", "Board-Ready Outputs", "Chat with your data, then export PDF, PowerPoint, Excel, "
             "CSV and a ready-to-send executive email.")]):
            with col:
                st.markdown(f'<div class="glass"><div style="font-size:1.6rem;">{icon}</div>'
                            f'<div style="font-weight:800;margin:6px 0;">{esc(title)}</div>'
                            f'<div style="color:#B9C6DE;font-size:0.88rem;">{esc(body)}</div></div>',
                            unsafe_allow_html=True)
        return

    df = prepare_df(df_raw)
    if df.empty:
        render_hero()
        st.warning("⚠️ The file was read successfully but contains no data rows.")
        return
    prof = profile_dataset(df)
    render_hero(prof, source)
    fdf = apply_global_filters(df, prof)
    if fdf.empty:
        st.warning("⚠️ Global filters removed every row — adjust or reset them in the sidebar.")
        return
    if len(fdf) != len(df):
        st.caption(f"🌍 Global filters active: **{len(fdf):,}** of **{len(df):,}** rows.")
    ana = analyze(fdf, prof)
    ins = build_insights(fdf, prof, ana)
    render_kpis(ana)
    st.write("")

    tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs([
        "🏠 Executive Dashboard", "🤖 AI Insights", "🔎 Root Cause Analysis",
        "💬 Ask Your Data", "📈 Charts Gallery", "📊 Data Explorer",
        "📄 Reports", "📥 Export",
    ])
    with tab1:
        render_dashboard(fdf, prof, ana, ins)
    with tab2:
        render_insights_tab(ins)
    with tab3:
        render_rca_tab(fdf, prof, ana)
    with tab4:
        render_chat_tab(fdf, prof, ana, ins, settings)
    with tab5:
        render_gallery_tab(fdf, prof, ana)
    with tab6:
        render_explorer_tab(fdf, prof)
    with tab7:
        render_reports_tab(fdf, prof, ana, ins, source)
    with tab8:
        render_export_tab(fdf, prof, ana, ins, source)
    st.divider()
    st.caption(f"🚀 {APP_NAME} {APP_VERSION} · Streamlit · Pandas · NumPy · Plotly · OpenPyXL · "
               "python-pptx · Built-in PDF engine")


if __name__ == "__main__":
    main()
