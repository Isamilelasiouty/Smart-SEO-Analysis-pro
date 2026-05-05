# -*- coding: utf-8 -*-
"""
Smart SEO Analysis Tool — Streamlit Edition
Supports: Google Sheets (OAuth / Service Account) + CSV/Excel upload
GitHub Deploy ready
"""

import io
import re
import json
import time
import warnings
import dataclasses
import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
from collections import defaultdict
from plotly.subplots import make_subplots

warnings.filterwarnings("ignore")

# ─────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Smart SEO Analysis",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────
# CUSTOM CSS — Dark Professional Theme
# ─────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@400;600;700;800&family=DM+Sans:ital,wght@0,300;0,400;0,500;1,300&display=swap');

:root {
    --primary: #6b4ddf;
    --primary-light: #8a6be8;
    --accent: #00d4aa;
    --accent2: #f59e0b;
    --danger: #ef4444;
    --bg: #0d0f1a;
    --bg2: #13162a;
    --bg3: #1a1e35;
    --card: #1e2340;
    --border: #2a2f50;
    --text: #e8eaf6;
    --text-dim: #8892b0;
    --success: #10b981;
}

html, body, [class*="css"] {
    font-family: 'DM Sans', sans-serif;
    background-color: var(--bg) !important;
    color: var(--text) !important;
}

/* Hide Streamlit branding */
#MainMenu, footer, header { visibility: hidden; }
.stDeployButton { display: none; }

/* Main background */
.stApp {
    background: linear-gradient(135deg, #0d0f1a 0%, #0f1222 50%, #0d1520 100%) !important;
}

/* Sidebar */
[data-testid="stSidebar"] {
    background: var(--bg2) !important;
    border-right: 1px solid var(--border);
}
[data-testid="stSidebar"] * { color: var(--text) !important; }

/* Metric cards */
[data-testid="metric-container"] {
    background: var(--card) !important;
    border: 1px solid var(--border);
    border-radius: 16px;
    padding: 20px !important;
    transition: transform .2s, box-shadow .2s;
}
[data-testid="metric-container"]:hover {
    transform: translateY(-2px);
    box-shadow: 0 8px 24px rgba(107,77,223,.25);
}
[data-testid="metric-container"] label {
    color: var(--text-dim) !important;
    font-size: 12px !important;
    text-transform: uppercase;
    letter-spacing: 1px;
}
[data-testid="metric-container"] [data-testid="stMetricValue"] {
    color: var(--text) !important;
    font-family: 'Syne', sans-serif !important;
    font-size: 28px !important;
    font-weight: 700 !important;
}

/* Tabs */
[data-testid="stTabs"] button {
    font-family: 'DM Sans', sans-serif !important;
    color: var(--text-dim) !important;
    border-radius: 8px 8px 0 0 !important;
    font-weight: 500;
}
[data-testid="stTabs"] button[aria-selected="true"] {
    color: var(--primary-light) !important;
    border-bottom: 2px solid var(--primary) !important;
    background: rgba(107,77,223,.1) !important;
}

/* Buttons */
.stButton > button {
    background: linear-gradient(135deg, var(--primary), var(--primary-light)) !important;
    color: white !important;
    border: none !important;
    border-radius: 12px !important;
    padding: 12px 28px !important;
    font-family: 'Syne', sans-serif !important;
    font-weight: 700 !important;
    font-size: 15px !important;
    transition: all .2s !important;
    box-shadow: 0 4px 15px rgba(107,77,223,.4) !important;
}
.stButton > button:hover {
    transform: translateY(-2px) !important;
    box-shadow: 0 8px 25px rgba(107,77,223,.5) !important;
}

/* Inputs */
.stTextInput input, .stTextArea textarea, .stSelectbox select {
    background: var(--card) !important;
    border: 1px solid var(--border) !important;
    border-radius: 10px !important;
    color: var(--text) !important;
    font-family: 'DM Sans', sans-serif !important;
}

/* DataFrames */
[data-testid="stDataFrame"] {
    border-radius: 12px !important;
    border: 1px solid var(--border) !important;
    overflow: hidden !important;
}

/* Expander */
[data-testid="stExpander"] {
    background: var(--card) !important;
    border: 1px solid var(--border) !important;
    border-radius: 12px !important;
}

/* Alerts */
.stAlert {
    border-radius: 12px !important;
    border: none !important;
}

/* Progress */
.stProgress > div > div {
    background: linear-gradient(90deg, var(--primary), var(--accent)) !important;
    border-radius: 10px !important;
}

/* File uploader */
[data-testid="stFileUploader"] {
    background: var(--card) !important;
    border: 2px dashed var(--border) !important;
    border-radius: 16px !important;
    padding: 20px !important;
}

/* Radio */
.stRadio label { color: var(--text) !important; }
.stRadio [data-testid="stMarkdownContainer"] { color: var(--text-dim) !important; }

/* Scrollbar */
::-webkit-scrollbar { width: 6px; height: 6px; }
::-webkit-scrollbar-track { background: var(--bg); }
::-webkit-scrollbar-thumb { background: var(--primary); border-radius: 3px; }

/* Section headers */
.section-header {
    font-family: 'Syne', sans-serif;
    font-size: 22px;
    font-weight: 800;
    color: var(--text);
    margin: 24px 0 16px;
    padding-bottom: 10px;
    border-bottom: 2px solid var(--primary);
    display: flex;
    align-items: center;
    gap: 10px;
}

/* Stat card */
.stat-card {
    background: var(--card);
    border: 1px solid var(--border);
    border-radius: 16px;
    padding: 24px;
    text-align: center;
    transition: all .2s;
}
.stat-card:hover { border-color: var(--primary); transform: translateY(-3px); }
.stat-card .value {
    font-family: 'Syne', sans-serif;
    font-size: 32px;
    font-weight: 800;
    color: var(--primary-light);
}
.stat-card .label {
    font-size: 12px;
    color: var(--text-dim);
    text-transform: uppercase;
    letter-spacing: 1px;
    margin-top: 6px;
}

/* Badge */
.badge {
    display: inline-block;
    padding: 3px 10px;
    border-radius: 20px;
    font-size: 11px;
    font-weight: 600;
}
.badge-success { background: rgba(16,185,129,.15); color: #10b981; }
.badge-warning { background: rgba(245,158,11,.15); color: #f59e0b; }
.badge-danger  { background: rgba(239,68,68,.15);  color: #ef4444; }

/* Hero */
.hero-banner {
    background: linear-gradient(135deg, #1a0a3d 0%, #0d1a3a 50%, #0a2a2a 100%);
    border: 1px solid var(--border);
    border-radius: 24px;
    padding: 40px 48px;
    margin-bottom: 32px;
    position: relative;
    overflow: hidden;
}
.hero-banner::before {
    content: '';
    position: absolute;
    top: -50%;
    right: -10%;
    width: 500px;
    height: 500px;
    background: radial-gradient(circle, rgba(107,77,223,.15) 0%, transparent 70%);
    pointer-events: none;
}
.hero-title {
    font-family: 'Syne', sans-serif;
    font-size: 42px;
    font-weight: 800;
    background: linear-gradient(135deg, #fff 30%, var(--primary-light) 100%);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    margin-bottom: 8px;
}
.hero-sub { color: var(--text-dim); font-size: 16px; }

/* Connection card */
.conn-card {
    background: var(--card);
    border: 2px solid transparent;
    border-radius: 20px;
    padding: 28px;
    cursor: pointer;
    transition: all .2s;
    text-align: center;
}
.conn-card:hover, .conn-card.active {
    border-color: var(--primary);
    background: rgba(107,77,223,.08);
    transform: translateY(-4px);
    box-shadow: 0 12px 32px rgba(107,77,223,.2);
}
.conn-icon { font-size: 36px; margin-bottom: 12px; }
.conn-title {
    font-family: 'Syne', sans-serif;
    font-weight: 700;
    font-size: 16px;
    color: var(--text);
    margin-bottom: 6px;
}
.conn-desc { font-size: 13px; color: var(--text-dim); }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# CONFIG DATACLASS
# ─────────────────────────────────────────────
@dataclasses.dataclass
class SEOConfig:
    primary:   str = "#6b4ddf"
    accent:    str = "#00d4aa"
    success:   str = "#10b981"
    warning:   str = "#f59e0b"
    error:     str = "#ef4444"
    dynamic_percentile:          int   = 25
    click_dominance_threshold:   float = 0.70
    impression_share_threshold:  float = 0.20
    min_impressions_for_opportunity:  int = 100
    min_impressions_for_ctr_problem:  int = 500

CFG = SEOConfig()

# ─────────────────────────────────────────────
# HELPER FUNCTIONS
# ─────────────────────────────────────────────
def safe_divide(a, b, multiplier=1, default=0):
    try:
        if b is None or (isinstance(b, float) and np.isnan(b)) or b == 0:
            return default
        return a / b * multiplier
    except Exception:
        return default

def normalize_text(value):
    text = str(value).strip().lower()
    text = text.replace("_", " ").replace("-", " ")
    return re.sub(r"\s+", " ", text)

def score_column_name(col, keywords, weight=1):
    name = normalize_text(col)
    return sum(1 for k in keywords if k in name) * weight

def clean_numeric_series(series):
    return pd.to_numeric(
        series.astype(str)
              .str.replace(",", "", regex=False)
              .str.replace("%", "", regex=False)
              .str.strip()
              .replace({"": np.nan, "nan": np.nan, "None": np.nan}),
        errors="coerce",
    )

def clean_ctr_series(series):
    raw = series.astype(str).str.strip()
    has_percent = raw.str.contains("%", regex=False, na=False)
    numeric = clean_numeric_series(series).values
    numeric = np.where(has_percent, numeric / 100, numeric)
    numeric = np.where((~has_percent) & (numeric > 1), numeric / 100, numeric)
    return pd.Series(numeric, index=series.index, dtype="float64")

def looks_like_url_series(s):
    s = s.astype(str)
    return (
        s.str.contains("http://",  case=False, regex=False, na=False)
        | s.str.contains("https://", case=False, regex=False, na=False)
        | s.str.contains("www.",    case=False, regex=False, na=False)
        | s.str.contains("/",                  regex=False, na=False)
    )

def make_unique_columns(columns):
    seen = {}
    result = []
    for col in columns:
        base = str(col).strip() or "Unnamed"
        count = seen.get(base, 0) + 1
        seen[base] = count
        result.append(base if count == 1 else f"{base} __{count}")
    return result

def expected_ctr_by_position(position):
    if pd.isna(position): return 0.02
    p = float(position)
    if p <= 1:  return 0.22
    if p <= 2:  return 0.13
    if p <= 3:  return 0.09
    if p <= 5:  return 0.05
    if p <= 10: return 0.03
    if p <= 20: return 0.012
    if p <= 50: return 0.004
    return 0.001

def dynamic_report_limit(df, cap=10000):
    rows = len(df) if df is not None else 0
    if rows < 1000:  return rows
    if rows < 10000: return min(rows, cap)
    if rows < 100000: return max(500, min(int(rows * 0.20), cap))
    return max(1000, min(int(rows * 0.10), cap))

# ─────────────────────────────────────────────
# COLUMN DETECTION
# ─────────────────────────────────────────────
def standardize_columns(df):
    df = df.copy()
    df = df.dropna(how="all").dropna(how="all", axis=1)
    raw_columns = list(df.columns)

    def get_text(col):
        return df[col].dropna().astype(str).str.strip().head(500)

    def get_num(col):
        s = get_text(col)
        s = s.str.replace(",", "", regex=False).str.replace("%", "", regex=False).str.strip()
        return pd.to_numeric(s, errors="coerce").dropna()

    mapping = {}

    # Query
    best_query, best_score = None, -999
    for col in raw_columns:
        s = get_text(col)
        if len(s) == 0: continue
        numeric_ratio = pd.to_numeric(s.str.replace(",", "", regex=False), errors="coerce").notna().mean()
        url_ratio = looks_like_url_series(s).mean()
        avg_len = s.str.len().mean()
        text_ratio = s.str.contains(r"[A-Za-zء-ي]", regex=True, na=False).mean()
        score = (score_column_name(col, ["query","keyword","search query","search term","term","استعلام","كلمة","بحث"], 6)
                 + text_ratio * 4 + (3 if 1 <= avg_len <= 100 else 0) - numeric_ratio * 8 - url_ratio * 5)
        if score > best_score:
            best_score, best_query = score, col
    if best_query and best_score >= 4:
        mapping[best_query] = "Query string"

    # Landing page
    best_page, best_score = None, -999
    for col in raw_columns:
        if col == best_query: continue
        s = get_text(col)
        if len(s) == 0: continue
        url_ratio = looks_like_url_series(s).mean()
        score = score_column_name(col, ["landing","page","url","link","address","الصفحة","الرابط"], 5) + url_ratio * 10
        if score > best_score:
            best_score, best_page = score, col
    if best_page and best_score >= 4:
        mapping[best_page] = "Landing page"

    # CTR
    best_ctr, best_score = None, -999
    for col in raw_columns:
        if col in mapping: continue
        s = get_text(col)
        nums = get_num(col)
        if len(nums) == 0: continue
        percent_ratio = s.str.contains("%", regex=False, na=False).mean()
        small_ratio = ((nums >= 0) & (nums <= 1)).mean()
        score = (score_column_name(col, ["ctr","click through rate","rate","نسبة النقر","معدل النقر"], 8)
                 + percent_ratio * 8 + small_ratio * 4)
        if score > best_score:
            best_score, best_ctr = score, col
    if best_ctr and best_score >= 4:
        mapping[best_ctr] = "CTR"

    # Position
    best_position, best_score = None, -999
    for col in raw_columns:
        if col in mapping: continue
        nums = get_num(col)
        if len(nums) == 0: continue
        range_ratio = ((nums >= 1) & (nums <= 150)).mean()
        decimal_ratio = ((nums % 1) != 0).mean()
        score = (score_column_name(col, ["position","avg position","average position","rank","ranking","ترتيب","الموضع"], 8)
                 + range_ratio * 6 + decimal_ratio * 2)
        if nums.quantile(0.95) > 300: score -= 8
        if score > best_score:
            best_score, best_position = score, col
    if best_position and best_score >= 4:
        mapping[best_position] = "Position"

    # Clicks & Impressions
    numeric_candidates = []
    for col in raw_columns:
        if col in mapping: continue
        nums = get_num(col)
        if len(nums) == 0: continue
        name = normalize_text(col)
        numeric_candidates.append({
            "col": col, "sum": nums.sum(),
            "score": ((nums % 1) == 0).mean() * 4 + (nums >= 0).mean() * 3,
            "is_clicks": any(k in name for k in ["click","clicks","نقر","نقرات"]),
            "is_impressions": any(k in name for k in ["impression","impressions","ظهور"]),
        })

    clicks_col = impressions_col = None
    clicks_named = [x for x in numeric_candidates if x["is_clicks"]]
    impressions_named = [x for x in numeric_candidates if x["is_impressions"]]
    if clicks_named: clicks_col = sorted(clicks_named, key=lambda x: x["score"], reverse=True)[0]["col"]
    if impressions_named: impressions_col = sorted(impressions_named, key=lambda x: x["score"], reverse=True)[0]["col"]

    if clicks_col is None or impressions_col is None or clicks_col == impressions_col:
        by_sum = sorted(numeric_candidates, key=lambda x: x["sum"], reverse=True)
        if impressions_col is None and len(by_sum) >= 1: impressions_col = by_sum[0]["col"]
        if clicks_col is None and len(by_sum) >= 2: clicks_col = by_sum[1]["col"]

    if clicks_col and clicks_col not in mapping: mapping[clicks_col] = "Clicks"
    if impressions_col and impressions_col not in mapping: mapping[impressions_col] = "Impressions"

    clean_mapping = {}
    used_targets = set()
    for source, target in mapping.items():
        if target not in used_targets:
            clean_mapping[source] = target
            used_targets.add(target)

    return df.rename(columns=clean_mapping)


def clean_data(df):
    df = standardize_columns(df)
    df["Clicks"] = clean_numeric_series(df["Clicks"]).fillna(0) if "Clicks" in df.columns else pd.Series(0, index=df.index)
    df["Impressions"] = clean_numeric_series(df["Impressions"]).fillna(0) if "Impressions" in df.columns else pd.Series(0, index=df.index)
    df["Position"] = clean_numeric_series(df["Position"]) if "Position" in df.columns else pd.Series(np.nan, index=df.index)
    if "CTR" in df.columns:
        df["CTR"] = clean_ctr_series(df["CTR"])
    else:
        df["CTR"] = np.where(df["Impressions"] > 0, df["Clicks"] / df["Impressions"], 0)
    if "Query string" in df.columns: df["Query string"] = df["Query string"].astype(str).str.strip()
    if "Landing page" in df.columns: df["Landing page"] = df["Landing page"].astype(str).str.strip()
    return df

# ─────────────────────────────────────────────
# COMPARISON TAB DETECTION
# ─────────────────────────────────────────────
def normalize_comparison_columns(df):
    df = df.copy()
    original_cols = list(df.columns)

    def col_text(col): return normalize_text(str(col).replace("__", " "))

    previous_words = ["previous","prev","past","last","old","before","former","comparison","السابق","سابقة","قبل","المقارنة"]
    current_words  = ["current","curr","now","latest","الحالي","حالية"]
    metric_words = {
        "Clicks":      ["click","clicks","نقر","نقرات"],
        "Impressions": ["impression","impressions","ظهور"],
        "CTR":         ["ctr","click through rate","نسبة النقر","معدل النقر"],
        "Position":    ["position","avg position","average position","rank","ranking","ترتيب","الموضع"],
    }

    def get_text(col):
        try:
            v = df[col]
            if isinstance(v, pd.DataFrame): v = v.iloc[:,0]
            return v.dropna().astype(str).str.strip().head(500)
        except: return pd.Series(dtype=str)

    def get_num(col):
        s = get_text(col)
        s = s.str.replace(",","",regex=False).str.replace("%","",regex=False).str.strip()
        return pd.to_numeric(s, errors="coerce").dropna()

    def identity_score(col, kind):
        name = col_text(col)
        s = get_text(col)
        if len(s) == 0: return -999
        numeric_ratio = pd.to_numeric(s.str.replace(",","",regex=False), errors="coerce").notna().mean()
        url_ratio = looks_like_url_series(s).mean()
        text_ratio = s.str.contains(r"[A-Za-zء-ي]", regex=True, na=False).mean()
        avg_len = s.str.len().mean()
        if kind == "query":
            ns = score_column_name(col, ["query","keyword","search query","search term","term","استعلام","كلمة","بحث"], 8)
            return ns + text_ratio*4 + (3 if 1<=avg_len<=120 else 0) - numeric_ratio*8 - url_ratio*5
        if kind == "page":
            ns = score_column_name(col, ["landing","page","url","link","address","الصفحة","الرابط"], 8)
            return ns + url_ratio*12
        return -999

    rename_map = {}
    query_scores = sorted([(identity_score(c,"query"),c) for c in original_cols], reverse=True)
    page_scores  = sorted([(identity_score(c,"page"),c)  for c in original_cols], reverse=True)
    query_col = query_scores[0][1] if query_scores and query_scores[0][0]>=4 else None
    page_col  = page_scores[0][1]  if page_scores  and page_scores[0][0]>=4  else None
    if query_col: rename_map[query_col] = "Query string"
    if page_col and page_col != query_col: rename_map[page_col] = "Landing page"
    identity_cols = {query_col, page_col}

    def period_type(col):
        name = col_text(col)
        prev_score = sum(1 for w in previous_words if w in name)
        curr_score = sum(1 for w in current_words  if w in name)
        if re.search(r"\b__?\s*2\b", str(col).lower()) or str(col).strip().endswith("__2"): prev_score += 2
        if prev_score > curr_score: return "previous"
        if curr_score > prev_score: return "current"
        return "unknown"

    def metric_score_fn(col, metric):
        if col in identity_cols: return -999
        name = col_text(col)
        nums = get_num(col)
        if len(nums) == 0: return -999
        score = sum(1 for w in metric_words[metric] if w in name) * 10
        if metric == "CTR":
            raw = get_text(col)
            pct = raw.str.contains("%", regex=False, na=False).mean() if len(raw) else 0
            sr  = ((nums>=0)&(nums<=1)).mean()
            score += pct*8 + sr*4
        elif metric == "Position":
            rr = ((nums>=1)&(nums<=200)).mean()
            dr = ((nums%1)!=0).mean()
            score += rr*5 + dr*2
            if nums.quantile(0.95)>500: score -= 8
        else:
            score += ((nums%1)==0).mean()*3 + (nums>=0).mean()*3
        return score

    for metric in ["Clicks","Impressions","CTR","Position"]:
        candidates = [{"col":col,"score":metric_score_fn(col,metric),"period":period_type(col),"order":original_cols.index(col)}
                      for col in original_cols if col not in identity_cols and metric_score_fn(col,metric)>=8]
        if not candidates: continue
        candidates = sorted(candidates, key=lambda x:(x["score"],-x["order"]), reverse=True)
        prev_c = [x for x in candidates if x["period"]=="previous"]
        curr_c = [x for x in candidates if x["period"]=="current"]
        current_col  = curr_c[0]["col"]  if curr_c  else None
        previous_col = prev_c[0]["col"]  if prev_c  else None
        if previous_col and not current_col:
            others = [x for x in candidates if x["col"]!=previous_col]
            if others: current_col = sorted(others, key=lambda x:x["score"], reverse=True)[0]["col"]
        if current_col and not previous_col:
            others = [x for x in candidates if x["col"]!=current_col]
            if others: previous_col = sorted(others, key=lambda x:x["score"], reverse=True)[0]["col"]
        if not current_col and not previous_col and len(candidates)>=2:
            by_order = sorted(candidates, key=lambda x:x["order"])
            current_col  = by_order[0]["col"]
            previous_col = by_order[1]["col"]
        if not current_col and len(candidates)==1: current_col = candidates[0]["col"]
        if current_col and current_col not in rename_map: rename_map[current_col] = f"{metric} Current"
        if previous_col and previous_col!=current_col and previous_col not in rename_map: rename_map[previous_col] = f"{metric} Previous"

    return df.rename(columns=rename_map)


def clean_comparison_data(df):
    df = normalize_comparison_columns(df)
    if "Query string" in df.columns: df["Query string"] = df["Query string"].astype(str).str.strip()
    if "Landing page" in df.columns: df["Landing page"] = df["Landing page"].astype(str).str.strip()
    for col in ["Clicks Current","Clicks Previous","Impressions Current","Impressions Previous","Position Current","Position Previous"]:
        if col in df.columns: df[col] = clean_numeric_series(df[col]).fillna(0)
    for col in ["CTR Current","CTR Previous"]:
        if col in df.columns: df[col] = clean_ctr_series(df[col]).fillna(0)
    if "CTR Current" not in df.columns and {"Clicks Current","Impressions Current"}.issubset(df.columns):
        df["CTR Current"] = np.where(df["Impressions Current"]>0, df["Clicks Current"]/df["Impressions Current"], 0)
    if "CTR Previous" not in df.columns and {"Clicks Previous","Impressions Previous"}.issubset(df.columns):
        df["CTR Previous"] = np.where(df["Impressions Previous"]>0, df["Clicks Previous"]/df["Impressions Previous"], 0)
    return df

def available_comparison_pairs(df):
    return [m for m in ["Clicks","Impressions","CTR","Position"]
            if f"{m} Current" in df.columns and f"{m} Previous" in df.columns]

def comparison_score(df, sheet_name=""):
    name = normalize_text(sheet_name)
    cols = set(df.columns)
    pairs = available_comparison_pairs(df)
    if not pairs or ("Query string" not in cols and "Landing page" not in cols): return 0
    score = len(pairs)*35
    if "Query string" in cols: score += 20
    if "Landing page" in cols: score += 20
    if any(w in name for w in ["comparison","compare","period","previous","prev","decline","drop","مقارنة","السابق","هبوط"]): score += 30
    score += min(len(df),1000)/1000
    return score

def source_score(df, sheet_name=""):
    name = normalize_text(sheet_name)
    score = (("Query string" in df.columns)*50 + ("Landing page" in df.columns)*40
             + ("Clicks" in df.columns)*10 + ("Impressions" in df.columns)*10
             + ("Position" in df.columns)*10)
    if "query" in name or "queries" in name or "keyword" in name: score += 25
    if "landing" in name or "page" in name: score += 15
    score += min(len(df),1000)/1000
    return score

# ─────────────────────────────────────────────
# KEYWORD ANALYSIS FUNCTIONS
# ─────────────────────────────────────────────
def normalize_keyword_text(value):
    text = str(value).strip().lower()
    for src, dst in {"أ":"ا","إ":"ا","آ":"ا","ى":"ي","ة":"ه","ؤ":"و","ئ":"ي"}.items():
        text = text.replace(src, dst)
    text = re.sub(r"[^a-z0-9ء-ي ]", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    stop_words = {"the","a","an","and","or","for","to","of","in","on","with","store","shop","website","official","egypt","eg","متجر","موقع","رسمي","مصر","في","من","على","و","ال"}
    tokens = [t for t in text.split() if t not in stop_words]
    return " ".join(tokens) if tokens else text

def infer_brand_terms(df, top_n=8):
    if "Query string" not in df.columns: return []
    top_queries = (df.groupby("Query string", as_index=False)["Clicks"].sum()
                   .sort_values("Clicks", ascending=False).head(50)["Query string"].astype(str).tolist())
    counts = defaultdict(int)
    for q in top_queries:
        norm = normalize_keyword_text(q)
        for token in set(norm.split()):
            if len(token) >= 3: counts[token] += 1
    return [k for k, v in counts.items() if v >= 3][:top_n]

def keyword_cluster_key(keyword, brand_terms=None):
    normalized = normalize_keyword_text(keyword)
    if brand_terms:
        for brand in brand_terms:
            brand_norm = normalize_keyword_text(brand)
            if brand_norm and brand_norm in normalized: return f"brand::{brand_norm}"
    tokens = normalized.split()
    if not tokens: return normalized
    if len(tokens) <= 3: return " ".join(sorted(set(tokens)))
    return " ".join(sorted(set(tokens[:4])))

def detect_query_intent(query):
    q = normalize_keyword_text(query)
    patterns = {
        "Brand":         r"yamama|elyamama|اليمامه|يمامه",
        "Price":         r"price|cost|سعر|اسعار|تكلفة",
        "Comparison":    r"vs|versus|compare|comparison|مقارنة|افضل|أفضل",
        "Commercial":    r"best|review|buy|deal|discount|شراء|مراجعة|عرض|خصم",
        "Transactional": r"order|shop|subscribe|register|download|اطلب|اشتري|تحميل|اشترك",
        "Question":      r"how|what|why|when|where|can|should|طريقة|ما هو|ما هي|كيف|هل|شرح",
        "Navigational":  r"login|official|website|contact|app|دخول|رسمي|تواصل|تطبيق",
    }
    for intent, pattern in patterns.items():
        if re.search(pattern, q, flags=re.IGNORECASE): return intent
    return "Informational"

def detect_page_type(url):
    u = str(url).lower()
    if u in ("","nan","none"): return "Unknown"
    if "/blog" in u or "/blogs" in u: return "Blog"
    if "/collection" in u or "/collections" in u or "/category" in u: return "Collection"
    if "/product" in u or "/products" in u: return "Product"
    if "compare" in u or "vs" in u: return "Comparison"
    if u == "/" or u.endswith(".com") or u.endswith(".com/"): return "Homepage"
    return "Other"

def add_enrichment_columns(df):
    df = df.copy()
    if "Query string" in df.columns:
        brand_terms = infer_brand_terms(df)
        df["Keyword Cluster"] = df["Query string"].apply(lambda x: keyword_cluster_key(x, brand_terms))
        df["Intent"] = df["Query string"].apply(detect_query_intent)
        df["Is Brand"] = df["Keyword Cluster"].astype(str).str.startswith("brand::")
    if "Landing page" in df.columns:
        df["Page Type"] = df["Landing page"].apply(detect_page_type)
    return df

# ─────────────────────────────────────────────
# ANALYSIS FUNCTIONS
# ─────────────────────────────────────────────
def create_overview(df):
    total_clicks = df["Clicks"].sum()
    total_impressions = df["Impressions"].sum()
    return {
        "Total Rows": f"{len(df):,}",
        "Unique Queries": f"{df['Query string'].nunique():,}" if "Query string" in df.columns else "N/A",
        "Unique Pages": f"{df['Landing page'].nunique():,}" if "Landing page" in df.columns else "N/A",
        "Total Clicks": f"{total_clicks:,.0f}",
        "Total Impressions": f"{total_impressions:,.0f}",
        "Avg CTR (%)": f"{safe_divide(total_clicks, total_impressions, 100):.2f}%",
        "Avg Position": f"{df['Position'].mean():.2f}" if "Position" in df.columns else "N/A",
    }

def create_top_pages(df):
    if "Landing page" not in df.columns: return pd.DataFrame({"Note":["No landing page column found."]})
    result = df.groupby("Landing page", as_index=False).agg(Clicks=("Clicks","sum"), Impressions=("Impressions","sum"), Position=("Position","mean"))
    result["CTR (%)"] = np.where(result["Impressions"]>0, result["Clicks"]/result["Impressions"]*100, 0).round(2)
    result["Position"] = result["Position"].round(2)
    result["Page Type"] = result["Landing page"].apply(detect_page_type)
    limit = dynamic_report_limit(result)
    return result.sort_values("Clicks", ascending=False).head(limit).rename(columns={"Landing page":"Page","Position":"Average Position"})

def create_smart_top_keywords(df):
    if "Query string" not in df.columns: return pd.DataFrame({"Note":["No query column found."]})
    enriched = add_enrichment_columns(df)
    result = enriched.groupby("Query string", as_index=False).agg(
        Clicks=("Clicks","sum"), Impressions=("Impressions","sum"), Position=("Position","mean"),
        Keyword_Cluster=("Keyword Cluster","first"), Intent=("Intent","first"), Is_Brand=("Is Brand","first")
    ).rename(columns={"Keyword_Cluster":"Keyword Cluster","Is_Brand":"Is Brand"})
    result["CTR (%)"] = np.where(result["Impressions"]>0, result["Clicks"]/result["Impressions"]*100, 0).round(2)
    result["Position"] = result["Position"].round(2)
    cluster_totals = result.groupby("Keyword Cluster", as_index=False).agg(
        Clicks=("Clicks","sum"), Impressions=("Impressions","sum"),
        **{"Query string":("Query string","nunique")}
    ).rename(columns={"Query string":"Similar Variants"})
    reps = result.sort_values(["Clicks","Impressions"], ascending=False).drop_duplicates("Keyword Cluster")
    final = reps.merge(cluster_totals, on="Keyword Cluster", suffixes=("","_Cluster"))
    limit = dynamic_report_limit(final)
    final = final.sort_values(["Clicks_Cluster","Impressions_Cluster"], ascending=False).head(limit)
    return final[["Query string","Keyword Cluster","Similar Variants","Intent","Is Brand",
                  "Clicks","Impressions","Clicks_Cluster","Impressions_Cluster","CTR (%)","Position"]].rename(
        columns={"Query string":"Representative Keyword","Clicks":"Rep Clicks","Impressions":"Rep Impressions",
                 "Clicks_Cluster":"Cluster Clicks","Impressions_Cluster":"Cluster Impressions","Position":"Average Position"})

def create_opportunity_score_report(df):
    if "Query string" not in df.columns: return pd.DataFrame({"Note":["No query column found."]})
    enriched = add_enrichment_columns(df)
    temp = enriched.groupby("Query string", as_index=False).agg(
        Clicks=("Clicks","sum"), Impressions=("Impressions","sum"), CTR=("CTR","mean"),
        Position=("Position","mean"), Intent=("Intent","first"), Is_Brand=("Is Brand","first"),
        Keyword_Cluster=("Keyword Cluster","first")
    ).rename(columns={"Is_Brand":"Is Brand","Keyword_Cluster":"Keyword Cluster"})
    temp = temp[temp["Impressions"]>=CFG.min_impressions_for_opportunity].copy()
    if temp.empty: return pd.DataFrame({"Note":["No opportunity candidates found."]})
    temp["Expected CTR"] = temp["Position"].apply(expected_ctr_by_position)
    temp["CTR Gap"] = (temp["Expected CTR"] - temp["CTR"]).clip(lower=0)
    temp["Potential Click Gain"] = (temp["CTR Gap"] * temp["Impressions"]).round(0)
    temp["Position Score"] = np.select([temp["Position"].between(4,10), temp["Position"].between(11,20), temp["Position"].between(21,50)], [35,28,16], default=5)
    temp["Intent Score"] = temp["Intent"].map({"Commercial":20,"Transactional":20,"Price":18,"Comparison":18,"Question":12,"Informational":10,"Navigational":5,"Brand":3}).fillna(8)
    temp["Brand Penalty"] = np.where(temp["Is Brand"],-15,0)
    temp["Volume Score"] = np.log10(temp["Impressions"].clip(lower=1))*8
    temp["CTR Gap Score"] = temp["CTR Gap"]*100
    temp["Opportunity Score"] = (temp["Position Score"]+temp["Intent Score"]+temp["Volume Score"]+temp["CTR Gap Score"]+temp["Brand Penalty"]).clip(lower=0,upper=100).round(2)
    limit = dynamic_report_limit(temp)
    return temp.sort_values(["Opportunity Score","Potential Click Gain"], ascending=False).head(limit)[[
        "Query string","Keyword Cluster","Intent","Is Brand","Clicks","Impressions","CTR","Expected CTR","CTR Gap","Position","Potential Click Gain","Opportunity Score"
    ]].rename(columns={"Query string":"Keyword","Position":"Average Position","CTR":"Current CTR"})

def create_ctr_problems(df):
    group_col = "Landing page" if "Landing page" in df.columns else "Query string" if "Query string" in df.columns else None
    if not group_col: return pd.DataFrame({"Note":["Not enough data."]})
    result = df.groupby(group_col, as_index=False).agg(Clicks=("Clicks","sum"), Impressions=("Impressions","sum"), Position=("Position","mean"))
    result["CTR"] = np.where(result["Impressions"]>0, result["Clicks"]/result["Impressions"], 0)
    result["CTR (%)"] = (result["CTR"]*100).round(2)
    result["Expected CTR"] = result["Position"].apply(expected_ctr_by_position)
    result["Lost Clicks"] = ((result["Expected CTR"]-result["CTR"]).clip(lower=0)*result["Impressions"]).round(0)
    result = result[(result["Impressions"]>=CFG.min_impressions_for_ctr_problem)&(result["Lost Clicks"]>0)].copy()
    if result.empty: return pd.DataFrame({"Note":["No major CTR problems found."]})
    limit = dynamic_report_limit(result)
    return result.sort_values("Lost Clicks", ascending=False).head(limit)

def create_cannibalization(df):
    required = {"Query string","Landing page","Clicks","Impressions"}
    if df is None or not required.issubset(set(df.columns)): return pd.DataFrame({"Note":["Cannibalization requires Query x Landing Page data."]})
    temp = df.dropna(subset=["Query string","Landing page"]).copy()
    temp = temp[(temp["Query string"].astype(str).str.strip()!="")&(temp["Landing page"].astype(str).str.strip()!="")]
    if temp.empty: return pd.DataFrame({"Note":["No valid data."]})
    temp["Page Type"] = temp["Landing page"].apply(detect_page_type)
    temp["Intent"] = temp["Query string"].apply(detect_query_intent)
    temp = temp.groupby(["Query string","Landing page"], as_index=False).agg(
        Clicks=("Clicks","sum"), Impressions=("Impressions","sum"), Position=("Position","mean"),
        Page_Type=("Page Type","first"), Intent=("Intent","first")
    ).rename(columns={"Page_Type":"Page Type"})
    multi = temp.groupby("Query string")["Landing page"].nunique()
    multi_queries = multi[multi>1].index
    df_multi = temp[temp["Query string"].isin(multi_queries)].copy()
    if df_multi.empty: return pd.DataFrame({"Note":["No keyword has multiple competing pages."]})
    threshold = np.percentile(df_multi.groupby("Query string")["Impressions"].sum(), CFG.dynamic_percentile)
    rows = []
    for query, group in df_multi.groupby("Query string"):
        total_clicks = group["Clicks"].sum()
        total_impressions = group["Impressions"].sum()
        if total_impressions<threshold or total_impressions<=0: continue
        group = group.copy()
        group["click_share"] = np.where(total_clicks>0, group["Clicks"]/total_clicks, 0)
        group["impression_share"] = group["Impressions"]/total_impressions
        pages_with_share = int((group["impression_share"]>=CFG.impression_share_threshold).sum())
        max_click_share = group["click_share"].max()
        if not (max_click_share<=CFG.click_dominance_threshold and pages_with_share>=2): continue
        pages_count = group["Landing page"].nunique()
        severity = (1-max_click_share)*100*(1+(pages_count-2)*0.1)
        position_spread = group["Position"].max()-group["Position"].min()
        if position_spread<=5: severity+=10
        if group["Page Type"].nunique()>1: severity+=5
        severity = min(round(severity,2),100)
        main_page = group.sort_values(["Clicks","Impressions"], ascending=False).iloc[0]["Landing page"]
        rows.append({"Keyword":query,"Intent":group["Intent"].iloc[0],"Competing Pages":int(pages_count),
                     "Total Clicks":int(total_clicks),"Total Impressions":int(total_impressions),
                     "Severity Score":severity,"Max Click Share (%)":round(max_click_share*100,2),
                     "Current Strongest Page":main_page,"Competing URLs":" | ".join(group.sort_values("Impressions",ascending=False)["Landing page"].tolist()[:8]),
                     "Recommended Action":"Choose a primary page, strengthen internal links, and use canonical/merge/redirect where needed."})
    if not rows: return pd.DataFrame({"Note":["No significant cannibalization found."]})
    result_df = pd.DataFrame(rows)
    limit = dynamic_report_limit(result_df)
    return result_df.sort_values(["Severity Score","Total Impressions"], ascending=False).head(limit)

def create_position_distribution(df):
    bins = [0,1,3,10,20,50,100,float("inf")]
    labels = ["#1","#2-3","#4-10","#11-20","#21-50","#51-100","+100"]
    temp = df.copy()
    temp["Position Range"] = pd.cut(temp["Position"], bins=bins, labels=labels, include_lowest=True).astype(str)
    result = temp.groupby("Position Range", observed=True).agg(Clicks=("Clicks","sum"), Impressions=("Impressions","sum")).reset_index()
    result["CTR (%)"] = np.where(result["Impressions"]>0, result["Clicks"]/result["Impressions"]*100, 0).round(2)
    return result

def create_brand_nonbrand(df):
    if "Query string" not in df.columns: return pd.DataFrame({"Note":["No query column."]})
    enriched = add_enrichment_columns(df)
    result = enriched.groupby("Is Brand", as_index=False).agg(
        Clicks=("Clicks","sum"), Impressions=("Impressions","sum"),
        Unique_Keywords=("Query string","nunique"), Position=("Position","mean")
    ).rename(columns={"Unique_Keywords":"Unique Keywords"})
    result["CTR (%)"] = np.where(result["Impressions"]>0, result["Clicks"]/result["Impressions"]*100, 0).round(2)
    result["Position"] = result["Position"].round(2)
    result["Segment"] = np.where(result["Is Brand"],"Brand","Non-brand")
    return result[["Segment","Unique Keywords","Clicks","Impressions","CTR (%)","Position"]]

def create_intent_analysis(df):
    if "Query string" not in df.columns: return pd.DataFrame({"Note":["No query column."]})
    enriched = add_enrichment_columns(df)
    result = enriched.groupby("Intent", as_index=False).agg(
        Clicks=("Clicks","sum"), Impressions=("Impressions","sum"),
        Unique_Keywords=("Query string","nunique"), Position=("Position","mean")
    ).rename(columns={"Unique_Keywords":"Unique Keywords"})
    result["CTR (%)"] = np.where(result["Impressions"]>0, result["Clicks"]/result["Impressions"]*100, 0).round(2)
    result["Position"] = result["Position"].round(2)
    return result.sort_values("Clicks", ascending=False)

# ─────────────────────────────────────────────
# DECLINE ANALYSIS
# ─────────────────────────────────────────────
def has_current_previous_pair(df, metric):
    return f"{metric} Current" in df.columns and f"{metric} Previous" in df.columns

def add_drop_metrics(df):
    df = df.copy()
    pairs = available_comparison_pairs(df)
    if not pairs:
        df["Drop Priority Score"] = 0
        return df
    if has_current_previous_pair(df,"Clicks"):
        df["Clicks Change"] = df["Clicks Current"]-df["Clicks Previous"]
        df["Clicks Drop"] = (df["Clicks Previous"]-df["Clicks Current"]).clip(lower=0)
        df["Clicks Drop %"] = np.where(df["Clicks Previous"]>0, df["Clicks Drop"]/df["Clicks Previous"]*100, 0).round(2)
    else:
        df["Clicks Drop"] = df["Clicks Drop %"] = 0
    if has_current_previous_pair(df,"Impressions"):
        df["Impressions Change"] = df["Impressions Current"]-df["Impressions Previous"]
        df["Impressions Drop"] = (df["Impressions Previous"]-df["Impressions Current"]).clip(lower=0)
        df["Impressions Drop %"] = np.where(df["Impressions Previous"]>0, df["Impressions Drop"]/df["Impressions Previous"]*100, 0).round(2)
    else:
        df["Impressions Drop"] = df["Impressions Drop %"] = 0
    if has_current_previous_pair(df,"CTR"):
        df["CTR Drop"] = (df["CTR Previous"]-df["CTR Current"]).clip(lower=0)
        df["CTR Drop %"] = np.where(df["CTR Previous"]>0, df["CTR Drop"]/df["CTR Previous"]*100, 0).round(2)
    else:
        df["CTR Drop"] = df["CTR Drop %"] = 0
    if has_current_previous_pair(df,"Position"):
        df["Position Drop"] = (df["Position Current"]-df["Position Previous"]).clip(lower=0)
    else:
        df["Position Drop"] = 0
    score = pd.Series(0.0, index=df.index)
    if has_current_previous_pair(df,"Clicks"):
        score += np.log10(df["Clicks Previous"].clip(lower=1))*12 + df["Clicks Drop %"].clip(upper=100)*0.45 + df["Clicks Drop"].clip(upper=500)*0.03
    if has_current_previous_pair(df,"Impressions"):
        score += np.log10(df["Impressions Previous"].clip(lower=1))*8 + df["Impressions Drop %"].clip(upper=100)*0.20
    if has_current_previous_pair(df,"CTR"):
        score += df["CTR Drop %"].clip(upper=100)*0.15
    if has_current_previous_pair(df,"Position"):
        score += df["Position Drop"].clip(upper=20)*2.5
    df["Drop Priority Score"] = score.clip(lower=0,upper=100).round(2)
    return df

def has_any_drop(row):
    return (row.get("Clicks Drop",0)>0 or row.get("Impressions Drop",0)>0
            or row.get("CTR Drop",0)>0 or row.get("Position Drop",0)>0)

def infer_drop_reason(row):
    if row.get("Position Drop",0)>=2: return "Ranking drop"
    if row.get("CTR Drop",0)>0 and row.get("Clicks Drop",0)>0 and not row.get("Impressions Drop",0): return "CTR drop"
    if row.get("Impressions Drop",0)>0 and row.get("Clicks Drop",0)>0: return "Demand or visibility drop"
    if row.get("Impressions Drop",0)>0: return "Visibility drop"
    if row.get("Clicks Drop",0)>0: return "Clicks dropped"
    if row.get("CTR Drop",0)>0: return "CTR drop"
    return "No major drop"

def infer_drop_action(row):
    actions = {
        "Ranking drop": "Refresh content, strengthen internal links, check SERP competitors, and review technical/indexing issues.",
        "CTR drop": "Rewrite title/meta description, improve SERP value proposition, and align snippet with intent.",
        "Demand or visibility drop": "Check seasonality, indexing, SERP changes, lost impressions by query/page, and content freshness.",
        "Visibility drop": "Check seasonality, indexing, SERP changes, and content freshness.",
        "Clicks dropped": "Review query/page performance details and compare CTR, ranking, and impressions.",
    }
    return actions.get(row.get("Likely Reason",""), "Monitor only.")

def create_declining_keywords(comparison_df):
    if comparison_df is None: return pd.DataFrame({"Note":["No comparison tab found."]})
    if "Query string" not in comparison_df.columns: return pd.DataFrame({"Note":["No query column in comparison tab."]})
    pairs = available_comparison_pairs(comparison_df)
    if not pairs: return pd.DataFrame({"Note":["No current/previous metric pairs detected."]})
    df = add_drop_metrics(comparison_df)
    df = df[df.apply(has_any_drop, axis=1)].copy()
    if df.empty: return pd.DataFrame({"Note":["No keyword decline detected."]})
    df["Likely Reason"] = df.apply(infer_drop_reason, axis=1)
    df["Recommended Action"] = df.apply(infer_drop_action, axis=1)
    columns = ["Query string"] + [c for c in ["Clicks Current","Clicks Previous","Clicks Drop","Clicks Drop %",
        "Impressions Current","Impressions Previous","Impressions Drop","Impressions Drop %",
        "CTR Current","CTR Previous","CTR Drop","CTR Drop %","Position Current","Position Previous","Position Drop",
        "Drop Priority Score","Likely Reason","Recommended Action"] if c in df.columns]
    limit = dynamic_report_limit(df)
    return df.sort_values(["Drop Priority Score","Clicks Drop","Impressions Drop"], ascending=False).head(limit)[columns].rename(columns={"Query string":"Keyword"})

def create_declining_pages(comparison_df):
    if comparison_df is None: return pd.DataFrame({"Note":["No comparison tab found."]})
    if "Landing page" not in comparison_df.columns: return pd.DataFrame({"Note":["No page column in comparison tab."]})
    pairs = available_comparison_pairs(comparison_df)
    if not pairs: return pd.DataFrame({"Note":["No current/previous metric pairs detected."]})
    df = comparison_df.copy()
    agg = {}
    for metric in ["Clicks","Impressions"]:
        if has_current_previous_pair(df,metric):
            agg[f"{metric} Current"] = "sum"; agg[f"{metric} Previous"] = "sum"
    for metric in ["CTR","Position"]:
        if has_current_previous_pair(df,metric):
            agg[f"{metric} Current"] = "mean"; agg[f"{metric} Previous"] = "mean"
    if not agg: return pd.DataFrame({"Note":["No usable comparison metrics."]})
    page_df = df.groupby("Landing page", as_index=False).agg(agg)
    page_df = add_drop_metrics(page_df)
    page_df = page_df[page_df.apply(has_any_drop, axis=1)].copy()
    if page_df.empty: return pd.DataFrame({"Note":["No page decline detected."]})
    page_df["Page Type"] = page_df["Landing page"].apply(detect_page_type)
    page_df["Likely Reason"] = page_df.apply(infer_drop_reason, axis=1)
    page_df["Recommended Action"] = page_df.apply(infer_drop_action, axis=1)
    columns = ["Landing page","Page Type"] + [c for c in ["Clicks Current","Clicks Previous","Clicks Drop","Clicks Drop %",
        "Impressions Current","Impressions Previous","Impressions Drop","Impressions Drop %",
        "CTR Current","CTR Previous","CTR Drop","CTR Drop %","Position Current","Position Previous","Position Drop",
        "Drop Priority Score","Likely Reason","Recommended Action"] if c in page_df.columns]
    limit = dynamic_report_limit(page_df)
    return page_df.sort_values(["Drop Priority Score","Clicks Drop","Impressions Drop"], ascending=False).head(limit)[columns].rename(columns={"Landing page":"Page"})

# ─────────────────────────────────────────────
# DATA LOADING — 3 MODES
# ─────────────────────────────────────────────
def load_from_csv_excel(uploaded_file):
    """Load from uploaded CSV or Excel file — returns list of (name, df) tuples"""
    name = uploaded_file.name
    if name.endswith(".csv"):
        df = pd.read_csv(uploaded_file, encoding="utf-8", on_bad_lines="skip")
        return [("Sheet1", df)]
    else:
        xl = pd.ExcelFile(uploaded_file)
        sheets = []
        for sheet in xl.sheet_names:
            try:
                df = xl.parse(sheet)
                sheets.append((sheet, df))
            except Exception:
                pass
        return sheets

def get_gspread_client(service_account_json):
    """Return authorized gspread client with full read+write scopes."""
    import gspread
    from google.oauth2.service_account import Credentials
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(json.loads(service_account_json), scopes=scopes)
    return gspread.authorize(creds)


def load_from_google_sheets_service_account(sheet_url, service_account_json):
    """Load via Service Account JSON"""
    try:
        gc = get_gspread_client(service_account_json)
        sh = gc.open_by_url(sheet_url)
        sheets = []
        for ws in sh.worksheets():
            data = ws.get_all_values()
            if data and len(data) >= 2:
                headers = make_unique_columns(data[0])
                df = pd.DataFrame(data[1:], columns=headers)
                sheets.append((ws.title, df))
        return sheets
    except Exception as e:
        raise RuntimeError(f"Service Account error: {e}")


def prepare_for_sheets(df):
    """Convert DataFrame to list-of-lists safe for Sheets upload."""
    df = df.copy()
    for col in df.columns:
        try:
            if isinstance(df[col].dtype, pd.CategoricalDtype):
                df[col] = df[col].astype(str)
        except Exception:
            pass
    df = df.replace([np.inf, -np.inf], np.nan).fillna("")
    # Convert all values to native Python types
    rows = []
    for row in df.values.tolist():
        rows.append([
            (float(v) if isinstance(v, (np.floating,)) else
             int(v)   if isinstance(v, (np.integer,))  else
             bool(v)  if isinstance(v, (np.bool_,))    else
             str(v)   if not isinstance(v, (str, int, float, bool)) else v)
            for v in row
        ])
    return [df.columns.tolist()] + rows


def save_results_to_google_sheet(results, overview_data, service_account_json, source_sheet_url=None):
    """
    Create a brand-new Google Sheet with all analysis tabs.
    Returns the URL of the new sheet.
    """
    from gspread.utils import rowcol_to_a1

    gc     = get_gspread_client(service_account_json)
    title  = f"Smart SEO Analysis — {datetime.now().strftime('%Y-%m-%d %H:%M')}"
    sh     = gc.create(title)

    # Make it accessible to anyone with the link (reader)
    try:
        sh.share("", perm_type="anyone", role="reader")
    except Exception:
        pass  # Drive API might not be enabled — non-fatal

    tab_order = list(results.items())

    # Header colour: purple #6b4ddf → RGB fractions
    HEADER_BG = {"red": 0.42, "green": 0.30, "blue": 0.87}
    HEADER_FG = {"red": 1.0,  "green": 1.0,  "blue": 1.0}

    first = True
    for tab_name, df in tab_order:
        if df is None or df.empty:
            df = pd.DataFrame({"Note": ["No data available."]})

        safe_name = re.sub(r"[\\/*?:\[\]]", "", tab_name)[:100]
        rows_n    = max(len(df) + 10, 20)
        cols_n    = max(len(df.columns) + 2, 5)

        if first:
            # Rename the default Sheet1 instead of adding a new tab
            ws = sh.sheet1
            ws.update_title(safe_name)
            ws = sh.get_worksheet(0)
            first = False
        else:
            ws = sh.add_worksheet(title=safe_name, rows=str(rows_n), cols=str(cols_n))

        data_to_write = prepare_for_sheets(df)
        ws.update("A1", data_to_write)

        # Format header row
        last_cell = rowcol_to_a1(1, len(df.columns))
        ws.format(f"A1:{last_cell}", {
            "backgroundColor": HEADER_BG,
            "textFormat": {
                "foregroundColor": HEADER_FG,
                "bold": True,
            },
            "horizontalAlignment": "CENTER",
        })
        ws.freeze(rows=1)

    return sh.url

def load_from_google_sheets_public(sheet_url):
    """Load via public CSV export (no auth needed for public sheets)"""
    try:
        import requests
        # Extract sheet ID
        match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", sheet_url)
        if not match:
            raise ValueError("Invalid Google Sheets URL")
        sheet_id = match.group(1)
        # Try to get gid (tab id)
        gid_match = re.search(r"[#&]gid=(\d+)", sheet_url)
        gid = gid_match.group(1) if gid_match else "0"
        csv_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}"
        r = requests.get(csv_url, timeout=30)
        if r.status_code != 200:
            raise ValueError(f"Could not access sheet. Make sure it's shared publicly. Status: {r.status_code}")
        df = pd.read_csv(io.StringIO(r.text), on_bad_lines="skip")
        return [("Sheet1", df)]
    except Exception as e:
        raise RuntimeError(f"Public sheet error: {e}")

def process_sheets(raw_sheets):
    """Process raw (name, df) tuples into analysis sources"""
    sources = []
    comparison_sources = []
    for name, raw_df in raw_sheets:
        try:
            cleaned = clean_data(raw_df.copy())
            ns = source_score(cleaned, name)
            sources.append({"name": name, "df": cleaned, "score": ns})
            comp_cleaned = clean_comparison_data(raw_df.copy())
            cs = comparison_score(comp_cleaned, name)
            if cs > 0:
                comparison_sources.append({"name": name, "df": comp_cleaned, "score": cs})
        except Exception:
            pass
    sources = sorted(sources, key=lambda x: x["score"], reverse=True)
    comparison_sources = sorted(comparison_sources, key=lambda x: x["score"], reverse=True)
    return sources, comparison_sources

# ─────────────────────────────────────────────
# PDF EXPORT
# ─────────────────────────────────────────────
def generate_pdf_report(overview_data, results):
    try:
        from fpdf import FPDF
        pdf = FPDF()
        pdf.add_page()
        pdf.set_auto_page_break(auto=True, margin=15)
        # Title
        pdf.set_font("Helvetica", "B", 24)
        pdf.set_text_color(107, 77, 223)
        pdf.cell(0, 15, "Smart SEO Analysis Report", ln=True, align="C")
        pdf.set_font("Helvetica", "", 11)
        pdf.set_text_color(100, 100, 100)
        pdf.cell(0, 8, f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}", ln=True, align="C")
        pdf.ln(5)
        # Overview metrics
        pdf.set_font("Helvetica", "B", 14)
        pdf.set_text_color(30, 35, 64)
        pdf.cell(0, 10, "Overview Metrics", ln=True)
        pdf.set_font("Helvetica", "", 10)
        pdf.set_text_color(50, 50, 50)
        for k, v in overview_data.items():
            pdf.cell(80, 8, k + ":", border=0)
            pdf.cell(0, 8, str(v), ln=True)
        pdf.ln(5)
        # Results summary
        for tab_name, df in results.items():
            if df is None or "Note" in df.columns: continue
            pdf.set_font("Helvetica", "B", 13)
            pdf.set_text_color(107, 77, 223)
            pdf.cell(0, 10, tab_name, ln=True)
            pdf.set_font("Helvetica", "", 9)
            pdf.set_text_color(50, 50, 50)
            pdf.cell(0, 6, f"Rows: {len(df):,} | Columns: {', '.join(df.columns[:5].tolist())}", ln=True)
            pdf.ln(3)
        buf = io.BytesIO()
        pdf_bytes = pdf.output()
        return bytes(pdf_bytes)
    except ImportError:
        return None

# ─────────────────────────────────────────────
# CHARTS
# ─────────────────────────────────────────────
CHART_THEME = {
    "paper_bgcolor": "rgba(0,0,0,0)",
    "plot_bgcolor":  "rgba(0,0,0,0)",
    "font":          {"color": "#e8eaf6", "family": "DM Sans"},
    "margin":        {"l":40,"r":40,"t":50,"b":40},
    "colorway":      ["#6b4ddf","#00d4aa","#f59e0b","#ef4444","#8a6be8","#10b981","#3b82f6","#ec4899"],
}

def apply_dark_theme(fig):
    fig.update_layout(
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(25,28,55,0.6)",
        font=dict(color="#e8eaf6", family="DM Sans"),
        margin=dict(l=40,r=40,t=50,b=40),
        legend=dict(bgcolor="rgba(0,0,0,0)", bordercolor="rgba(255,255,255,.1)", borderwidth=1),
        xaxis=dict(gridcolor="rgba(255,255,255,.05)", linecolor="rgba(255,255,255,.1)"),
        yaxis=dict(gridcolor="rgba(255,255,255,.05)", linecolor="rgba(255,255,255,.1)"),
    )
    return fig

def chart_position_distribution(df):
    if df is None or "Note" in df.columns: return None
    labels = ["#1","#2-3","#4-10","#11-20","#21-50","#51-100","+100"]
    df_sorted = df.set_index("Position Range").reindex([l for l in labels if l in df["Position Range"].values]).reset_index()
    colors = ["#00d4aa","#10b981","#6b4ddf","#8a6be8","#f59e0b","#ef4444","#4b5563"]
    fig = go.Figure(go.Bar(
        x=df_sorted["Position Range"], y=df_sorted["Clicks"],
        marker_color=colors[:len(df_sorted)], text=df_sorted["Clicks"].apply(lambda x: f"{x:,.0f}"),
        textposition="outside", textfont=dict(color="#e8eaf6", size=11),
    ))
    fig.update_layout(title="Click Distribution by SERP Position", **{k:v for k,v in CHART_THEME.items() if k!="colorway"})
    return apply_dark_theme(fig)

def chart_intent_breakdown(df):
    if df is None or "Note" in df.columns: return None
    colors = ["#6b4ddf","#00d4aa","#f59e0b","#ef4444","#8a6be8","#10b981","#3b82f6","#ec4899"]
    fig = go.Figure(go.Pie(
        labels=df["Intent"], values=df["Clicks"],
        hole=0.55, marker=dict(colors=colors[:len(df)], line=dict(color="rgba(0,0,0,0)",width=2)),
        textinfo="label+percent", textfont=dict(size=12, color="#e8eaf6"),
    ))
    fig.update_layout(title="Clicks by Search Intent", **{k:v for k,v in CHART_THEME.items() if k!="colorway"})
    return apply_dark_theme(fig)

def chart_opportunity_scatter(df):
    if df is None or "Note" in df.columns: return None
    display_df = df.head(50).copy()
    fig = px.scatter(
        display_df, x="Average Position", y="Opportunity Score",
        size="Impressions", color="Opportunity Score",
        hover_name="Keyword", color_continuous_scale=["#1a1e35","#6b4ddf","#00d4aa"],
        size_max=40, labels={"Average Position":"Avg Position","Opportunity Score":"Score"},
    )
    fig.update_layout(title="Opportunity Map (Size = Impressions)", **{k:v for k,v in CHART_THEME.items() if k!="colorway"})
    fig.update_coloraxes(showscale=False)
    return apply_dark_theme(fig)

def chart_top_pages(df):
    if df is None or "Note" in df.columns: return None
    top10 = df.head(10).copy()
    fig = go.Figure(go.Bar(
        x=top10["Clicks"], y=top10["Page"].apply(lambda x: x[-45:] if len(str(x))>45 else x),
        orientation="h", marker=dict(color=top10["Clicks"], colorscale=[[0,"#1a1e35"],[1,"#6b4ddf"]]),
        text=top10["Clicks"].apply(lambda x: f"{x:,.0f}"), textposition="outside",
        textfont=dict(color="#e8eaf6"),
    ))
    fig.update_layout(title="Top 10 Pages by Clicks", yaxis=dict(autorange="reversed"), **{k:v for k,v in CHART_THEME.items() if k!="colorway"})
    return apply_dark_theme(fig)

def chart_brand_vs_nonbrand(df):
    if df is None or "Note" in df.columns or len(df)==0: return None
    fig = go.Figure()
    metrics = ["Clicks","Impressions","Unique Keywords"]
    for metric in metrics:
        if metric not in df.columns: continue
    fig = make_subplots(rows=1, cols=2, specs=[[{"type":"pie"},{"type":"bar"}]], subplot_titles=["Clicks Share","Position Comparison"])
    clicks_data = df[["Segment","Clicks"]] if "Segment" in df.columns else None
    if clicks_data is not None:
        fig.add_trace(go.Pie(labels=df["Segment"],values=df["Clicks"],hole=0.5,marker=dict(colors=["#6b4ddf","#00d4aa"]),showlegend=False), row=1,col=1)
    if "Position" in df.columns:
        fig.add_trace(go.Bar(x=df["Segment"],y=df["Position"],marker_color=["#6b4ddf","#00d4aa"],showlegend=False), row=1,col=2)
    fig.update_layout(title="Brand vs Non-brand Analysis", **{k:v for k,v in CHART_THEME.items() if k!="colorway"})
    return apply_dark_theme(fig)

def chart_declining_keywords(df):
    if df is None or "Note" in df.columns or len(df)==0: return None
    top = df.head(15).copy()
    kw_col = "Keyword" if "Keyword" in top.columns else top.columns[0]
    drop_col = "Clicks Drop" if "Clicks Drop" in top.columns else None
    if drop_col is None: return None
    fig = go.Figure(go.Bar(
        x=top[drop_col], y=top[kw_col].apply(lambda x: x[:35]+"..." if len(str(x))>35 else x),
        orientation="h", marker=dict(color=top[drop_col], colorscale=[[0,"#f59e0b"],[1,"#ef4444"]]),
        text=top[drop_col].apply(lambda x: f"-{x:,.0f}"), textposition="outside",
        textfont=dict(color="#e8eaf6"),
    ))
    fig.update_layout(title="Top 15 Declining Keywords (Click Loss)", yaxis=dict(autorange="reversed"),
                      **{k:v for k,v in CHART_THEME.items() if k!="colorway"})
    return apply_dark_theme(fig)

def chart_ctr_problems(df):
    if df is None or "Note" in df.columns or len(df)==0: return None
    group_col = "Landing page" if "Landing page" in df.columns else df.columns[0]
    top = df.head(12).copy()
    fig = go.Figure()
    x_labels = top[group_col].apply(lambda x: str(x)[-35:] if len(str(x))>35 else str(x))
    if "CTR (%)" in top.columns:
        fig.add_trace(go.Bar(name="Current CTR %", x=x_labels, y=top["CTR (%)"], marker_color="#f59e0b"))
    if "Expected CTR" in top.columns:
        fig.add_trace(go.Scatter(name="Expected CTR %", x=x_labels, y=top["Expected CTR"]*100,
                                 mode="markers+lines", marker=dict(color="#00d4aa",size=8), line=dict(color="#00d4aa",dash="dash")))
    fig.update_layout(title="CTR Problems: Current vs Expected", barmode="group",
                      **{k:v for k,v in CHART_THEME.items() if k!="colorway"})
    return apply_dark_theme(fig)

# ─────────────────────────────────────────────
# EXCEL EXPORT
# ─────────────────────────────────────────────
def export_to_excel(results):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for name, df in results.items():
            if df is not None and not df.empty:
                safe_name = re.sub(r"[\\/*?:\[\]]","",name)[:31]
                df.to_excel(writer, sheet_name=safe_name, index=False)
    buf.seek(0)
    return buf.getvalue()

# ─────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────
def render_sidebar():
    with st.sidebar:
        st.markdown("""
        <div style='text-align:center;padding:20px 0 10px;'>
            <div style='font-family:Syne,sans-serif;font-size:22px;font-weight:800;
                        background:linear-gradient(135deg,#fff,#8a6be8);
                        -webkit-background-clip:text;-webkit-text-fill-color:transparent;'>
                📊 Smart SEO
            </div>
            <div style='color:#8892b0;font-size:12px;margin-top:4px;'>Analysis Tool v2.0</div>
        </div>
        <hr style='border-color:#2a2f50;margin:10px 0;'>
        """, unsafe_allow_html=True)

        st.markdown("### 🔗 Data Source")
        mode = st.radio("Choose input method:", [
            "📁 Upload CSV / Excel",
            "🌐 Google Sheets (Public)",
            "🔐 Google Sheets (Service Account)",
        ], label_visibility="collapsed")

        st.markdown("<hr style='border-color:#2a2f50;margin:10px 0;'>", unsafe_allow_html=True)
        st.markdown("### ⚙️ Settings")
        min_impressions = st.number_input("Min impressions for opportunities", min_value=10, value=100, step=10)
        min_ctr_impressions = st.number_input("Min impressions for CTR problems", min_value=50, value=500, step=50)
        CFG.min_impressions_for_opportunity = min_impressions
        CFG.min_impressions_for_ctr_problem = min_ctr_impressions

        st.markdown("<hr style='border-color:#2a2f50;margin:10px 0;'>", unsafe_allow_html=True)
        st.markdown("""
        <div style='color:#8892b0;font-size:11px;line-height:1.7;'>
        <strong style='color:#e8eaf6;'>📖 Supported formats:</strong><br>
        • Google Search Console exports<br>
        • Looker Studio exports<br>
        • CSV / Excel with queries, pages, clicks, impressions, CTR, position<br>
        • Multi-tab Google Sheets<br>
        • Comparison tabs (current vs previous)
        </div>
        """, unsafe_allow_html=True)

    return mode

# ─────────────────────────────────────────────
# MAIN APP
# ─────────────────────────────────────────────
def main():
    # Hero banner
    st.markdown("""
    <div class="hero-banner">
        <div class="hero-title">📊 Smart SEO Analysis</div>
        <div class="hero-sub">Upload your data, get instant insights — keyword clusters, opportunities, cannibalization, declines & more</div>
    </div>
    """, unsafe_allow_html=True)

    mode = render_sidebar()

    # ── DATA INPUT SECTION ──
    raw_sheets = None
    load_error = None

    if mode == "📁 Upload CSV / Excel":
        st.markdown('<div class="section-header">📁 Upload Your Data File</div>', unsafe_allow_html=True)
        col1, col2 = st.columns([2,1])
        with col1:
            uploaded = st.file_uploader(
                "Drop your CSV or Excel file here",
                type=["csv","xlsx","xls"],
                help="Export from Google Search Console → Performance → Download CSV"
            )
            if uploaded:
                with st.spinner("Reading file..."):
                    try:
                        raw_sheets = load_from_csv_excel(uploaded)
                        st.success(f"✅ Loaded **{uploaded.name}** — {len(raw_sheets)} sheet(s)")
                    except Exception as e:
                        load_error = str(e)
        with col2:
            st.markdown("""
            <div style='background:rgba(107,77,223,.08);border:1px solid rgba(107,77,223,.3);
                        border-radius:16px;padding:20px;margin-top:4px;'>
                <div style='font-weight:700;color:#8a6be8;margin-bottom:10px;font-size:14px;'>💡 Tips</div>
                <div style='color:#8892b0;font-size:12px;line-height:1.8;'>
                    • Export from GSC → Performance<br>
                    • Include: Query, Page, Clicks, Impressions, CTR, Position<br>
                    • For comparison: add a 2nd date range in GSC<br>
                    • Arabic & English supported<br>
                    • Multi-tab Excel fully supported
                </div>
            </div>
            """, unsafe_allow_html=True)

    elif mode == "🌐 Google Sheets (Public)":
        st.markdown('<div class="section-header">🌐 Google Sheets — Public Access</div>', unsafe_allow_html=True)
        st.info("ℹ️ The sheet must be shared publicly (Anyone with the link → Viewer).")
        col1, col2 = st.columns([3,1])
        with col1:
            sheet_url = st.text_input("Google Sheets URL", placeholder="https://docs.google.com/spreadsheets/d/...")
        with col2:
            st.write("")
            st.write("")
            load_btn = st.button("🚀 Load Sheet", use_container_width=True)
        if load_btn and sheet_url:
            with st.spinner("Fetching sheet..."):
                try:
                    raw_sheets = load_from_google_sheets_public(sheet_url)
                    st.success(f"✅ Sheet loaded — {len(raw_sheets)} tab(s)")
                except Exception as e:
                    load_error = str(e)

    elif mode == "🔐 Google Sheets (Service Account)":
        st.markdown('<div class="section-header">🔐 Google Sheets — Service Account</div>', unsafe_allow_html=True)

        with st.expander("ℹ️ How to set up a Service Account", expanded=False):
            st.markdown("""
            1. Go to [Google Cloud Console](https://console.cloud.google.com/)
            2. Create a project → Enable **Google Sheets API** and **Google Drive API**
            3. Create a **Service Account** → Download JSON key
            4. Share your Google Sheet with the service account email (`xxx@xxx.iam.gserviceaccount.com`) as Viewer
            5. Paste the JSON content below
            """)

        col1, col2 = st.columns([1,1])
        with col1:
            sheet_url = st.text_input("Google Sheets URL", placeholder="https://docs.google.com/spreadsheets/d/...")
        with col2:
            sa_json = st.text_area("Service Account JSON", placeholder='{"type": "service_account", ...}', height=120)
        load_btn = st.button("🔐 Connect & Load", use_container_width=False)
        if load_btn and sheet_url and sa_json:
            with st.spinner("Connecting via Service Account..."):
                try:
                    raw_sheets = load_from_google_sheets_service_account(sheet_url, sa_json)
                    st.session_state["sa_json_connected"] = sa_json  # save for output sheet
                    st.success(f"✅ Connected — {len(raw_sheets)} tab(s) loaded")
                except Exception as e:
                    load_error = str(e)

    if load_error:
        st.error(f"❌ {load_error}")

    if raw_sheets is None:
        # Landing state
        st.markdown("""
        <div style='text-align:center;padding:60px 20px;color:#8892b0;'>
            <div style='font-size:48px;margin-bottom:16px;'>🔍</div>
            <div style='font-family:Syne,sans-serif;font-size:20px;font-weight:700;color:#e8eaf6;margin-bottom:8px;'>
                Ready to Analyze Your SEO Data
            </div>
            <div style='font-size:14px;max-width:500px;margin:0 auto;line-height:1.7;'>
                Choose a data source from the sidebar. Supports Google Search Console exports,
                Looker Studio CSVs, or connect directly to Google Sheets.
            </div>
        </div>
        """, unsafe_allow_html=True)
        return

    # ── PROCESS DATA ──
    with st.spinner("🧠 Detecting columns and processing data..."):
        sources, comparison_sources = process_sheets(raw_sheets)

    if not sources:
        st.error("❌ Could not detect valid SEO columns. Please check your file format.")
        return

    # Pick best sources
    query_page_s = next((s for s in sources if "Query string" in s["df"].columns and "Landing page" in s["df"].columns), None)
    query_s      = next((s for s in sources if "Query string" in s["df"].columns), None)
    page_s       = next((s for s in sources if "Landing page" in s["df"].columns), None)
    comp_s       = comparison_sources[0] if comparison_sources else None

    main_df       = query_page_s["df"] if query_page_s else sources[0]["df"]
    query_df      = query_s["df"]      if query_s      else main_df
    page_df       = page_s["df"]       if page_s       else main_df
    query_page_df = query_page_s["df"] if query_page_s else None
    comparison_df = comp_s["df"]       if comp_s       else None

    # ── COLUMN DETECTION INFO ──
    detected_cols = [c for c in ["Query string","Landing page","Clicks","Impressions","CTR","Position"] if c in main_df.columns]
    if detected_cols:
        cols_str = " · ".join([f"**{c}**" for c in detected_cols])
        st.success(f"🧠 Detected columns: {cols_str}")
    if comparison_df is not None:
        st.info(f"📉 Comparison tab detected: **{comp_s['name']}**")

    # ── RUN ALL ANALYSES ──
    with st.spinner("⚡ Running all analyses..."):
        overview_data    = create_overview(main_df)
        top_pages_df     = create_top_pages(page_df)
        top_kw_df        = create_smart_top_keywords(query_df)
        brand_df         = create_brand_nonbrand(query_df)
        intent_df        = create_intent_analysis(query_df)
        opportunity_df   = create_opportunity_score_report(query_df)
        ctr_problems_df  = create_ctr_problems(main_df)
        cannib_df        = create_cannibalization(query_page_df)
        position_df      = create_position_distribution(query_df)
        declining_kw_df  = create_declining_keywords(comparison_df)
        declining_pg_df  = create_declining_pages(comparison_df)

    all_results = {
        "1. Overview":                pd.DataFrame(list(overview_data.items()), columns=["Metric","Value"]),
        "2. Top Pages":               top_pages_df,
        "3. Smart Top Keywords":      top_kw_df,
        "4. Brand vs Non-brand":      brand_df,
        "5. Intent Analysis":         intent_df,
        "6. Opportunity Score":       opportunity_df,
        "7. CTR Problems":            ctr_problems_df,
        "8. Keyword Cannibalization": cannib_df,
        "9. Position Distribution":   position_df,
        "10. Declining Keywords":     declining_kw_df,
        "11. Declining Pages":        declining_pg_df,
    }

    # ─────────────────────────────
    # KPI METRICS ROW
    # ─────────────────────────────
    st.markdown('<div class="section-header">📈 Key Metrics</div>', unsafe_allow_html=True)
    m1,m2,m3,m4,m5,m6 = st.columns(6)
    kpis = [
        (m1, "Total Clicks",      overview_data.get("Total Clicks","—"),      "🖱️"),
        (m2, "Impressions",       overview_data.get("Total Impressions","—"),  "👁️"),
        (m3, "Avg CTR",           overview_data.get("Avg CTR (%)","—"),        "📊"),
        (m4, "Avg Position",      overview_data.get("Avg Position","—"),       "📍"),
        (m5, "Unique Queries",    overview_data.get("Unique Queries","—"),     "🔍"),
        (m6, "Unique Pages",      overview_data.get("Unique Pages","—"),       "📄"),
    ]
    for col, label, value, icon in kpis:
        with col:
            st.metric(f"{icon} {label}", value)

    # ─────────────────────────────
    # ALERTS
    # ─────────────────────────────
    alert_cols = st.columns(3)
    with alert_cols[0]:
        if cannib_df is not None and "Note" not in cannib_df.columns and len(cannib_df)>0:
            st.error(f"⚠️ **{len(cannib_df)} cannibalization** cases detected")
    with alert_cols[1]:
        if declining_kw_df is not None and "Note" not in declining_kw_df.columns and len(declining_kw_df)>0:
            st.warning(f"📉 **{len(declining_kw_df)} keywords** declining")
    with alert_cols[2]:
        if opportunity_df is not None and "Note" not in opportunity_df.columns and len(opportunity_df)>0:
            st.info(f"🚀 **{len(opportunity_df)} opportunities** found")

    # ─────────────────────────────
    # CHARTS DASHBOARD
    # ─────────────────────────────
    st.markdown('<div class="section-header">📊 Dashboard</div>', unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    with c1:
        fig = chart_position_distribution(position_df)
        if fig: st.plotly_chart(fig, use_container_width=True, config={"displayModeBar":False})
    with c2:
        fig = chart_intent_breakdown(intent_df)
        if fig: st.plotly_chart(fig, use_container_width=True, config={"displayModeBar":False})

    c3, c4 = st.columns(2)
    with c3:
        fig = chart_top_pages(top_pages_df)
        if fig: st.plotly_chart(fig, use_container_width=True, config={"displayModeBar":False})
    with c4:
        fig = chart_brand_vs_nonbrand(brand_df)
        if fig: st.plotly_chart(fig, use_container_width=True, config={"displayModeBar":False})

    c5, c6 = st.columns(2)
    with c5:
        fig = chart_opportunity_scatter(opportunity_df)
        if fig: st.plotly_chart(fig, use_container_width=True, config={"displayModeBar":False})
    with c6:
        fig = chart_ctr_problems(ctr_problems_df)
        if fig: st.plotly_chart(fig, use_container_width=True, config={"displayModeBar":False})

    if declining_kw_df is not None and "Note" not in declining_kw_df.columns:
        fig = chart_declining_keywords(declining_kw_df)
        if fig: st.plotly_chart(fig, use_container_width=True, config={"displayModeBar":False})

    # ─────────────────────────────
    # DATA TABLES — TABS
    # ─────────────────────────────
    st.markdown('<div class="section-header">📋 Detailed Reports</div>', unsafe_allow_html=True)

    tab_labels = ["🏆 Top Keywords","📄 Top Pages","🎯 Opportunities","⚠️ CTR Problems",
                  "🔄 Cannibalization","📉 Declining KWs","📉 Declining Pages",
                  "🏷️ Brand / Non-brand","💡 Intent","📍 Position Dist."]
    tab_dfs = [top_kw_df, top_pages_df, opportunity_df, ctr_problems_df,
               cannib_df, declining_kw_df, declining_pg_df,
               brand_df, intent_df, position_df]

    tabs = st.tabs(tab_labels)
    for tab, df in zip(tabs, tab_dfs):
        with tab:
            if df is None or "Note" in df.columns:
                st.info("ℹ️ " + (df["Note"].iloc[0] if df is not None and "Note" in df.columns else "No data available"))
            else:
                # Search filter
                search = st.text_input("🔍 Filter rows", key=f"search_{tab_labels[tab_dfs.index(df)]}", placeholder="Type to filter...")
                display_df = df.copy()
                if search:
                    mask = display_df.apply(lambda col: col.astype(str).str.contains(search, case=False, na=False)).any(axis=1)
                    display_df = display_df[mask]
                st.markdown(f"<div style='color:#8892b0;font-size:12px;margin-bottom:8px;'>Showing {len(display_df):,} rows</div>", unsafe_allow_html=True)
                st.dataframe(display_df, use_container_width=True, height=400)

    # ─────────────────────────────
    # EXPORT SECTION
    # ─────────────────────────────
    st.markdown('<div class="section-header">💾 Export Reports</div>', unsafe_allow_html=True)

    # ── Row 1: Google Sheets Save (full width, primary CTA) ──
    st.markdown("""
    <div style='background:linear-gradient(135deg,rgba(107,77,223,.15),rgba(0,212,170,.08));
                border:1px solid rgba(107,77,223,.4);border-radius:20px;padding:28px;margin-bottom:20px;'>
        <div style='font-family:Syne,sans-serif;font-size:18px;font-weight:800;color:#e8eaf6;margin-bottom:6px;'>
            🗂️ Save to Google Sheets
        </div>
        <div style='color:#8892b0;font-size:13px;margin-bottom:18px;'>
            Creates a brand-new Google Sheet with all analysis tabs, formatted headers, and frozen rows —
            just like the original Colab tool.
        </div>
    </div>
    """, unsafe_allow_html=True)

    # Determine if SA JSON is already available from this session
    sa_json_for_save = st.session_state.get("sa_json_connected", "")

    if mode == "🔐 Google Sheets (Service Account)" and sa_json_for_save:
        # Already connected — one-click save
        if st.button("🚀 Create New Google Sheet with Results", use_container_width=True):
            with st.spinner("Creating your Google Sheet... this may take 30–60 seconds ⏳"):
                try:
                    new_url = save_results_to_google_sheet(all_results, overview_data, sa_json_for_save)
                    st.success("✅ Google Sheet created successfully!")
                    st.markdown(f"""
                    <a href="{new_url}" target="_blank"
                       style="display:inline-block;background:linear-gradient(135deg,#6b4ddf,#00d4aa);
                              color:#fff;padding:14px 28px;border-radius:12px;font-size:15px;
                              text-decoration:none;font-weight:700;margin-top:8px;">
                        📊 Open Your New Report Sheet →
                    </a>
                    """, unsafe_allow_html=True)
                except Exception as e:
                    st.error(f"❌ Failed to create sheet: {e}")
                    st.info("Make sure your Service Account has Editor access to Google Drive.")
    else:
        # Ask for SA JSON inline (works for CSV/Excel and Public modes too)
        with st.expander("🔐 Enter Service Account JSON to save to Google Sheets", expanded=(mode != "📁 Upload CSV / Excel")):
            st.markdown("""
            <div style='color:#8892b0;font-size:12px;line-height:1.7;margin-bottom:10px;'>
            Don't have a Service Account yet?<br>
            1. <a href='https://console.cloud.google.com/' target='_blank' style='color:#8a6be8;'>Google Cloud Console</a>
               → Enable <strong>Sheets API</strong> + <strong>Drive API</strong><br>
            2. Create Service Account → Download JSON key<br>
            3. Paste the JSON content below
            </div>
            """, unsafe_allow_html=True)
            sa_input = st.text_area(
                "Service Account JSON",
                placeholder='{"type": "service_account", "project_id": "...", ...}',
                height=130,
                key="sa_json_for_save_input",
                label_visibility="collapsed",
            )
            if st.button("🚀 Create New Google Sheet with Results", use_container_width=True, key="save_sheet_btn"):
                if not sa_input.strip():
                    st.warning("Please paste your Service Account JSON first.")
                else:
                    with st.spinner("Creating your Google Sheet... this may take 30–60 seconds ⏳"):
                        try:
                            new_url = save_results_to_google_sheet(all_results, overview_data, sa_input.strip())
                            st.success("✅ Google Sheet created successfully!")
                            st.markdown(f"""
                            <a href="{new_url}" target="_blank"
                               style="display:inline-block;background:linear-gradient(135deg,#6b4ddf,#00d4aa);
                                      color:#fff;padding:14px 28px;border-radius:12px;font-size:15px;
                                      text-decoration:none;font-weight:700;margin-top:8px;">
                                📊 Open Your New Report Sheet →
                            </a>
                            """, unsafe_allow_html=True)
                        except Exception as e:
                            st.error(f"❌ Failed: {e}")
                            st.info("Tip: Make sure Sheets API and Drive API are enabled in your Google Cloud project.")

    st.markdown("<div style='height:8px;'></div>", unsafe_allow_html=True)

    # ── Row 2: Excel + PDF downloads ──
    ex1, ex2 = st.columns(2)

    with ex1:
        st.markdown("#### 📊 Excel Report")
        st.markdown("<div style='color:#8892b0;font-size:13px;margin-bottom:12px;'>All tabs in one Excel file</div>", unsafe_allow_html=True)
        excel_bytes = export_to_excel(all_results)
        st.download_button(
            label="⬇️ Download Excel (.xlsx)",
            data=excel_bytes,
            file_name=f"smart_seo_analysis_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    with ex2:
        st.markdown("#### 📄 PDF Summary")
        st.markdown("<div style='color:#8892b0;font-size:13px;margin-bottom:12px;'>Executive summary for sharing</div>", unsafe_allow_html=True)
        pdf_bytes = generate_pdf_report(overview_data, all_results)
        if pdf_bytes:
            st.download_button(
                label="⬇️ Download PDF Summary",
                data=pdf_bytes,
                file_name=f"smart_seo_summary_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
                mime="application/pdf",
                use_container_width=True,
            )
        else:
            st.warning("PDF export requires `fpdf2` — run `pip install fpdf2`")

    # ─────────────────────────────
    # FOOTER
    # ─────────────────────────────
    st.markdown("""
    <div style='text-align:center;padding:40px 20px;color:#4b5563;font-size:12px;border-top:1px solid #2a2f50;margin-top:40px;'>
        <span style='color:#6b4ddf;font-weight:700;'>Smart SEO Analysis Tool</span> —
        Supports Google Search Console · Looker Studio · Multi-tab Sheets · Arabic & English
    </div>
    """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()
