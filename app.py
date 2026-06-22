"""
MBC Media Solutions — Talent 9-Grid Dashboard  v3
==================================================
Source : 9Grid_Final.xlsx  →  sheets  Evalutaion 23 / 24 / 25
Brand  : #666EFF (purple) · #30BFA6 (green) · #FFFFFF · #EDEDEE · #2B2B2B
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import base64, warnings
warnings.filterwarnings("ignore")

# ── PAGE CONFIG ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="MMS Talent 9-Grid",
    page_icon="🎯",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── BRAND PALETTE ─────────────────────────────────────────────────────────────
C_PURPLE  = "#666EFF"
C_GREEN   = "#30BFA6"
C_WHITE   = "#FFFFFF"
C_LGRAY   = "#EDEDEE"
C_DARK    = "#2B2B2B"
C_SIDEBAR = "#1A1A2E"          # deep navy for sidebar
C_CANVAS  = "#F7F7FB"          # page background

# 9-box palette  (6 shades derived from brand purple→green gradient + accent reds)
BOX_COLORS = {
    "3C": {"dot": "#666EFF", "bg": "rgba(102,110,255,0.10)", "border": "#666EFF"},
    "3B": {"dot": "#30BFA6", "bg": "rgba(48,191,166,0.10)",  "border": "#30BFA6"},
    "3A": {"dot": "#4ECDC4", "bg": "rgba(78,205,196,0.10)",  "border": "#4ECDC4"},
    "2C": {"dot": "#A78BFA", "bg": "rgba(167,139,250,0.10)", "border": "#A78BFA"},
    "2B": {"dot": "#6B7280", "bg": "rgba(107,114,128,0.08)", "border": "#9CA3AF"},
    "2A": {"dot": "#F59E0B", "bg": "rgba(245,158,11,0.10)",  "border": "#F59E0B"},
    "1C": {"dot": "#F97316", "bg": "rgba(249,115,22,0.10)",  "border": "#F97316"},
    "1B": {"dot": "#EF4444", "bg": "rgba(239,68,68,0.10)",   "border": "#EF4444"},
    "1A": {"dot": "#991B1B", "bg": "rgba(153,27,27,0.13)",   "border": "#991B1B"},
}

GRID_META = {
    "3C": {"label": "Star",               "perf": 3, "pot": 3},
    "3B": {"label": "High Performer",     "perf": 3, "pot": 2},
    "3A": {"label": "Solid Performer",    "perf": 3, "pot": 1},
    "2C": {"label": "Growth Employee",    "perf": 2, "pot": 3},
    "2B": {"label": "Core Player",        "perf": 2, "pot": 2},
    "2A": {"label": "Average Performer",  "perf": 2, "pot": 1},
    "1C": {"label": "Potential Gem",      "perf": 1, "pot": 3},
    "1B": {"label": "Inconsistent Player","perf": 1, "pot": 2},
    "1A": {"label": "Risk",               "perf": 1, "pot": 1},
}
VALID_CODES = set(GRID_META.keys())
CAT_ORDER   = ["3C","3B","3A","2C","2B","2A","1C","1B","1A"]

LOC_MAP = {
    "MMS KSA": "Saudi Arabia",
    "MMS UAE": "UAE",
    "MMS Egypt": "Egypt",
}

# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown(f"""
<style>
/* ── canvas ── */
[data-testid="stAppViewContainer"] {{ background:{C_CANVAS}; }}
[data-testid="stMainBlockContainer"] {{ padding-top:1rem; }}

/* ── sidebar ── */
[data-testid="stSidebar"] {{ background:{C_SIDEBAR} !important; }}
[data-testid="stSidebar"] > div:first-child {{
    background:{C_SIDEBAR} !important;
    padding-top:0 !important;
}}
[data-testid="stSidebar"] * {{ color:#CBD5E1 !important; }}
[data-testid="stSidebar"] h1,
[data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3 {{ color:#F1F5F9 !important; }}
[data-testid="stSidebar"] .stSelectbox label,
[data-testid="stSidebar"] .stMultiSelect label {{
    color:#94A3B8 !important; font-size:11px !important;
    text-transform:uppercase; letter-spacing:.05em;
}}
[data-testid="stSidebar"] hr {{ border-color:rgba(255,255,255,.1) !important; }}
[data-testid="stSidebar"] .stSelectbox > div > div {{
    border:1.5px solid rgba(102,110,255,0.45) !important;
    border-radius:8px !important;
    background:rgba(255,255,255,0.05) !important;
}}
[data-testid="stSidebar"] .stSelectbox > div > div:focus-within {{
    border-color:{C_PURPLE} !important;
    box-shadow:0 0 0 2px rgba(102,110,255,0.25) !important;
}}
/* purple border for all selectboxes everywhere */
.stSelectbox > div > div {{
    border:1.5px solid rgba(102,110,255,0.35) !important;
    border-radius:8px !important;
}}
.stSelectbox > div > div:focus-within {{
    border-color:{C_PURPLE} !important;
    box-shadow:0 0 0 2px rgba(102,110,255,0.20) !important;
}}
/* ── single purple tab underline (no duplicate, no red) ── */

/* ── hide default header ── */
header[data-testid="stHeader"] {{
    background: transparent !important;
}}

/* ── KPI cards ── */
.kpi-wrap {{ background:#fff; border-radius:14px; padding:16px 18px;
             box-shadow:0 2px 8px rgba(0,0,0,.06);
             border-top:3px solid {C_PURPLE}; }}
.kpi-wrap.green  {{ border-top-color:{C_GREEN}; }}
.kpi-wrap.amber  {{ border-top-color:#F59E0B; }}
.kpi-wrap.red    {{ border-top-color:#991B1B; }}
.kpi-wrap.violet {{ border-top-color:#A78BFA; }}
.kpi-wrap.gray   {{ border-top-color:#6B7280; }}
.kpi-lbl  {{ font-size:10px; font-weight:700; color:#6B7280;
             text-transform:uppercase; letter-spacing:.07em; margin-bottom:4px; }}
.kpi-val  {{ font-size:28px; font-weight:800; color:{C_DARK}; line-height:1; }}
.kpi-sub  {{ font-size:10px; color:#9CA3AF; margin-top:3px; }}
.kpi-dlt  {{ font-size:11px; font-weight:700; margin-top:2px; }}
.kpi-dlt.up {{ color:{C_GREEN}; }}
.kpi-dlt.dn {{ color:#EF4444; }}

/* ── section header ── */
.sec-head {{
    font-size:15px; font-weight:800; color:{C_DARK};
    display:flex; align-items:center; gap:8px;
    border-bottom:2px solid {C_PURPLE}; padding-bottom:6px;
    margin:24px 0 14px;
}}
.sec-head .pill {{
    background:{C_PURPLE}; color:#fff; font-size:10px;
    padding:2px 8px; border-radius:999px; font-weight:700;
}}

/* ── top performer cards ── */
.tp-card {{
    background:#fff; border-radius:12px; padding:14px 16px;
    box-shadow:0 2px 8px rgba(0,0,0,.06);
    border-left:4px solid {C_PURPLE};
    margin-bottom:10px;
}}
.tp-card.green {{ border-left-color:{C_GREEN}; }}
.tp-card.amber {{ border-left-color:#F59E0B; }}
.tp-name  {{ font-size:13px; font-weight:700; color:{C_DARK}; }}
.tp-role  {{ font-size:11px; color:#6B7280; margin-top:1px; }}
.tp-meta  {{ font-size:11px; color:#374151; margin-top:6px; line-height:1.6; }}
.tp-badge {{
    display:inline-block; font-size:10px; font-weight:700;
    padding:2px 9px; border-radius:999px; margin-top:5px;
}}

/* ── 9-box cell ── */
.box9 {{
    border-radius:10px; padding:10px 12px;
    display:flex; flex-direction:column;
    font-family:inherit;
}}
.box9-code  {{ font-size:11px; font-weight:800; opacity:.7; }}
.box9-label {{ font-size:12px; font-weight:700; margin:2px 0; }}
.box9-cnt   {{ font-size:26px; font-weight:800; line-height:1; }}
.box9-names {{ font-size:9.5px; line-height:1.55; margin-top:5px;
               max-height:90px; overflow:hidden; }}

/* ── pill badge ── */
.badge {{
    display:inline-block; font-size:10px; font-weight:700;
    padding:2px 9px; border-radius:999px;
}}

/* ── Tab bar: single clean purple underline ── */
.stTabs [data-baseweb="tab-list"] {{
    gap: 0;
    border-bottom: 1px solid #E5E7EB;
}}
.stTabs [data-baseweb="tab-list"] button {{
    border: none !important;
    border-bottom: 3px solid transparent !important;
    background: transparent !important;
    box-shadow: none !important;
    padding-bottom: 10px;
    margin-bottom: -1px;
}}
.stTabs [data-baseweb="tab-list"] button[aria-selected="true"] {{
    border-bottom: 3px solid #666EFF !important;
    color: #666EFF !important;
    font-weight: 700 !important;
    background: transparent !important;
}}
.stTabs [data-baseweb="tab-list"] button[aria-selected="true"] p {{
    color: #666EFF !important;
}}
.stTabs [data-baseweb="tab-list"] button:not([aria-selected="true"]) {{
    color: #6B7280 !important;
}}
.stTabs [data-baseweb="tab-list"] button:not([aria-selected="true"]) p {{
    color: #6B7280 !important;
}}
/* Hide the BaseWeb animated highlight bar (prevents double underline / red line) */
[data-baseweb="tab-highlight"] {{
    display: none !important;
    background: transparent !important;
    height: 0 !important;
}}




/* ── Final reference radio style: visible circles, no text background ── */
div[role="radiogroup"] {{
    display: flex !important;
    align-items: center !important;
    gap: 26px !important;
}}

div[role="radiogroup"] label {{
    display: inline-flex !important;
    align-items: center !important;
    gap: 10px !important;
    background: transparent !important;
    border: none !important;
    box-shadow: none !important;
    padding: 0 !important;
    margin: 0 !important;
    cursor: pointer !important;
}}

div[role="radiogroup"] label * {{
    background: transparent !important;
    box-shadow: none !important;
}}

div[role="radiogroup"] label p,
div[role="radiogroup"] label span {{
    color: #2B2B2B !important;
    font-weight: 400 !important;
    background: transparent !important;
    margin: 0 !important;
}}

div[role="radiogroup"] label:has(input[type="radio"]:checked) p,
div[role="radiogroup"] label:has(input[type="radio"]:checked) span {{
    color: #666EFF !important;
    font-weight: 700 !important;
    background: transparent !important;
}}

div[role="radiogroup"] input[type="radio"] {{
    appearance: auto !important;
    -webkit-appearance: radio !important;
    accent-color: #666EFF !important;
    width: 24px !important;
    height: 24px !important;
    min-width: 24px !important;
    min-height: 24px !important;
    margin: 0 !important;
    opacity: 1 !important;
    position: relative !important;
    visibility: visible !important;
    cursor: pointer !important;
}}

/* BaseWeb wrapper cleanup: do not color or cover the text */
div[role="radiogroup"] [data-baseweb="radio"] {{
    background: transparent !important;
    border: none !important;
    box-shadow: none !important;
}}

div[role="radiogroup"] [data-baseweb="radio"] div {{
    background: transparent !important;
    box-shadow: none !important;
}}

div[role="radiogroup"] label:hover p,
div[role="radiogroup"] label:hover span {{
    color: #666EFF !important;
}}

</style>
""", unsafe_allow_html=True)


# ── DATA LOADING ──────────────────────────────────────────────────────────────
@st.cache_data(show_spinner="Loading talent data…")
def load_all(path: str) -> pd.DataFrame:
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    frames = []

    SHEET_YEARS = [
        ("Evalutaion 23", 2023),
        ("Evalutaion 24", 2024),
        ("Evalutaion 25", 2025),
    ]

    for sheet, year in SHEET_YEARS:
        ws   = wb[sheet]
        rows = list(ws.iter_rows(values_only=True))
        hdr  = next(
            i for i, r in enumerate(rows)
            if any(c == "Employee Name" for c in r)
        )
        cols = list(rows[hdr])
        data = [r for r in rows[hdr + 1:] if any(c is not None for c in r[:4])]
        df   = pd.DataFrame(data, columns=cols)

        # ── find column names flexibly ──────────────────────────────────────
        def find(keywords, exclude=None):
            exc = exclude or []
            for c in df.columns:
                cs = str(c).strip().lower() if c else ""
                if all(k in cs for k in keywords) and not any(e in cs for e in exc):
                    return c
            return None

        grid_col  = find(["zone"], exclude=["potential","perf","perfor"]) or \
                    find(["gird"]) or find(["grid"])
        score_col = find(["objectives", "comp"])
        obj_col   = find(["objectives", "score"], exclude=["comp","and"])
        cmp_col   = find(["competency", "score"], exclude=["objectives"])
        pot_col   = find(["potential"])
        perf_col  = find(["perfor", "zone"])
        loc_col   = "Location"   if "Location"         in df.columns else None
        dept_col  = "Department" if "Department"        in df.columns else None
        name_col  = "Employee Name"
        pos_col   = "Position"   if "Position"          in df.columns else None
        empid_col = "Employee Number" if "Employee Number" in df.columns else None

        # ── rename to standard names ────────────────────────────────────────
        rn = {}
        if grid_col:  rn[grid_col]  = "grid_zone"
        if score_col: rn[score_col] = "overall_score"
        if obj_col:   rn[obj_col]   = "obj_score"
        if cmp_col:   rn[cmp_col]   = "comp_score"
        if pot_col:   rn[pot_col]   = "pot_zone_raw"
        if perf_col:  rn[perf_col]  = "perf_zone_raw"
        if loc_col:   rn[loc_col]   = "location"
        if dept_col:  rn[dept_col]  = "dept"
        if pos_col:   rn[pos_col]   = "position"
        if empid_col: rn[empid_col] = "emp_id"
        rn[name_col] = "name"
        df = df.rename(columns=rn)

        # ── clean grid_zone ─────────────────────────────────────────────────
        if "grid_zone" in df.columns:
            df["grid_zone"] = df["grid_zone"].apply(
                lambda v: str(v).strip()
                if isinstance(v, str) and str(v).strip() in VALID_CODES
                else (str(v).strip() if isinstance(v, str) and str(v).strip() in VALID_CODES else np.nan)
            )
        else:
            df["grid_zone"] = np.nan

        # ── clean numeric score ──────────────────────────────────────────────
        NULL_TOKENS = {"NA", "NR", "-", "", "nan", "none"}
        for col in ["overall_score", "obj_score", "comp_score"]:
            if col in df.columns:
                def _num(v):
                    if v is None: return np.nan
                    if isinstance(v, (int, float)):
                        return float(v) if not np.isnan(float(v)) and float(v) > 0 else np.nan
                    sv = str(v).strip().lower()
                    if sv in NULL_TOKENS: return np.nan
                    try:
                        f = float(sv)
                        return f if f > 0 else np.nan
                    except Exception:
                        return np.nan
                df[col] = df[col].apply(_num)

        # ── clean text fields ────────────────────────────────────────────────
        for col in ["name", "dept", "position", "location", "pot_zone_raw", "perf_zone_raw"]:
            if col in df.columns:
                df[col] = df[col].apply(
                    lambda v: str(v).strip()
                    if v and str(v).strip().lower() not in NULL_TOKENS
                    else np.nan
                )

        # ── map location to readable names ───────────────────────────────────
        if "location" in df.columns:
            df["location"] = df["location"].map(
                lambda v: LOC_MAP.get(v, v) if isinstance(v, str) else v
            )

        # ── derived columns ──────────────────────────────────────────────────
        df["perf_tier"] = df["grid_zone"].apply(
            lambda v: int(v[0]) if isinstance(v, str) and len(v) == 2 else np.nan
        )
        df["pot_tier"] = df["grid_zone"].apply(
            lambda v: {"A": 1, "B": 2, "C": 3}.get(v[1], np.nan)
            if isinstance(v, str) and len(v) == 2 else np.nan
        )
        df["grid_label"] = df["grid_zone"].map(
            {k: m["label"] for k, m in GRID_META.items()}
        )
        df["year"] = year
        frames.append(df)

    out = pd.concat(frames, ignore_index=True)
    out["name"] = out["name"].astype(str).str.strip()
    out = out[out["name"].notna() & (out["name"] != "nan") & (out["name"] != "")]
    return out


# ── LOAD ──────────────────────────────────────────────────────────────────────
DATA_PATH = "9Grid_Final.xlsx"
try:
    df_all = load_all(DATA_PATH)
except FileNotFoundError:
    st.error(f"❌  Excel file `{DATA_PATH}` not found. Place it in the same folder as app.py.")
    st.stop()

# ── LOGO ──────────────────────────────────────────────────────────────────────
def img_b64(path):
    try:
        with open(path, "rb") as f:
            return base64.b64encode(f.read()).decode()
    except Exception:
        return None

logo_b64 = img_b64("logo.png")

# ── SIDEBAR ───────────────────────────────────────────────────────────────────
with st.sidebar:
    if logo_b64:
        st.markdown(
            f'<div style="text-align:center;padding:18px 0 10px">'
            f'<img src="data:image/png;base64,{logo_b64}" '
            f'style="max-width:170px;width:100%;border-radius:6px"></div>',
            unsafe_allow_html=True,
        )
    st.markdown(
        f'<div style="text-align:center;font-size:11px;color:#64748B;'
        f'margin-bottom:12px">Talent 9-Grid Dashboard</div>',
        unsafe_allow_html=True,
    )
    st.markdown("---")

    # ── Year ──────────────────────────────────────────────────────────────────
    sel_year = st.selectbox("📅  Year", [2025, 2024, 2023], index=0)

    # ── Department ────────────────────────────────────────────────────────────
    all_depts = ["All"] + sorted(df_all["dept"].dropna().unique())
    sel_dept  = st.selectbox("🏢  Department", all_depts, index=0)

    # ── Location ──────────────────────────────────────────────────────────────
    all_locs = ["All"] + sorted(df_all["location"].dropna().unique())
    sel_loc  = st.selectbox("📍  Location", all_locs, index=0)

    # ── Employee Name ─────────────────────────────────────────────────────────
    yr_names  = sorted(df_all[df_all["year"] == sel_year]["name"].dropna().unique())
    sel_name  = st.selectbox("👤  Employee Name", ["All"] + yr_names, index=0)

    st.markdown("---")
    st.caption("Source: 9Grid_Final.xlsx · Evalutaion sheets 23 / 24 / 25")


# ── FILTER HELPERS ────────────────────────────────────────────────────────────
def base_filter(df, year):
    """Year + dept + location (no name, no category)."""
    m = df["year"] == year
    if sel_dept != "All":
        m &= df["dept"] == sel_dept
    if sel_loc != "All":
        m &= df["location"] == sel_loc
    return df[m].copy()

def full_filter(df, year):
    """All slicers applied."""
    d = base_filter(df, year)
    if sel_name != "All":
        d = d[d["name"] == sel_name]
    return d

# filtered dataframes for selected year
df_base = base_filter(df_all, sel_year)              # dept+loc only
df      = full_filter(df_all, sel_year)              # all filters
gdf     = df[df["grid_zone"].notna()].copy()         # has valid grid zone
gdf_base= df_base[df_base["grid_zone"].notna()].copy()


# ── PAGE HEADER ───────────────────────────────────────────────────────────────
logo_img = (
    f'<img src="data:image/png;base64,{logo_b64}" style="height:46px;border-radius:6px">'
    if logo_b64 else ""
)
yr_badge = (
    f'<span style="background:{C_PURPLE};color:#fff;border-radius:999px;'
    f'padding:3px 14px;font-size:12px;font-weight:700">{sel_year}</span>'
)
loc_txt  = sel_loc  if sel_loc  != "All" else "All Locations"
dept_txt = sel_dept if sel_dept != "All" else "All Departments"
name_txt = f" · {sel_name}" if sel_name != "All" else ""

st.markdown(
    f'<div style="background:linear-gradient(135deg,{C_SIDEBAR} 0%,#2D2B55 100%);'
    f'border-radius:16px;padding:22px 28px;margin-bottom:20px;'
    f'display:flex;align-items:center;gap:18px">'
    f'<div>{logo_img}</div>'
    f'<div>'
    f'<div style="font-size:20px;font-weight:800;color:#F1F5F9;margin-bottom:4px">'
    f'Talent 9-Grid Dashboard</div>'
    f'<div style="font-size:12px;color:#94A3B8">'
    f'{yr_badge} &nbsp;{dept_txt} · {loc_txt}{name_txt}'
    f' &nbsp;·&nbsp; <b style="color:#30BFA6">{len(gdf)}</b> evaluated employees</div>'
    f'</div></div>',
    unsafe_allow_html=True,
)


# ── KPI CARDS ─────────────────────────────────────────────────────────────────
def kpi(label, val, sub="", cls="", delta="", ddir=""):
    d = (f'<div class="kpi-dlt {ddir}">{delta}</div>' if delta else "")
    return (
        f'<div class="kpi-wrap {cls}">'
        f'<div class="kpi-lbl">{label}</div>'
        f'<div class="kpi-val">{val}</div>'
        f'<div class="kpi-sub">{sub}</div>'
        f'{d}</div>'
    )

total   = len(gdf)
stars   = int((gdf["grid_zone"] == "3C").sum())
hiperf  = int((gdf["grid_zone"] == "3B").sum())
risk    = int((gdf["grid_zone"] == "1A").sum())
hipot   = int((gdf["pot_tier"]  == 3).sum())
avg_sc  = gdf["overall_score"].dropna().mean()
avg_pot = gdf["pot_tier"].dropna().mean()
hip_pct = hipot / total * 100 if total else 0

# YoY delta (stars)
prev_yr  = sel_year - 1
_pm2     = df_all["year"] == prev_yr
if sel_dept != "All":
    _pm2 &= df_all["dept"] == sel_dept
if sel_loc != "All":
    _pm2 &= df_all["location"] == sel_loc
prev_base = df_all[_pm2]
prev_stars = int((prev_base["grid_zone"] == "3C").sum()) if len(prev_base) else None
if prev_stars is not None:
    dn = stars - prev_stars
    s_delta = f"{'▲' if dn >= 0 else '▼'} {abs(dn)} vs {prev_yr}"
    s_ddir  = "up" if dn >= 0 else "dn"
else:
    s_delta, s_ddir = "Baseline year", ""

k1, k2, k3, k4, k5, k6, k7 = st.columns(7)
for col, html in zip(
    [k1, k2, k3, k4, k5, k6, k7],
    [
        kpi("Total Employees",   total,
            f"{sel_year} evaluation cycle", "gray"),
        kpi("⭐ Stars (3C)",     stars,
            f"{round(stars/total*100,1) if total else 0}% of workforce", "",
            s_delta, s_ddir),
        kpi("🟢 High Performers",hiperf,
            "Grid 3B", "green"),
        kpi("🔴 At Risk (1A)",   risk,
            "Immediate action", "red"),
        kpi("🎯 High Potential %",
            f"{hip_pct:.1f}%",
            f"{hipot} employees", "violet"),
        kpi("📊 Avg Performance",
            f"{avg_sc:.2f}" if not np.isnan(avg_sc) else "—",
            "Score out of 5.0", "green"),
        kpi("⚡ Avg Potential",
            f"{avg_pot:.2f}" if not np.isnan(avg_pot) else "—",
            "Tier 1–3", "amber"),
    ],
):
    with col:
        st.markdown(html, unsafe_allow_html=True)


# ── TABS ──────────────────────────────────────────────────────────────────────
tab1, tab2, tab3, tab4 = st.tabs([
    "🔲  9-Box Matrix",
    "🏆  Top Performers",
    "📈  Year Comparison",
    "📋  Employee Register",
])


# ══════════════════════════════════════════════════════════════════════════════
# TAB 1  —  9-BOX MATRIX  (card-based layout, no scatter)
# ══════════════════════════════════════════════════════════════════════════════
with tab1:

    # ── Inject card-grid CSS (scoped, injected once per render) ──────────────
    st.markdown("""
    <style>
    /* ── axis label strips ── */
    .grid-axis-top{display:grid;grid-template-columns:repeat(3,1fr);gap:10px;
        margin-bottom:6px;padding:0 2px;}
    .grid-axis-left{display:flex;flex-direction:column;justify-content:space-around;
        align-items:center;padding:2px 0;writing-mode:vertical-lr;
        transform:rotate(180deg);height:100%;}
    .ax-lbl{font-size:11px;font-weight:700;color:#6B7280;text-align:center;
        text-transform:uppercase;letter-spacing:.06em;}
    /* ── 3×3 outer wrapper ── */
    .grid9-wrap{display:grid;grid-template-columns:auto 1fr;gap:0;width:100%;}
    .grid9-rows{display:flex;flex-direction:column;gap:10px;flex:1;}
    .grid9-row{display:grid;grid-template-columns:repeat(3,1fr);gap:10px;}
    /* ── single box card ── */
    .g9box{border-radius:14px;padding:14px 14px 10px;background:#fff;
        border:1.5px solid #E5E7EB;
        box-shadow:0 2px 10px rgba(0,0,0,.06);
        display:flex;flex-direction:column;min-height:170px;
        transition:box-shadow .2s;}
    .g9box:hover{box-shadow:0 6px 20px rgba(0,0,0,.11);}
    /* header row inside box */
    .g9box-head{display:flex;justify-content:space-between;
        align-items:flex-start;margin-bottom:6px;}
    .g9box-code{font-size:10px;font-weight:800;letter-spacing:.08em;
        opacity:.65;line-height:1;}
    .g9box-title{font-size:12px;font-weight:800;line-height:1.2;
        margin-top:2px;color:#1E1E2E;}
    .g9box-cnt{font-size:26px;font-weight:900;line-height:1;
        text-align:right;}
    .g9box-cnt-lbl{font-size:9px;font-weight:600;text-align:right;
        opacity:.6;text-transform:uppercase;letter-spacing:.05em;}
    /* divider */
    .g9box-div{height:1px;background:rgba(0,0,0,.07);margin:8px 0 6px;}
    /* scrollable name list */
    .g9box-names{flex:1;overflow-y:auto;max-height:120px;
        scrollbar-width:thin;scrollbar-color:rgba(0,0,0,.15) transparent;}
    .g9box-names::-webkit-scrollbar{width:3px;}
    .g9box-names::-webkit-scrollbar-thumb{background:rgba(0,0,0,.15);border-radius:3px;}
    .g9name{font-size:10.5px;line-height:1.65;white-space:nowrap;
        overflow:hidden;text-overflow:ellipsis;padding:0 2px;}
    /* empty state */
    .g9-empty{font-size:10px;color:#9CA3AF;font-style:italic;text-align:center;
        margin-top:20px;}
    /* performance axis bottom */
    .grid-axis-bot{display:grid;grid-template-columns:auto repeat(3,1fr);
        gap:10px;margin-top:8px;}
    .perf-lbl{font-size:11px;font-weight:700;color:#6B7280;text-align:center;
        text-transform:uppercase;letter-spacing:.06em;}
    .grid-xaxis-title{font-size:11px;font-weight:700;color:#6B7280;text-align:center;
        margin-top:4px;letter-spacing:.05em;}
    .grid-yaxis-title{font-size:11px;font-weight:700;color:#6B7280;text-align:center;
        letter-spacing:.05em;margin-bottom:8px;}
    </style>
    """, unsafe_allow_html=True)

    # ── Section header ────────────────────────────────────────────────────────
    st.markdown(
        f'<div class="sec-head">🔲 9-Box Talent Matrix'
        f'<span class="pill">{sel_year}</span>'
        f'<span style="font-size:11px;font-weight:400;color:#6B7280;margin-left:auto">'
        f'{len(gdf)} employees evaluated</span></div>',
        unsafe_allow_html=True,
    )

    # ── Box color config (border + accent + bg) ───────────────────────────────
    # Layout: rows top→bottom = High Pot (3) → Med Pot (2) → Low Pot (1)
    #         cols left→right = Poor Perf(1) → Avg Perf(2) → High Perf(3)
    # Grid rendering order (row by row, top to bottom):
    #   Row 0 (High Pot):   1C, 2C, 3C
    #   Row 1 (Med Pot):    1B, 2B, 3B
    #   Row 2 (Low Pot):    1A, 2A, 3A
    RENDER_ORDER = [
        ["1C", "2C", "3C"],
        ["1B", "2B", "3B"],
        ["1A", "2A", "3A"],
    ]

    BOX_STYLE = {
        "3C": {"accent": "#666EFF", "bg": "#F0F1FF", "border": "#C7CAFF"},
        "3B": {"accent": "#30BFA6", "bg": "#EDFAF7", "border": "#A8EDE3"},
        "3A": {"accent": "#34D399", "bg": "#ECFDF5", "border": "#A7F3D0"},
        "2C": {"accent": "#818CF8", "bg": "#F5F3FF", "border": "#C4B5FD"},
        "2B": {"accent": "#6B7280", "bg": "#F9FAFB", "border": "#D1D5DB"},
        "2A": {"accent": "#F59E0B", "bg": "#FFFBEB", "border": "#FDE68A"},
        "1C": {"accent": "#A78BFA", "bg": "#FAF5FF", "border": "#DDD6FE"},
        "1B": {"accent": "#F97316", "bg": "#FFF7ED", "border": "#FED7AA"},
        "1A": {"accent": "#EF4444", "bg": "#FEF2F2", "border": "#FECACA"},
    }

    # pre-compute per-box employee lists
    box_data = {}
    for code in GRID_META:
        sub = gdf[gdf["grid_zone"] == code].copy()
        box_data[code] = {
            "count": len(sub),
            "names": sorted(sub["name"].tolist()),
        }

    # ── Build HTML for the 3×3 grid ──────────────────────────────────────────
    def box_html(code: str) -> str:
        meta  = GRID_META[code]
        style = BOX_STYLE[code]
        data  = box_data[code]
        cnt   = data["count"]
        names = data["names"]

        # name list HTML
        if names:
            name_items = "".join(
                f'<div class="g9name" '
                f'style="color:{style["accent"]}" '
                f'title="{n}">• {n}</div>'
                for n in names
            )
            names_html = f'<div class="g9box-names">{name_items}</div>'
        else:
            names_html = '<div class="g9-empty">No employees</div>'

        return (
            f'<div class="g9box" '
            f'style="background:{style["bg"]};border-color:{style["border"]}">'

            # header: code+title on left, count on right
            f'<div class="g9box-head">'
            f'  <div>'
            f'    <div class="g9box-code" style="color:{style["accent"]}">{code}</div>'
            f'    <div class="g9box-title">{meta["label"]}</div>'
            f'  </div>'
            f'  <div style="text-align:right">'
            f'    <div class="g9box-cnt" style="color:{style["accent"]}">{cnt}</div>'
            f'    <div class="g9box-cnt-lbl" style="color:{style["accent"]}">people</div>'
            f'  </div>'
            f'</div>'

            # divider
            f'<div class="g9box-div" style="background:{style["border"]}"></div>'

            # scrollable name list
            f'{names_html}'
            f'</div>'
        )

    # ── Assemble rows ─────────────────────────────────────────────────────────
    pot_labels   = ["High Potential", "Medium Potential", "Low Potential"]
    perf_labels  = ["Poor Performance", "Average Performance", "High Performance"]

    # top axis: performance labels
    top_axis = (
        '<div style="display:grid;grid-template-columns:42px repeat(3,1fr);'
        'gap:10px;margin-bottom:4px;padding-right:2px">'
        '<div></div>'  # spacer for Y-axis column
        + "".join(
            f'<div class="perf-lbl">{lbl}</div>' for lbl in perf_labels
        )
        + '</div>'
    )

    # rows with Y-axis label
    rows_html = ""
    for ri, (row_codes, pot_lbl) in enumerate(zip(RENDER_ORDER, pot_labels)):
        row_boxes = "".join(box_html(code) for code in row_codes)
        rows_html += (
            f'<div style="display:grid;grid-template-columns:42px 1fr;'
            f'gap:10px;margin-bottom:10px;align-items:stretch">'

            # Y-axis label cell
            f'<div style="display:flex;align-items:center;justify-content:center">'
            f'  <div style="writing-mode:vertical-lr;transform:rotate(180deg);'
            f'  font-size:10px;font-weight:700;color:#6B7280;'
            f'  text-transform:uppercase;letter-spacing:.06em;white-space:nowrap">'
            f'  {pot_lbl}</div>'
            f'</div>'

            # 3 boxes
            f'<div style="display:grid;grid-template-columns:repeat(3,1fr);gap:10px">'
            f'{row_boxes}'
            f'</div>'

            f'</div>'
        )

    # bottom X-axis title
    x_title = (
        '<div style="display:grid;grid-template-columns:42px 1fr;gap:10px;margin-top:2px">'
        '<div></div>'
        '<div class="grid-xaxis-title">← Performance →</div>'
        '</div>'
    )
    # left Y-axis title
    y_title = (
        '<div class="grid-yaxis-title">← Potential →</div>'
    )

    full_grid_html = (
        f'<div style="width:100%">'
        f'{y_title}'
        f'{top_axis}'
        f'{rows_html}'
        f'{x_title}'
        f'</div>'
    )

    st.markdown(full_grid_html, unsafe_allow_html=True)

    # ── Legend strip below grid ───────────────────────────────────────────────
    st.markdown("<br>", unsafe_allow_html=True)
    legend_items = "".join(
        f'<span style="display:inline-flex;align-items:center;gap:5px;'
        f'margin-right:14px;font-size:11px;color:#374151">'
        f'<span style="width:10px;height:10px;border-radius:3px;'
        f'background:{BOX_STYLE[k]["accent"]};display:inline-block"></span>'
        f'{k} {GRID_META[k]["label"]}</span>'
        for k in CAT_ORDER
    )
    st.markdown(
        f'<div style="display:flex;flex-wrap:wrap;gap:4px;padding:10px 12px;'
        f'background:#F9FAFB;border-radius:10px;border:1px solid #E5E7EB">'
        f'{legend_items}</div>',
        unsafe_allow_html=True,
    )


# TAB 2  —  TOP PERFORMERS
# ══════════════════════════════════════════════════════════════════════════════
with tab2:
    st.markdown(
        f'<div class="sec-head">🏆 Top Performers'
        f'<span class="pill">{sel_year}</span></div>',
        unsafe_allow_html=True,
    )

    # Rank: 3C first → 3B → then by score descending
    grade_order = {code: i for i, code in enumerate(CAT_ORDER)}
    top_df = gdf.copy()
    top_df["rank_key"] = top_df["grid_zone"].map(grade_order).fillna(99)
    top_df = top_df.sort_values(
        ["rank_key", "overall_score"], ascending=[True, False]
    ).reset_index(drop=True)

    # Summary stat row
    s1, s2, s3 = st.columns(3)
    with s1:
        n_top = int((top_df["grid_zone"].isin(["3C", "3B"])).sum())
        st.markdown(
            kpi("Top Talent (3C + 3B)", n_top,
                f"{round(n_top/total*100,1) if total else 0}% of workforce"),
            unsafe_allow_html=True,
        )
    with s2:
        best_dept = (
            top_df[top_df["grid_zone"] == "3C"]["dept"].value_counts().idxmax()
            if not top_df[top_df["grid_zone"] == "3C"].empty else "—"
        )
        st.markdown(
            kpi("Top Dept for Stars", best_dept, "Most 3C employees", "green"),
            unsafe_allow_html=True,
        )
    with s3:
        best_loc = (
            top_df[top_df["grid_zone"] == "3C"]["location"].value_counts().idxmax()
            if not top_df[top_df["grid_zone"] == "3C"].empty else "—"
        )
        st.markdown(
            kpi("Top Location for Stars", best_loc, "Most 3C employees", "violet"),
            unsafe_allow_html=True,
        )

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Cards view for top 15 ──────────────────────────────────────────────
    show_n  = min(15, len(top_df))
    n_cols  = 3
    rows_   = [top_df.iloc[i:i + n_cols] for i in range(0, show_n, n_cols)]

    for row_chunk in rows_:
        cols = st.columns(n_cols)
        for col, (_, emp) in zip(cols, row_chunk.iterrows()):
            code  = emp["grid_zone"]
            bc    = BOX_COLORS.get(code, {"dot": C_PURPLE})
            score = f"{emp['overall_score']:.2f}" if not np.isnan(emp.get("overall_score", np.nan)) else "—"
            pot_  = (
                "High"   if emp["pot_tier"] == 3
                else "Medium" if emp["pot_tier"] == 2
                else "Low"    if emp["pot_tier"] == 1
                else "—"
            )
            cls_  = ("green" if code in ("3C", "3B")
                     else "amber" if code.startswith("2")
                     else "")
            with col:
                st.markdown(
                    f'<div class="tp-card {cls_}">'
                    f'<div class="tp-name">{emp["name"]}</div>'
                    f'<div class="tp-role">'
                    f'{emp.get("position","—") or "—"}</div>'
                    f'<div class="tp-meta">'
                    f'🏢 {emp.get("dept","—") or "—"}<br>'
                    f'📍 {emp.get("location","—") or "—"}<br>'
                    f'📊 Score: <b>{score}</b> &nbsp;|&nbsp; Potential: <b>{pot_}</b>'
                    f'</div>'
                    f'<span class="tp-badge" '
                    f'style="background:{bc["bg"]};color:{bc["dot"]};'
                    f'border:1px solid {bc["border"]}">'
                    f'{code} · {GRID_META[code]["label"]}'
                    f'</span></div>',
                    unsafe_allow_html=True,
                )

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Full sortable table ────────────────────────────────────────────────
    st.markdown(
        f'<div class="sec-head" style="margin-top:0">Full Ranked Table</div>',
        unsafe_allow_html=True,
    )
    tbl = top_df[[
        "name", "position", "dept", "location",
        "grid_zone", "grid_label", "overall_score", "pot_tier",
    ]].copy()
    tbl.columns = [
        "Employee Name", "Position", "Department", "Location",
        "Grid Code", "Category", "Performance Score", "Potential (1-3)",
    ]
    tbl["Performance Score"] = tbl["Performance Score"].round(2)
    tbl = tbl.reset_index(drop=True)
    tbl.index += 1
    st.dataframe(tbl, use_container_width=True, height=400)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 3  —  YEAR COMPARISON  |  Executive Talent Analytics
# ══════════════════════════════════════════════════════════════════════════════
with tab3:
    st.markdown("""
<style>
.yc-hero{background:linear-gradient(135deg,#0B1025 0%,#17133F 55%,#0F172A 100%);border-radius:18px;padding:22px 26px;margin:2px 0 18px;color:#fff;box-shadow:0 10px 30px rgba(15,23,42,.18);position:relative;overflow:hidden;}
.yc-hero:before{content:'';position:absolute;top:-80px;right:-70px;width:220px;height:220px;border-radius:999px;background:radial-gradient(circle,rgba(102,110,255,.36),rgba(48,191,166,.10),transparent 68%);}
.yc-hero-kicker{color:#9CA3FF;font-size:11px;font-weight:800;text-transform:uppercase;letter-spacing:.13em;margin-bottom:8px;}
.yc-hero-title{font-size:25px;font-weight:900;letter-spacing:-.04em;margin-bottom:5px;color:#fff;}
.yc-hero-sub{color:#CBD5E1;font-size:12px;line-height:1.6;max-width:760px;}
.yc-card{background:#FFFFFF;border:1px solid #E8ECF4;border-radius:16px;padding:16px 17px;box-shadow:0 2px 8px rgba(15,23,42,.05),0 10px 25px rgba(15,23,42,.04);position:relative;overflow:hidden;min-height:104px;}
.yc-card:before{content:'';position:absolute;left:0;top:0;width:100%;height:3px;background:linear-gradient(90deg,#666EFF,#30BFA6);}
.yc-label{font-size:10px;font-weight:800;color:#8A94A6;text-transform:uppercase;letter-spacing:.08em;margin-bottom:7px;}
.yc-value{font-size:30px;line-height:1;font-weight:950;color:#101828;letter-spacing:-.05em;}
.yc-sub{font-size:11px;color:#667085;margin-top:7px;line-height:1.45;}
.yc-delta{display:inline-flex;align-items:center;gap:4px;font-size:11px;font-weight:800;padding:4px 9px;border-radius:999px;margin-top:8px;}
.yc-up{color:#027A48;background:#ECFDF3;border:1px solid #ABEFC6;}.yc-down{color:#B42318;background:#FEF3F2;border:1px solid #FECDCA;}.yc-neutral{color:#475467;background:#F2F4F7;border:1px solid #EAECF0;}
.yc-panel{background:#FFFFFF;border:1px solid #E8ECF4;border-radius:18px;padding:18px 20px;box-shadow:0 2px 8px rgba(15,23,42,.05),0 12px 28px rgba(15,23,42,.04);margin-bottom:16px;}
.yc-panel-title{display:flex;align-items:center;justify-content:space-between;gap:12px;margin-bottom:14px;}
.yc-panel-title h3{margin:0;font-size:15px;font-weight:900;letter-spacing:-.02em;color:#101828;}
.yc-panel-title span{font-size:10px;font-weight:800;letter-spacing:.08em;color:#667085;text-transform:uppercase;background:#F4F5FF;border:1px solid #DDE0FF;padding:4px 9px;border-radius:999px;}
.move-row{display:grid;grid-template-columns:1fr 36px 1fr 70px;gap:10px;align-items:center;background:#F8FAFC;border:1px solid #EEF2F6;border-radius:14px;padding:10px 12px;margin-bottom:9px;}
.move-from,.move-to{font-size:12px;font-weight:800;color:#263238;}.move-caption{font-size:10px;color:#98A2B3;font-weight:700;text-transform:uppercase;letter-spacing:.08em;margin-bottom:2px;}
.move-arrow{width:34px;height:34px;border-radius:999px;display:flex;align-items:center;justify-content:center;background:linear-gradient(135deg,#666EFF,#30BFA6);color:#fff;font-weight:900;box-shadow:0 6px 16px rgba(102,110,255,.26);}
.move-count{text-align:center;font-size:20px;font-weight:950;color:#666EFF;line-height:1;}.move-count small{display:block;color:#98A2B3;font-size:9px;font-weight:800;margin-top:3px;text-transform:uppercase;letter-spacing:.07em;}
.trend-row{display:grid;grid-template-columns:145px 1fr;gap:12px;align-items:start;margin-bottom:13px;}.trend-name{font-size:12px;color:#101828;font-weight:850;padding-top:3px;}
.bar-line{display:grid;grid-template-columns:42px 1fr 42px;gap:8px;align-items:center;margin-bottom:6px;}.bar-year{font-size:10px;color:#667085;font-weight:800;}.bar-track{height:10px;background:#EEF2F6;border-radius:999px;overflow:hidden;}.bar-fill{height:10px;border-radius:999px;background:linear-gradient(90deg,#666EFF,#30BFA6);}.bar-val{font-size:10px;color:#344054;font-weight:800;text-align:right;}
.heat-table{width:100%;border-collapse:separate;border-spacing:0 8px;}.heat-table th{font-size:10px;color:#98A2B3;text-transform:uppercase;letter-spacing:.08em;text-align:center;padding:2px 8px;}.heat-table th:first-child{text-align:left;}.heat-table td{font-size:12px;font-weight:850;color:#101828;padding:10px 10px;background:#F8FAFC;border-top:1px solid #EEF2F6;border-bottom:1px solid #EEF2F6;text-align:center;}.heat-table td:first-child{text-align:left;border-left:1px solid #EEF2F6;border-radius:12px 0 0 12px;background:#FFFFFF;}.heat-table td:last-child{border-right:1px solid #EEF2F6;border-radius:0 12px 12px 0;}
.spotlight-card{background:linear-gradient(135deg,#FFFFFF,#F7F8FF);border:1px solid #E4E7FF;border-radius:15px;padding:12px 14px;margin-bottom:10px;display:flex;align-items:center;gap:12px;}.rank-badge{width:35px;height:35px;border-radius:12px;display:flex;align-items:center;justify-content:center;background:linear-gradient(135deg,#666EFF,#30BFA6);color:#fff;font-weight:950;box-shadow:0 6px 16px rgba(102,110,255,.23);flex-shrink:0;}.spot-name{font-size:12px;color:#101828;font-weight:900;line-height:1.25;}.spot-meta{font-size:10px;color:#667085;margin-top:3px;line-height:1.45;}.spot-score{margin-left:auto;text-align:right;font-size:18px;font-weight:950;color:#666EFF;}.spot-score small{display:block;font-size:9px;color:#98A2B3;font-weight:800;text-transform:uppercase;letter-spacing:.06em;}
.insight-box{background:linear-gradient(135deg,#101828 0%,#1D1745 100%);color:#fff;border-radius:18px;padding:18px 20px;box-shadow:0 12px 30px rgba(16,24,40,.17);height:100%;}.insight-box h3{margin:0 0 12px 0;font-size:15px;font-weight:900;color:#fff;}.insight-box ul{margin:0;padding-left:18px;color:#D0D5DD;font-size:12px;line-height:1.8;}.insight-box b{color:#FFFFFF;}.mini-note{font-size:10px;color:#98A2B3;margin-top:8px;}
</style>
""", unsafe_allow_html=True)

    st.markdown(f'<div class="yc-hero"><div class="yc-hero-kicker">Executive Talent Analytics</div><div class="yc-hero-title">Year-over-Year Talent Health</div><div class="yc-hero-sub">A simplified executive view of how talent quality, risk exposure, performance, and grid movement changed across 2023, 2024, and 2025.</div></div>', unsafe_allow_html=True)

    YEARS = [2023, 2024, 2025]
    trend_rows, dist_rows = [], []
    year_frames = {}
    for yr in YEARS:
        ydf = base_filter(df_all, yr)
        ygdf = ydf[ydf["grid_zone"].notna()].copy()
        year_frames[yr] = ygdf
        trend_rows.append({
            "year": yr,
            "total": len(ygdf),
            "Stars": int((ygdf["grid_zone"] == "3C").sum()),
            "High Performers": int((ygdf["grid_zone"] == "3B").sum()),
            "Growth Employees": int((ygdf["grid_zone"] == "2C").sum()),
            "Core Players": int((ygdf["grid_zone"] == "2B").sum()),
            "At Risk": int((ygdf["grid_zone"] == "1A").sum()),
            "High Potential": int((ygdf["pot_tier"] == 3).sum()),
            "avg_score": ygdf["overall_score"].dropna().mean(),
        })
        for code in CAT_ORDER:
            dist_rows.append({"year": yr, "category": GRID_META[code]["label"], "code": code, "count": int((ygdf["grid_zone"] == code).sum())})

    trend = pd.DataFrame(trend_rows)
    dist = pd.DataFrame(dist_rows)

    def _safe(v, default=0):
        try:
            return default if pd.isna(v) else v
        except Exception:
            return v

    def _yr_val(year, col):
        row = trend[trend["year"] == year]
        return 0 if row.empty else _safe(row.iloc[0][col], 0)

    def _delta(curr, prev, lower_is_better=False, suffix=""):
        curr, prev = _safe(curr, 0), _safe(prev, 0)
        diff = curr - prev
        good = diff >= 0
        if lower_is_better:
            good = diff <= 0
        cls = "yc-up" if good else "yc-down"
        arrow = "▲" if diff >= 0 else "▼"
        if abs(diff) < 0.0001:
            cls, arrow = "yc-neutral", "→"
        value = f"{abs(diff):.2f}" if isinstance(diff, float) and not float(diff).is_integer() else f"{abs(int(diff))}"
        return f'<div class="yc-delta {cls}">{arrow} {value}{suffix} vs 2023</div>'

    stars_now = int(_yr_val(2025, "Stars")); stars_base = int(_yr_val(2023, "Stars"))
    hp_now = int(_yr_val(2025, "High Potential")); hp_base = int(_yr_val(2023, "High Potential"))
    risk_now = int(_yr_val(2025, "At Risk")); risk_base = int(_yr_val(2023, "At Risk"))
    avg_now = float(_safe(_yr_val(2025, "avg_score"), 0)); avg_base = float(_safe(_yr_val(2023, "avg_score"), 0))

    c1, c2, c3, c4 = st.columns(4)
    cards = [
        ("⭐ Stars", stars_now, "Top 9-Grid box in 2025", _delta(stars_now, stars_base)),
        ("🚀 High Potential", hp_now, "Employees with high potential", _delta(hp_now, hp_base)),
        ("⚠️ Risk Exposure", risk_now, "At Risk employees in 2025", _delta(risk_now, risk_base, lower_is_better=True)),
        ("📈 Avg Performance", f"{avg_now:.2f}" if avg_now else "—", "Average score in 2025", _delta(avg_now, avg_base)),
    ]
    for col, (lbl, val, sub, dlt) in zip([c1, c2, c3, c4], cards):
        with col:
            st.markdown(f'<div class="yc-card"><div class="yc-label">{lbl}</div><div class="yc-value">{val}</div><div class="yc-sub">{sub}</div>{dlt}</div>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    _pm = pd.Series(True, index=df_all.index)
    if sel_dept != "All": _pm &= df_all["dept"] == sel_dept
    if sel_loc != "All": _pm &= df_all["location"] == sel_loc
    pivot = df_all[_pm].pivot_table(index=["name", "dept", "location"], columns="year", values="grid_zone", aggfunc="first").reset_index()
    pivot.columns.name = None
    for yr in YEARS:
        if yr not in pivot.columns: pivot[yr] = np.nan

    def _grid_score(g):
        if pd.isna(g): return np.nan
        return GRID_META.get(g, {}).get("perf", 0) + GRID_META.get(g, {}).get("pot", 0)

    move_df = pivot[pivot[2024].notna() & pivot[2025].notna()].copy()
    move_df["from_label"] = move_df[2024].map(lambda x: GRID_META.get(x, {}).get("label", str(x)))
    move_df["to_label"] = move_df[2025].map(lambda x: GRID_META.get(x, {}).get("label", str(x)))
    move_df["from_score"] = move_df[2024].map(_grid_score)
    move_df["to_score"] = move_df[2025].map(_grid_score)
    move_df["movement_type"] = np.where(move_df["to_score"] > move_df["from_score"], "Improved", np.where(move_df["to_score"] < move_df["from_score"], "Declined", "Stable"))
    transitions = move_df.groupby(["from_label", "to_label", "movement_type"]).size().reset_index(name="count").sort_values("count", ascending=False)
    top_transitions = transitions[transitions["from_label"] != transitions["to_label"]].head(5)
    if top_transitions.empty: top_transitions = transitions.head(5)

    left_col, right_col = st.columns([1.25, 1], gap="medium")
    with left_col:
        move_rows_html = ""
        if top_transitions.empty:
            move_rows_html = '<div class="mini-note">No movement data available for the selected filters.</div>'
        else:
            for _, r in top_transitions.iterrows():
                move_rows_html += f'<div class="move-row"><div><div class="move-caption">From</div><div class="move-from">{r["from_label"]}</div></div><div class="move-arrow">→</div><div><div class="move-caption">To</div><div class="move-to">{r["to_label"]}</div></div><div class="move-count">{int(r["count"])}<small>people</small></div></div>'
        st.markdown(f'<div class="yc-panel"><div class="yc-panel-title"><h3>Talent Movement</h3><span>2024 → 2025</span></div>{move_rows_html}</div>', unsafe_allow_html=True)

    with right_col:
        improved = int((move_df["movement_type"] == "Improved").sum()) if not move_df.empty else 0
        declined = int((move_df["movement_type"] == "Declined").sum()) if not move_df.empty else 0
        stable = int((move_df["movement_type"] == "Stable").sum()) if not move_df.empty else 0
        total_move = len(move_df) if len(move_df) else 1
        improved_pct, stable_pct, declined_pct = improved/total_move*100, stable/total_move*100, declined/total_move*100
        movement_summary_html = f'<div class="trend-row"><div class="trend-name">Improved</div><div><div class="bar-line"><div class="bar-year">%</div><div class="bar-track"><div class="bar-fill" style="width:{improved_pct:.1f}%"></div></div><div class="bar-val">{improved_pct:.0f}%</div></div></div></div><div class="trend-row"><div class="trend-name">Stable</div><div><div class="bar-line"><div class="bar-year">%</div><div class="bar-track"><div class="bar-fill" style="width:{stable_pct:.1f}%;background:linear-gradient(90deg,#98A2B3,#667085)"></div></div><div class="bar-val">{stable_pct:.0f}%</div></div></div></div><div class="trend-row"><div class="trend-name">Declined</div><div><div class="bar-line"><div class="bar-year">%</div><div class="bar-track"><div class="bar-fill" style="width:{declined_pct:.1f}%;background:linear-gradient(90deg,#F97316,#EF4444)"></div></div><div class="bar-val">{declined_pct:.0f}%</div></div></div></div>'
        st.markdown(f'<div class="yc-panel"><div class="yc-panel-title"><h3>Movement Health</h3><span>{len(move_df)} matched employees</span></div>{movement_summary_html}</div>', unsafe_allow_html=True)

    col_trend, col_spot = st.columns([1.2, 1], gap="medium")
    trend_metrics = [("Stars", "Stars"), ("High Performers", "High Performers"), ("Core Players", "Core Players"), ("At Risk", "At Risk")]
    max_val = max([int(_yr_val(y, mcol)) for _, mcol in trend_metrics for y in YEARS] + [1])
    trend_html = ""
    for display_name, metric_col in trend_metrics:
        trend_html += f'<div class="trend-row"><div class="trend-name">{display_name}</div><div>'
        for yr in YEARS:
            val = int(_yr_val(yr, metric_col)); width = max((val / max_val * 100), 3) if val else 0
            color = "linear-gradient(90deg,#F97316,#EF4444)" if metric_col == "At Risk" else "linear-gradient(90deg,#666EFF,#30BFA6)"
            trend_html += f'<div class="bar-line"><div class="bar-year">{yr}</div><div class="bar-track"><div class="bar-fill" style="width:{width:.1f}%;background:{color}"></div></div><div class="bar-val">{val}</div></div>'
        trend_html += '</div></div>'
    with col_trend:
        st.markdown(f'<div class="yc-panel"><div class="yc-panel-title"><h3>Talent Distribution Trend</h3><span>Headcount by year</span></div>{trend_html}</div>', unsafe_allow_html=True)

    with col_spot:
        latest_df = year_frames.get(2025, pd.DataFrame()).copy()
        grade_order = {code: i for i, code in enumerate(CAT_ORDER)}
        if not latest_df.empty:
            latest_df["rank_key"] = latest_df["grid_zone"].map(grade_order).fillna(99)
            latest_df = latest_df.sort_values(["rank_key", "overall_score"], ascending=[True, False]).head(3)
        spot_html = ""
        if latest_df.empty:
            spot_html = '<div class="mini-note">No top performer data available for 2025.</div>'
        else:
            for medal, (_, emp) in zip(["1", "2", "3"], latest_df.iterrows()):
                score = emp.get("overall_score", np.nan); score_txt = f"{score:.2f}" if not pd.isna(score) else "—"
                spot_html += f'<div class="spotlight-card"><div class="rank-badge">{medal}</div><div><div class="spot-name">{emp.get("name","—")}</div><div class="spot-meta">{emp.get("position","—") or "—"}<br>{emp.get("dept","—") or "—"} · {emp.get("grid_zone","—")}</div></div><div class="spot-score">{score_txt}<small>score</small></div></div>'
        st.markdown(f'<div class="yc-panel"><div class="yc-panel-title"><h3>Top Talent Spotlight</h3><span>2025</span></div>{spot_html}</div>', unsafe_allow_html=True)

    col_heat, col_ins = st.columns([1.35, 1], gap="medium")
    with col_heat:
        heat_rows = ""
        for code in CAT_ORDER:
            label = GRID_META[code]["label"]
            vals = [int(dist[(dist["code"] == code) & (dist["year"] == yr)]["count"].sum()) for yr in YEARS]
            row_max = max(vals + [1]); cells = ""
            for val in vals:
                alpha = 0.08 + (val / row_max) * 0.25 if row_max else 0.08
                color = "239,68,68" if code == "1A" else "102,110,255"
                cells += f'<td style="background:rgba({color},{alpha:.2f});color:#101828">{val}</td>'
            heat_rows += f'<tr><td>{code} · {label}</td>{cells}</tr>'
        heat_html = f'<table class="heat-table"><thead><tr><th>Category</th><th>2023</th><th>2024</th><th>2025</th></tr></thead><tbody>{heat_rows}</tbody></table>'
        st.markdown(f'<div class="yc-panel"><div class="yc-panel-title"><h3>Talent Heatmap</h3><span>9-Grid density</span></div>{heat_html}</div>', unsafe_allow_html=True)

    with col_ins:
        total_2025 = int(_yr_val(2025, "total")) or 1
        stars_pct = stars_now / total_2025 * 100
        risk_pct = risk_now / total_2025 * 100
        avg_change = avg_now - avg_base
        top_dept = "—"
        y2025 = year_frames.get(2025, pd.DataFrame()).copy()
        if not y2025.empty and "dept" in y2025.columns:
            star_depts = y2025[y2025["grid_zone"] == "3C"]["dept"].dropna()
            if not star_depts.empty: top_dept = star_depts.value_counts().idxmax()
        st.markdown(f'<div class="insight-box"><h3>AI-style Executive Insights</h3><ul><li><b>{stars_pct:.1f}%</b> of the 2025 evaluated population is classified as Stars.</li><li>Risk exposure is currently <b>{risk_pct:.1f}%</b> of the evaluated population.</li><li>Average performance changed by <b>{avg_change:+.2f}</b> points vs 2023.</li><li><b>{top_dept}</b> produced the highest number of Stars in 2025.</li><li><b>{improved}</b> employees improved their grid position from 2024 to 2025.</li></ul></div>', unsafe_allow_html=True)

    st.markdown('<div class="mini-note">Note: Year Comparison respects Department and Location filters. Employee Name filter is excluded to keep year-over-year analytics meaningful.</div>', unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 4  —  EMPLOYEE REGISTER
# ══════════════════════════════════════════════════════════════════════════════
with tab4:
    st.markdown(
        f'<div class="sec-head">📋 Full Employee Register</div>',
        unsafe_allow_html=True,
    )

    dept_mask = (df_all["dept"] == sel_dept) if sel_dept != "All" else pd.Series(True, index=df_all.index)
    loc_mask  = (df_all["location"] == sel_loc) if sel_loc != "All" else pd.Series(True, index=df_all.index)
    reg_base  = df_all[dept_mask & loc_mask]

    reg = (
        reg_base
        .pivot_table(
            index=["name", "dept", "location", "position"],
            columns="year",
            values=["grid_zone", "overall_score"],
            aggfunc="first",
        )
        .reset_index()
    )
    reg.columns = [
        f"{b}_{a}" if b else a for a, b in reg.columns
    ]
    reg = reg.rename(columns={
        "name": "Name", "dept": "Department",
        "location": "Location", "position": "Role",
        "grid_zone_2023": "Grid 2023", "grid_zone_2024": "Grid 2024",
        "grid_zone_2025": "Grid 2025",
        "overall_score_2023": "Score 2023", "overall_score_2024": "Score 2024",
        "overall_score_2025": "Score 2025",
    })
    for c in ["Score 2023", "Score 2024", "Score 2025"]:
        if c in reg.columns:
            reg[c] = pd.to_numeric(reg[c], errors="coerce").round(2)
    for c in ["Grid 2023", "Grid 2024", "Grid 2025"]:
        if c in reg.columns:
            reg[c] = reg[c].fillna("—")

    if sel_name != "All":
        reg = reg[reg["Name"] == sel_name]

    st.dataframe(reg.reset_index(drop=True), use_container_width=True, height=560)
    st.download_button(
        label="⬇️  Download CSV",
        data=reg.to_csv(index=False).encode("utf-8"),
        file_name=f"talent_register_{sel_year}.csv",
        mime="text/csv",
    )


# ── FOOTER ────────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown(
    f'<p style="text-align:center;color:#9CA3AF;font-size:10px">'
    f'MBC Media Solutions · Talent 9-Grid Dashboard · '
    f'Source: 9Grid_Final.xlsx (Evalutaion 23 / 24 / 25) · '
    f'Built with Streamlit & Plotly</p>',
    unsafe_allow_html=True,
)
