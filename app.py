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
/* tab selected state purple underline */
[data-testid="stTab"][aria-selected="true"] {{
    border-bottom:3px solid #666EFF !important; {C_PURPLE} !important;
    color:{C_PURPLE} !important;
}}
button[data-testid="stTab"][aria-selected="true"] {{
    color:{C_PURPLE} !important;
    border-bottom:3px solid {C_PURPLE} !important;
}}
button[data-testid="stTab"]:hover,
[data-testid="stTab"]:hover {{
    color:{C_PURPLE} !important;
    border-bottom-color:{C_PURPLE} !important;
}}
[data-testid="stTabs"] button:hover p,
[data-testid="stTabs"] button:hover span {{
    color:{C_PURPLE} !important;
}}
div[role="radiogroup"] label:hover,
div[role="radiogroup"] label:hover * {{
    color:{C_PURPLE} !important;
}}
div[role="radiogroup"] label div:first-child {{     border-color:#666EFF !important; }}  div[role="radiogroup"] label[aria-checked="true"] div:first-child, div[role="radiogroup"] label[data-checked="true"] div:first-child {{     background-color:#666EFF !important;     border-color:#666EFF !important; }}  div[role="radiogroup"] label[aria-checked="true"], div[role="radiogroup"] label[data-checked="true"] {{     color:#666EFF !important;     font-weight:700 !important; }} {     border-color:#666EFF !important; }  div[role="radiogroup"] label[aria-checked="true"] div:first-child, div[role="radiogroup"] label[data-checked="true"] div:first-child {     background-color:#666EFF !important;     border-color:#666EFF !important; }  div[role="radiogroup"] label[aria-checked="true"], div[role="radiogroup"] label[data-checked="true"] {     color:#666EFF !important;     font-weight:700 !important; } {     border-color:#666EFF !important; }  div[role="radiogroup"] label[aria-checked="true"] div:first-child, div[role="radiogroup"] label[data-checked="true"] div:first-child {     background-color:#666EFF !important;     border-color:#666EFF !important; }  div[role="radiogroup"] label[aria-checked="true"], div[role="radiogroup"] label[data-checked="true"] {     color:#666EFF !important;     font-weight:700 !important; } {{
    border-color:{C_PURPLE} !important;
}}
div[role="radiogroup"] label[aria-checked="true"] div:first-child,
div[role="radiogroup"] label[data-checked="true"] div:first-child {{
    background-color:{C_PURPLE} !important;
    border-color:{C_PURPLE} !important;
}}
div[role="radiogroup"] label[aria-checked="true"],
div[role="radiogroup"] label[data-checked="true"] {{
    color:{C_PURPLE} !important;
    font-weight:700 !important;
}}

/* ── hide default header ── */
header[data-testid="stHeader"] {{ display:none; }}

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
# TAB 3  —  YEAR COMPARISON
# ══════════════════════════════════════════════════════════════════════════════
with tab3:
    st.markdown(
        f'<div class="sec-head">📈 Year-over-Year Comparison  2023 · 2024 · 2025</div>',
        unsafe_allow_html=True,
    )

    # Build per-year stats (respecting dept/loc filter but not name/category)
    YEARS = [2023, 2024, 2025]
    trend_rows, dist_rows = [], []
    score_data = {}   # year → array of scores (for bell curve)

    for yr in YEARS:
        ydf = base_filter(df_all, yr)
        ygdf = ydf[ydf["grid_zone"].notna()].copy()
        trend_rows.append({
            "year": yr, "total": len(ygdf),
            "Stars (3C)":       int((ygdf["grid_zone"] == "3C").sum()),
            "High Performers":  int((ygdf["grid_zone"] == "3B").sum()),
            "Growth Employees": int((ygdf["grid_zone"] == "2C").sum()),
            "Core Players":     int((ygdf["grid_zone"] == "2B").sum()),
            "At Risk (1A)":     int((ygdf["grid_zone"] == "1A").sum()),
            "High Potential":   int((ygdf["pot_tier"] == 3).sum()),
            "avg_score":        ygdf["overall_score"].dropna().mean(),
        })
        for code in CAT_ORDER:
            dist_rows.append({
                "year": str(yr),
                "category": GRID_META[code]["label"],
                "code": code,
                "count": int((ygdf["grid_zone"] == code).sum()),
            })
        scores = ygdf["overall_score"].dropna().values
        score_data[yr] = scores

    trend = pd.DataFrame(trend_rows)
    dist  = pd.DataFrame(dist_rows)

    # ── Bell Curve — year selector + overlay ────────────────────────────────
    st.markdown(
        f'<div style="font-size:13px;font-weight:700;color:{C_DARK};margin-bottom:12px">' 
        "Performance Distribution — Bell Curve</div>",
        unsafe_allow_html=True,
    )

    # year selector for bell curve (independent of main sidebar year)
    bell_yr_options = ["All Years", "2025", "2024", "2023"]
    bell_sel = st.radio(
        "Show bell curve for:",
        bell_yr_options,
        index=0,
        horizontal=True,
        key="bell_yr_radio",
    )

    # colour + dash per year
    BELL_CFG = {
        2025: {"color": C_PURPLE,  "dash": "solid",  "fill": "rgba(102,110,255,0.09)"},
        2024: {"color": C_GREEN,   "dash": "dash",   "fill": "rgba(48,191,166,0.09)"},
        2023: {"color": "#A78BFA", "dash": "dot",    "fill": "rgba(167,139,250,0.09)"},
    }

    years_to_plot = [2025, 2024, 2023] if bell_sel == "All Years" else [int(bell_sel)]

    # ── reference distribution guideline (ideal bell, normalised to data) ────
    # Build a smooth normal curve as the "guideline" using combined score mean/std
    all_scores_combined = np.concatenate([score_data[y] for y in years_to_plot if len(score_data.get(y, [])) > 0])
    ref_mu  = all_scores_combined.mean() if len(all_scores_combined) else 3.5
    ref_std = all_scores_combined.std()  if len(all_scores_combined) else 0.5
    ref_x   = np.linspace(max(1.0, ref_mu - 3.5*ref_std), min(5.5, ref_mu + 3.5*ref_std), 200)
    ref_y   = (1 / (ref_std * np.sqrt(2 * np.pi))) * np.exp(-0.5 * ((ref_x - ref_mu) / ref_std) ** 2)

    fig_bell = go.Figure()

    # ── per-year bars (headcount per score bucket) + smoothed KDE line ────────
    PERF_ZONES = [
        ("Exceptional",        4.6, 5.0),
        ("Exceeds Expectations",3.6, 4.5),
        ("Meets Expectations",  3.0, 3.5),
        ("Below Expectations",  2.0, 2.9),
        ("Needs Improvement",   0.0, 1.9),
    ]

    bar_group_width = 0.25
    offsets = {2025: -bar_group_width, 2024: 0, 2023: bar_group_width}
    zone_centers = [1, 2, 3, 4, 5]   # numeric x positions for 5 zones
    zone_labels  = [z[0] for z in PERF_ZONES]

    for yr in years_to_plot:
        sc  = score_data.get(yr, np.array([]))
        if len(sc) == 0:
            continue
        cfg = BELL_CFG[yr]

        # headcount per zone
        hc = []
        for _, lo, hi in PERF_ZONES:
            hc.append(int(((sc >= lo) & (sc <= hi)).sum()))

        total_yr = sum(hc) or 1
        pcts = [h / total_yr * 100 for h in hc]

        # bars
        offset = offsets.get(yr, 0) if bell_sel == "All Years" else 0
        fig_bell.add_trace(go.Bar(
            x=[z + offset for z in zone_centers],
            y=hc,
            name=f"{yr} HC",
            marker_color=cfg["color"],
            marker_opacity=0.55,
            width=bar_group_width * 0.92,
            customdata=np.array(pcts).round(1),
            hovertemplate=(
                f"<b>{yr}</b><br>"
                "HC: %{y}<br>"
                "Share: %{customdata}%<br>"
                "<extra></extra>"
            ),
        ))

        # smooth KDE line over the continuous score range
        hist_vals, bin_edges = np.histogram(sc, bins=40, density=True)
        bin_c    = (bin_edges[:-1] + bin_edges[1:]) / 2
        smoothed = pd.Series(hist_vals).rolling(5, center=True, min_periods=1).mean().values
        # scale KDE to match bar heights for visual alignment
        scale = (max(hc) / max(smoothed)) if max(smoothed) > 0 else 1
        # map continuous scores to zone x positions for the line
        def score_to_zone_x(s):
            for xi, (_, lo, hi) in enumerate(PERF_ZONES):
                if lo <= s <= hi:
                    return zone_centers[xi]
            return zone_centers[2]

        zone_x_line = np.array([score_to_zone_x(s) for s in bin_c])
        # aggregate smoothed values per zone x
        from collections import defaultdict
        zone_vals = defaultdict(list)
        for zx, sv in zip(zone_x_line, smoothed * scale):
            zone_vals[zx].append(sv)
        zx_sorted = sorted(zone_vals.keys())
        zy_sorted = [np.mean(zone_vals[zx]) for zx in zx_sorted]

        fig_bell.add_trace(go.Scatter(
            x=zx_sorted,
            y=zy_sorted,
            mode="lines",
            name=f"{yr}",
            line=dict(color=cfg["color"], width=3, dash=cfg["dash"]),
            hovertemplate=f"<b>{yr} trend</b><extra></extra>",
        ))

    # ── reference guideline (red curve, like reference image) ─────────────────
    # map ref_x (1–5 score) to zone centres
    from collections import defaultdict as _dd
    zone_ref = _dd(list)
    for xi, yi in zip(ref_x, ref_y):
        for zi, (_, lo, hi) in enumerate(PERF_ZONES):
            if lo <= xi <= hi:
                zone_ref[zone_centers[zi]].append(yi)
                break
    if zone_ref:
        zxr = sorted(zone_ref.keys())
        # scale to match bar height
        all_hc_flat = []
        for yr in years_to_plot:
            sc = score_data.get(yr, np.array([]))
            for _, lo, hi in PERF_ZONES:
                all_hc_flat.append(int(((sc >= lo) & (sc <= hi)).sum()))
        max_hc = max(all_hc_flat) if all_hc_flat else 1
        ref_vals_raw = [np.mean(zone_ref[z]) for z in zxr]
        ref_scale = max_hc / max(ref_vals_raw) if max(ref_vals_raw) > 0 else 1
        zyr = [v * ref_scale for v in ref_vals_raw]
        fig_bell.add_trace(go.Scatter(
            x=zxr, y=zyr,
            mode="lines",
            name="Distribution Guideline",
            line=dict(color=C_PURPLE, width=2.5),
            hoverinfo="skip",
        ))

    # ── Headcount annotations above bars with staggered offsets to prevent overlap ──
    label_shift_all = {2025: 22, 2024: 12, 2023: 2}
    for yr in years_to_plot:
        sc  = score_data.get(yr, np.array([]))
        if len(sc) == 0:
            continue
        offset = offsets.get(yr, 0) if bell_sel == "All Years" else 0
        label_shift = label_shift_all.get(yr, 12) if bell_sel == "All Years" else 14
        for zi, (_, lo, hi) in enumerate(PERF_ZONES):
            hc_val = int(((sc >= lo) & (sc <= hi)).sum())
            if hc_val > 0:
                fig_bell.add_annotation(
                    x=zone_centers[zi] + offset,
                    y=hc_val,
                    text=f"<b>{hc_val}</b>",
                    showarrow=False,
                    yshift=label_shift,
                    font=dict(size=10, color=BELL_CFG[yr]["color"]),
                )

    fig_bell.update_layout(
        height=420,
        plot_bgcolor="#FAFBFC",
        paper_bgcolor=C_WHITE,
        barmode="group",
        xaxis=dict(
            tickvals=zone_centers,
            ticktext=zone_labels,
            showgrid=False,
            tickfont=dict(size=10, color="#374151"),
            title=dict(text="Performance Zone", font=dict(size=11, color="#6B7280")),
        ),
        yaxis=dict(
            title=dict(text="HC", font=dict(size=11, color="#6B7280")),
            showgrid=True,
            gridcolor=C_LGRAY,
            zeroline=True,
            zerolinecolor=C_LGRAY,
        ),
        legend=dict(
            orientation="h",
            y=-0.22,
            x=0.5,
            xanchor="center",
            font=dict(size=10),
            bgcolor="rgba(255,255,255,0.8)",
            bordercolor=C_LGRAY,
            borderwidth=1,
        ),
        margin=dict(l=50, r=20, t=30, b=90),
        font=dict(family="Inter, Arial", size=10),
    )
    st.plotly_chart(fig_bell, use_container_width=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Row 2: Stacked bar + Avg score ────────────────────────────────────────    # ── Row 2: Stacked bar + Avg score ────────────────────────────────────────
    col_stack, col_avg = st.columns([1, 1], gap="medium")

    with col_stack:
        st.markdown(
            f'<div style="font-size:13px;font-weight:700;color:{C_DARK};margin-bottom:8px">'
            f'Full Distribution by Category & Year</div>',
            unsafe_allow_html=True,
        )
        fig_st = px.bar(
            dist, x="year", y="count", color="category",
            color_discrete_map={
                GRID_META[k]["label"]: BOX_COLORS[k]["dot"] for k in GRID_META
            },
            barmode="stack",
            labels={"year": "Year", "count": "Employees", "category": "Category"},
        )
        fig_st.update_layout(
            height=320,
            plot_bgcolor="#FAFBFC", paper_bgcolor=C_WHITE,
            margin=dict(l=40, r=20, t=10, b=40),
            legend=dict(orientation="h", y=-0.40, font=dict(size=9)),
            font=dict(family="Inter, Arial", size=10),
            xaxis=dict(showgrid=False),
            yaxis=dict(showgrid=True, gridcolor=C_LGRAY),
        )
        st.plotly_chart(fig_st, use_container_width=True)

    with col_avg:
        st.markdown(
            f'<div style="font-size:13px;font-weight:700;color:{C_DARK};margin-bottom:8px">'
            f'Average Performance Score by Year</div>',
            unsafe_allow_html=True,
        )
        bar_clrs = [C_PURPLE, C_GREEN, "#A78BFA"]
        fig_avg = go.Figure(go.Bar(
            x=trend["year"].astype(str),
            y=trend["avg_score"].round(2),
            marker_color=bar_clrs,
            text=trend["avg_score"].round(2),
            textposition="outside",
            textfont=dict(size=11, color=C_DARK),
            width=0.4,
        ))
        fig_avg.update_layout(
            height=320,
            plot_bgcolor="#FAFBFC", paper_bgcolor=C_WHITE,
            xaxis=dict(showgrid=False),
            yaxis=dict(range=[0, 5.5], showgrid=True, gridcolor=C_LGRAY),
            margin=dict(l=40, r=20, t=10, b=40),
            font=dict(family="Inter, Arial", size=10),
            showlegend=False,
        )
        st.plotly_chart(fig_avg, use_container_width=True)

    # ── Row 3: Individual movement table ─────────────────────────────────────
    st.markdown(
        f'<div class="sec-head" style="margin-top:4px">Individual Grid Movement — All Years</div>',
        unsafe_allow_html=True,
    )
    _pm = pd.Series(True, index=df_all.index)
    if sel_dept != "All":
        _pm &= df_all["dept"] == sel_dept
    if sel_loc != "All":
        _pm &= df_all["location"] == sel_loc
    pivot = (
        df_all[_pm]
        .pivot_table(
            index=["name", "dept", "location"],
            columns="year",
            values="grid_zone",
            aggfunc="first",
        )
        .reset_index()
    )
    pivot.columns.name = None
    for yr in YEARS:
        if yr not in pivot.columns:
            pivot[yr] = np.nan

    def mvmt(g1, g2):
        if pd.isna(g1) or pd.isna(g2): return "🆕 New"
        s1 = GRID_META.get(g1, {}).get("perf", 0) + GRID_META.get(g1, {}).get("pot", 0)
        s2 = GRID_META.get(g2, {}).get("perf", 0) + GRID_META.get(g2, {}).get("pot", 0)
        return "⬆️ Improved" if s2 > s1 else "➡️ Stable" if s2 == s1 else "⬇️ Declined"

    pivot["Movement 23→25"] = pivot.apply(
        lambda r: mvmt(r.get(2023), r.get(2025)), axis=1
    )
    pivot = pivot.rename(columns={
        "name": "Name", "dept": "Dept", "location": "Location",
        2023: "Grid 2023", 2024: "Grid 2024", 2025: "Grid 2025",
    })
    for c in ["Grid 2023", "Grid 2024", "Grid 2025"]:
        if c in pivot.columns:
            pivot[c] = pivot[c].fillna("—")
    st.dataframe(pivot.reset_index(drop=True), use_container_width=True, height=340)


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
