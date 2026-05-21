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
[data-testid="stSidebar"] {{ background:{C_SIDEBAR}; }}
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
# TAB 1  —  9-BOX MATRIX
# ══════════════════════════════════════════════════════════════════════════════
with tab1:
    st.markdown(
        f'<div class="sec-head">🔲 9-Box Talent Matrix'
        f'<span class="pill">{sel_year}</span></div>',
        unsafe_allow_html=True,
    )

    left_col, right_col = st.columns([3, 1], gap="medium")

    with left_col:
        # ── Plotly 9-box scatter ──────────────────────────────────────────────
        rng     = np.random.default_rng(seed=42)
        plot_df = gdf.copy()
        jitter  = 0.19
        plot_df["jx"] = (plot_df["perf_tier"].astype(float)
                         + rng.uniform(-jitter, jitter, len(plot_df)))
        plot_df["jy"] = (plot_df["pot_tier"].astype(float)
                         + rng.uniform(-jitter, jitter, len(plot_df)))

        fig = go.Figure()

        # cell backgrounds + labels
        for code in CAT_ORDER:
            meta = GRID_META[code]
            bc   = BOX_COLORS[code]
            px_c, py_c = meta["perf"], meta["pot"]
            cnt = int((plot_df["grid_zone"] == code).sum())

            fig.add_shape(
                type="rect",
                x0=px_c - 0.5, x1=px_c + 0.5,
                y0=py_c - 0.5, y1=py_c + 0.5,
                fillcolor=bc["bg"],
                line=dict(color=bc["border"], width=1.5),
                layer="below",
            )
            # code (top-left)
            fig.add_annotation(
                x=px_c - 0.44, y=py_c + 0.42,
                text=f"<b>{code}</b>",
                showarrow=False,
                font=dict(size=9, color=bc["dot"]),
                xanchor="left", yanchor="top", opacity=0.7,
            )
            # label (below code)
            fig.add_annotation(
                x=px_c - 0.44, y=py_c + 0.30,
                text=f"<b>{meta['label']}</b>",
                showarrow=False,
                font=dict(size=9, color=bc["dot"]),
                xanchor="left", yanchor="top", opacity=0.9,
            )
            # count (bottom-right)
            fig.add_annotation(
                x=px_c + 0.43, y=py_c - 0.36,
                text=f"<b>n = {cnt}</b>",
                showarrow=False,
                font=dict(size=10, color=bc["dot"]),
                xanchor="right", yanchor="bottom", opacity=0.85,
            )

        # dots + first-name labels
        for code in CAT_ORDER:
            bc   = BOX_COLORS[code]
            sub  = plot_df[plot_df["grid_zone"] == code]
            if sub.empty:
                continue
            hover = (
                "<b>" + sub["name"] + "</b><br>"
                + "Dept: "     + sub["dept"].fillna("—") + "<br>"
                + "Location: " + sub["location"].fillna("—") + "<br>"
                + "Grid: "     + code + " · " + GRID_META[code]["label"] + "<br>"
                + "Score: "    + sub["overall_score"].apply(
                    lambda v: f"{v:.2f}" if not np.isnan(v) else "—")
            )
            fig.add_trace(go.Scatter(
                x=sub["jx"], y=sub["jy"],
                mode="markers+text",
                name=f"{code} {GRID_META[code]['label']}",
                marker=dict(
                    color=bc["dot"], size=9, opacity=0.85,
                    line=dict(width=1, color=C_WHITE),
                ),
                text=sub["name"].str.split().str[0],
                textposition="top center",
                textfont=dict(size=7, color=C_DARK),
                hovertext=hover,
                hoverinfo="text",
                showlegend=True,
            ))

        # axis zone labels
        for v, lbl in [(1, "Poor"), (2, "Average"), (3, "High")]:
            fig.add_annotation(
                x=v, y=0.45, text=f"<b>{lbl}</b>",
                showarrow=False,
                font=dict(size=9, color="#6B7280"), yanchor="top",
            )
        for v, lbl in [(1, "Low"), (2, "Medium"), (3, "High")]:
            fig.add_annotation(
                x=0.46, y=v, text=f"<b>{lbl}</b>",
                showarrow=False,
                font=dict(size=9, color="#6B7280"), xanchor="right",
            )
        # dividers
        for v in [1.5, 2.5]:
            fig.add_vline(x=v, line=dict(color="rgba(107,114,128,.35)", width=1, dash="dot"))
            fig.add_hline(y=v, line=dict(color="rgba(107,114,128,.35)", width=1, dash="dot"))

        fig.update_layout(
            height=580,
            plot_bgcolor="#FAFBFC",
            paper_bgcolor=C_WHITE,
            xaxis=dict(
                range=[0.42, 3.58], tickvals=[1, 2, 3], ticktext=["", "", ""],
                showgrid=False, zeroline=False,
                title="Performance  →",
                title_font=dict(size=11, color="#6B7280"),
            ),
            yaxis=dict(
                range=[0.42, 3.58], tickvals=[1, 2, 3], ticktext=["", "", ""],
                showgrid=False, zeroline=False,
                title="Potential  →",
                title_font=dict(size=11, color="#6B7280"),
            ),
            margin=dict(l=60, r=20, t=20, b=70),
            legend=dict(
                orientation="h", yanchor="bottom", y=-0.14,
                xanchor="center", x=0.5,
                font=dict(size=9), itemsizing="constant",
            ),
            font=dict(family="Inter, Arial, sans-serif"),
        )
        st.plotly_chart(fig, use_container_width=True)

    with right_col:
        # ── Category count bar ────────────────────────────────────────────────
        st.markdown(
            f'<div style="font-size:12px;font-weight:700;color:{C_DARK};'
            f'margin-bottom:8px">Category Counts</div>',
            unsafe_allow_html=True,
        )
        labels_bar = [f"{k} {GRID_META[k]['label']}" for k in CAT_ORDER]
        counts_bar = [int((gdf["grid_zone"] == k).sum()) for k in CAT_ORDER]
        colors_bar = [BOX_COLORS[k]["dot"] for k in CAT_ORDER]

        fig_b = go.Figure(go.Bar(
            x=counts_bar, y=labels_bar,
            orientation="h",
            marker_color=colors_bar,
            text=counts_bar,
            textposition="outside",
            textfont=dict(size=10, color=C_DARK),
        ))
        fig_b.update_layout(
            height=320,
            plot_bgcolor=C_WHITE, paper_bgcolor=C_WHITE,
            margin=dict(l=10, r=40, t=10, b=10),
            xaxis=dict(showgrid=True, gridcolor=C_LGRAY, zeroline=False),
            yaxis=dict(showgrid=False, autorange="reversed"),
            font=dict(size=10, family="Inter, Arial"),
            showlegend=False,
        )
        st.plotly_chart(fig_b, use_container_width=True)

        # ── Names per box ─────────────────────────────────────────────────────
        st.markdown(
            f'<div style="font-size:12px;font-weight:700;color:{C_DARK};'
            f'margin:12px 0 6px">Employee Names by Box</div>',
            unsafe_allow_html=True,
        )
        for code in CAT_ORDER:
            bc   = BOX_COLORS[code]
            sub  = gdf[gdf["grid_zone"] == code]
            if sub.empty:
                continue
            with st.expander(
                f"{code}  {GRID_META[code]['label']}  ({len(sub)})"
            ):
                for nm in sorted(sub["name"].tolist()):
                    st.markdown(
                        f'<span style="font-size:11px;color:{bc["dot"]}">• {nm}</span>',
                        unsafe_allow_html=True,
                    )


# ══════════════════════════════════════════════════════════════════════════════
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

    # ── Row 1: Bell curve + talent movement ──────────────────────────────────
    col_bell, col_move = st.columns([1, 1], gap="medium")

    with col_bell:
        st.markdown(
            f'<div style="font-size:13px;font-weight:700;color:{C_DARK};margin-bottom:8px">'
            f'Performance Distribution — Bell Curve</div>',
            unsafe_allow_html=True,
        )
        fig_bell = go.Figure()
        bell_colors = {2023: "#A78BFA", 2024: C_GREEN, 2025: C_PURPLE}
        bell_dash   = {2023: "dot", 2024: "dash", 2025: "solid"}

        for yr in YEARS:
            sc = score_data[yr]
            if len(sc) < 3:
                continue
            # KDE via histogram + smooth
            hist_vals, bin_edges = np.histogram(sc, bins=30, density=True)
            bin_centres = (bin_edges[:-1] + bin_edges[1:]) / 2
            # Gaussian smoothing via simple rolling mean
            smoothed = pd.Series(hist_vals).rolling(3, center=True, min_periods=1).mean().values

            fig_bell.add_trace(go.Scatter(
                x=bin_centres, y=smoothed,
                mode="lines",
                name=str(yr),
                fill="tozeroy",
                fillcolor=f"rgba{tuple(int(bell_colors[yr].lstrip('#')[i:i+2],16) for i in (0,2,4))+(0.10,)}",
                line=dict(color=bell_colors[yr], width=2.5, dash=bell_dash[yr]),
            ))
            # mean line
            mu = sc.mean()
            fig_bell.add_vline(
                x=mu,
                line=dict(color=bell_colors[yr], width=1.5, dash="dot"),
                annotation_text=f"{yr} μ={mu:.2f}",
                annotation_position="top",
                annotation_font=dict(size=9, color=bell_colors[yr]),
            )

        fig_bell.update_layout(
            height=320,
            plot_bgcolor="#FAFBFC", paper_bgcolor=C_WHITE,
            xaxis=dict(title="Performance Score", showgrid=True,
                       gridcolor=C_LGRAY, range=[1, 5.5]),
            yaxis=dict(title="Density", showgrid=True, gridcolor=C_LGRAY),
            legend=dict(orientation="h", y=-0.20, font=dict(size=10)),
            margin=dict(l=50, r=20, t=20, b=60),
            font=dict(family="Inter, Arial", size=10),
        )
        st.plotly_chart(fig_bell, use_container_width=True)

    with col_move:
        st.markdown(
            f'<div style="font-size:13px;font-weight:700;color:{C_DARK};margin-bottom:8px">'
            f'Key Talent Tier Movement</div>',
            unsafe_allow_html=True,
        )
        line_cfg = {
            "Stars (3C)":       (C_PURPLE, "solid"),
            "High Performers":  (C_GREEN,  "dash"),
            "Growth Employees": ("#A78BFA","dot"),
            "At Risk (1A)":     ("#991B1B","dot"),
        }
        fig_mv = go.Figure()
        for col_name, (clr, dsh) in line_cfg.items():
            fig_mv.add_trace(go.Scatter(
                x=trend["year"], y=trend[col_name],
                name=col_name, mode="lines+markers+text",
                line=dict(color=clr, width=2.5, dash=dsh),
                marker=dict(size=8, color=clr),
                text=trend[col_name].astype(int),
                textposition="top center",
                textfont=dict(size=9, color=clr),
            ))
        fig_mv.update_layout(
            height=320,
            plot_bgcolor="#FAFBFC", paper_bgcolor=C_WHITE,
            xaxis=dict(tickvals=YEARS, showgrid=False),
            yaxis=dict(showgrid=True, gridcolor=C_LGRAY),
            legend=dict(orientation="h", y=-0.22, font=dict(size=9)),
            margin=dict(l=40, r=20, t=20, b=70),
            font=dict(family="Inter, Arial", size=10),
        )
        st.plotly_chart(fig_mv, use_container_width=True)

    # ── Row 2: Stacked bar + Avg score ────────────────────────────────────────
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
