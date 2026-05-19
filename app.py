"""
MBC Media Solutions — Talent 9-Grid Dashboard
=============================================
Data source : 9Grid_Final.xlsx
             Sheets used: Evalutaion 25 / Evalutaion 24 / Evalutaion 23
             (These sheets contain Performance Zone + Potential zone + Grid Zone)

Column name constants are at the top — edit only the right-hand strings
if your file uses different headers.
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
import base64, os, warnings
warnings.filterwarnings("ignore")

# ─── PAGE CONFIG ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="MMS · Talent 9-Grid Dashboard",
    page_icon="🎯",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── GRID METADATA ───────────────────────────────────────────────────────────
GRID_META = {
    "3C": {"label": "Star",               "perf": 3, "pot": 3, "color": "#1F4E79", "bg": "rgba(31,78,121,0.12)"},
    "3B": {"label": "High Performer",     "perf": 3, "pot": 2, "color": "#166534", "bg": "rgba(22,101,52,0.12)"},
    "3A": {"label": "Solid Performer",    "perf": 3, "pot": 1, "color": "#0F766E", "bg": "rgba(15,118,110,0.12)"},
    "2C": {"label": "Growth Employee",    "perf": 2, "pot": 3, "color": "#6D28D9", "bg": "rgba(109,40,217,0.12)"},
    "2B": {"label": "Core Player",        "perf": 2, "pot": 2, "color": "#0369A1", "bg": "rgba(3,105,161,0.09)"},
    "2A": {"label": "Average Performer",  "perf": 2, "pot": 1, "color": "#B45309", "bg": "rgba(180,83,9,0.12)"},
    "1C": {"label": "Potential Gem",      "perf": 1, "pot": 3, "color": "#7E22CE", "bg": "rgba(126,34,206,0.12)"},
    "1B": {"label": "Inconsistent Player","perf": 1, "pot": 2, "color": "#C2410C", "bg": "rgba(194,65,12,0.12)"},
    "1A": {"label": "Risk",               "perf": 1, "pot": 1, "color": "#991B1B", "bg": "rgba(153,27,27,0.12)"},
}
VALID_CODES = set(GRID_META.keys())

# ─── CSS ─────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
[data-testid="stAppViewContainer"]  { background:#F4F6FB; }
[data-testid="stSidebar"]           { background:#0D1B2A; }
[data-testid="stSidebar"] *         { color:#CBD5E1 !important; }
[data-testid="stSidebar"] h1,
[data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3        { color:#F1F5F9 !important; }
[data-testid="stSidebar"] .stSelectbox label,
[data-testid="stSidebar"] .stMultiSelect label {
  color:#94A3B8 !important; font-size:11px !important;
  letter-spacing:.04em; text-transform:uppercase;
}
[data-testid="stSidebar"] hr { border-color:#1E3A5F !important; }

.dash-header {
  background:linear-gradient(135deg,#0D1B2A 0%,#1E3A5F 100%);
  border-radius:14px; padding:22px 28px; margin-bottom:20px;
  display:flex; align-items:center; gap:20px;
}
.dash-header .titles h1 { color:#F1F5F9; font-size:22px; margin:0; font-weight:700; }
.dash-header .titles p  { color:#94A3B8; margin:3px 0 0; font-size:12px; }
.badge { background:rgba(46,117,182,.25); color:#93C5FD; border-radius:999px;
         padding:3px 12px; font-size:11px; font-weight:600; display:inline-block; margin-top:6px; }

.kpi { background:#fff; border-radius:12px; padding:14px 16px;
       border-left:4px solid #2E75B6; box-shadow:0 1px 4px rgba(0,0,0,.06);
       margin-bottom:4px; }
.kpi.g  { border-left-color:#166534; }
.kpi.v  { border-left-color:#6D28D9; }
.kpi.a  { border-left-color:#B45309; }
.kpi.r  { border-left-color:#991B1B; }
.kpi.t  { border-left-color:#0F766E; }
.kpi.s  { border-left-color:#0369A1; }
.kpi-lbl  { font-size:10px; font-weight:600; color:#64748B;
            text-transform:uppercase; letter-spacing:.06em; margin-bottom:2px; }
.kpi-val  { font-size:24px; font-weight:700; color:#0F172A; line-height:1.1; }
.kpi-sub  { font-size:10px; color:#94A3B8; margin-top:2px; }
.kpi-delta { font-size:11px; font-weight:600; }
.kpi-delta.up { color:#166534; }
.kpi-delta.dn { color:#991B1B; }

.stitle { font-size:14px; font-weight:700; color:#1E3A5F;
          border-left:3px solid #2E75B6; padding-left:9px;
          margin:20px 0 10px; }

.succ-card { background:#fff; border-radius:10px; padding:14px 16px;
             border:0.5px solid #E2E8F0; box-shadow:0 1px 3px rgba(0,0,0,.05);
             margin-bottom:8px; }
.succ-name  { font-size:13px; font-weight:600; color:#1E293B; }
.succ-dept  { font-size:11px; color:#64748B; }
.succ-badge { display:inline-block; font-size:10px; font-weight:600;
              padding:2px 8px; border-radius:999px; margin-top:4px; }
.ready-high   { background:#DCFCE7; color:#166534; }
.ready-medium { background:#FEF3C7; color:#92400E; }
.ready-low    { background:#FEE2E2; color:#991B1B; }
</style>
""", unsafe_allow_html=True)


# ─── DATA LOADING ─────────────────────────────────────────────────────────────
@st.cache_data(show_spinner="Loading talent data…")
def load_all_years(path: str) -> pd.DataFrame:
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)

    def load_sheet(sheet_name: str, year: int) -> pd.DataFrame:
        ws = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))
        # find header row (contains 'Employee Name')
        hdr_idx = next(
            (i for i, r in enumerate(rows) if any(c == "Employee Name" for c in r)), None)
        if hdr_idx is None:
            return pd.DataFrame()
        headers = list(rows[hdr_idx])
        data = [r for r in rows[hdr_idx + 1:]
                if any(c is not None for c in r[:6])]
        df = pd.DataFrame(data, columns=headers)

        # ── rename to standard names ──
        rename = {}
        for c in df.columns:
            cs = str(c).strip().lower() if c else ""
            if c == "Employee Name":               rename[c] = "name"
            elif c == "Employee Number":           rename[c] = "emp_id"
            elif c == "Department":                rename[c] = "dept"
            elif c == "Location":                  rename[c] = "location"
            elif c == "Position":                  rename[c] = "position"
            elif c == "Joining Date":              rename[c] = "join_date"
            elif "objectives and comp" in cs:      rename[c] = "overall_score"
            elif "objectives score" in cs:         rename[c] = "obj_score"
            elif "overall comp" in cs:             rename[c] = "comp_score"
            elif ("perfor" in cs) and ("zone" in cs): rename[c] = "perf_zone"
            elif ("potential" in cs):              rename[c] = "pot_zone"
            elif ("grid" in cs or "gird" in cs) and ("zone" in cs): rename[c] = "grid_zone"
        df = df.rename(columns=rename)
        df["year"] = year

        # ── clean ──
        NULL_SET = {"NA", "NR", "-", "", "nan", "None"}

        def clean_num(v):
            if v is None: return np.nan
            if isinstance(v, (int, float)):
                return float(v) if (not np.isnan(float(v)) and float(v) > 0) else np.nan
            sv = str(v).strip()
            if sv in NULL_SET: return np.nan
            try: return float(sv) if float(sv) > 0 else np.nan
            except: return np.nan

        for col in ["overall_score", "obj_score", "comp_score"]:
            if col in df.columns:
                df[col] = df[col].apply(clean_num)

        # clean grid_zone
        if "grid_zone" in df.columns:
            df["grid_zone"] = df["grid_zone"].apply(
                lambda v: str(v).strip() if str(v).strip() in VALID_CODES else np.nan)
        else:
            df["grid_zone"] = np.nan

        # clean text zones
        for col in ["perf_zone", "pot_zone"]:
            if col in df.columns:
                df[col] = df[col].apply(
                    lambda v: str(v).strip()
                    if v and str(v).strip() not in NULL_SET else np.nan)
            else:
                df[col] = np.nan

        # normalise location
        loc_map = {
            "MMS KSA": "Saudi Arabia", "MMS UAE": "UAE", "MMS Egypt": "Egypt",
            "Riyadh": "Saudi Arabia",  "Jeddah": "Saudi Arabia",
            "Dubai":  "UAE",           "Cairo":  "Egypt",
        }
        if "location" in df.columns:
            df["location"] = df["location"].apply(
                lambda v: loc_map.get(str(v).strip(), str(v).strip()) if v else np.nan)

        if "dept" in df.columns:
            df["dept"] = df["dept"].apply(
                lambda v: str(v).strip() if v else np.nan)

        # derived tiers
        df["perf_tier"] = df["grid_zone"].apply(
            lambda v: int(v[0]) if isinstance(v, str) and len(v) == 2 else np.nan)
        df["pot_tier"] = df["grid_zone"].apply(
            lambda v: {"A": 1, "B": 2, "C": 3}.get(v[1])
            if isinstance(v, str) and len(v) == 2 else np.nan)
        df["grid_label"] = df["grid_zone"].map(
            {k: m["label"] for k, m in GRID_META.items()})

        return df

    df25 = load_sheet("Evalutaion 25", 2025)
    df24 = load_sheet("Evalutaion 24", 2024)
    df23 = load_sheet("Evalutaion 23", 2023)

    combined = pd.concat([df25, df24, df23], ignore_index=True)
    combined["name"] = combined["name"].astype(str).str.strip()
    combined = combined[
        combined["name"].notna() & (combined["name"] != "nan") & (combined["name"] != "")]
    return combined


# ─── LOAD ─────────────────────────────────────────────────────────────────────
DATA_PATH = "9Grid_Final.xlsx"
try:
    df_all = load_all_years(DATA_PATH)
except FileNotFoundError:
    st.error(f"❌ Excel file not found: `{DATA_PATH}`.")
    st.stop()


# ─── LOGO ─────────────────────────────────────────────────────────────────────
def img_b64(path):
    try:
        with open(path, "rb") as f:
            return base64.b64encode(f.read()).decode()
    except Exception:
        return None

logo_b64 = img_b64("logo.png")


# ─── SIDEBAR ──────────────────────────────────────────────────────────────────
with st.sidebar:
    if logo_b64:
        st.markdown(
            f'<div style="text-align:center;padding:16px 0 8px">'
            f'<img src="data:image/png;base64,{logo_b64}" '
            f'style="max-width:180px;width:100%;border-radius:6px"></div>',
            unsafe_allow_html=True)
    st.markdown("---")

    sel_year = st.selectbox("📅 Year", [2025, 2024, 2023], index=0)

    all_depts = sorted(df_all["dept"].dropna().unique())
    sel_depts = st.multiselect("🏢 Department", all_depts, default=all_depts)

    all_locs = sorted(df_all["location"].dropna().unique())
    sel_locs = st.multiselect("📍 Location / Country", all_locs, default=all_locs)

    all_roles = sorted(df_all["position"].dropna().unique())
    sel_roles = st.multiselect("💼 Role / Position", all_roles, default=all_roles)

    all_names = sorted(
        df_all[df_all["year"] == sel_year]["name"].dropna().unique())
    sel_names = st.multiselect("👤 Employee Name (optional)", all_names, default=[])

    st.markdown("---")
    all_cats = [GRID_META[k]["label"]
                for k in ["3C", "3B", "3A", "2C", "2B", "2A", "1C", "1B", "1A"]]
    sel_cats = st.multiselect("📊 Grid Category", all_cats, default=all_cats)

    st.caption("Source: 9Grid_Final.xlsx · Evalutaion sheets 23/24/25")


# ─── FILTER HELPER ────────────────────────────────────────────────────────────
def apply_filters(df, year, include_cat=True, include_name=True):
    m = ((df["year"] == year) &
         df["dept"].isin(sel_depts) &
         df["location"].isin(sel_locs))
    if sel_roles:
        m &= df["position"].isin(sel_roles)
    if include_name and sel_names:
        m &= df["name"].isin(sel_names)
    if include_cat and sel_cats:
        m &= df["grid_label"].isin(sel_cats)
    return df[m].copy()


df       = apply_filters(df_all, sel_year)
df_base  = apply_filters(df_all, sel_year, include_cat=False, include_name=False)


# ─── HEADER ───────────────────────────────────────────────────────────────────
logo_html = (f'<img src="data:image/png;base64,{logo_b64}" style="height:52px;border-radius:6px">'
             if logo_b64 else "🎯")
st.markdown(f"""
<div class="dash-header">
  <div>{logo_html}</div>
  <div class="titles">
    <h1>Talent 9-Grid Dashboard</h1>
    <p>Executive HR Analytics · <strong style="color:#93C5FD">{sel_year}</strong>
       · <strong style="color:#93C5FD">{len(df[df["grid_zone"].notna()])}</strong>
       evaluated employees</p>
    <span class="badge">MBC Media Solutions</span>
  </div>
</div>
""", unsafe_allow_html=True)


# ─── KPI CARDS ────────────────────────────────────────────────────────────────
def kpi_card(label, val, sub="", cls="", delta="", ddir=""):
    d = f'<div class="kpi-delta {ddir}">{delta}</div>' if delta else ""
    return (f'<div class="kpi {cls}">'
            f'<div class="kpi-lbl">{label}</div>'
            f'<div class="kpi-val">{val}</div>'
            f'<div class="kpi-sub">{sub}</div>{d}</div>')

gdf = df[df["grid_zone"].notna()]
total   = len(gdf)
stars   = int((gdf["grid_zone"] == "3C").sum())
hiperf  = int((gdf["grid_zone"] == "3B").sum())
growth  = int((gdf["grid_zone"] == "2C").sum())
at_risk = int((gdf["grid_zone"] == "1A").sum())
hipot   = int((gdf["pot_tier"]  == 3).sum())
avg_sc  = gdf["overall_score"].dropna().mean()
avg_pot = gdf["pot_tier"].dropna().mean()

prev_yr   = sel_year - 1
prev_base = df_all[(df_all["year"] == prev_yr) &
                   df_all["dept"].isin(sel_depts) &
                   df_all["location"].isin(sel_locs)]
prev_stars = int((prev_base["grid_zone"] == "3C").sum()) if len(prev_base) else None

if prev_stars is not None:
    delta_n = stars - prev_stars
    s_delta = f"{'▲' if delta_n >= 0 else '▼'} {abs(delta_n)} vs {prev_yr}"
    s_ddir  = "up" if delta_n >= 0 else "dn"
else:
    s_delta, s_ddir = "baseline year", ""

c1, c2, c3, c4, c5, c6, c7 = st.columns(7)
cards = [
    ("Total Employees",  total,  f"{sel_year} eval cycle", "s", "", ""),
    ("⭐ Stars (3C)",    stars,  f"{round(stars/total*100,1) if total else 0}% share", "g", s_delta, s_ddir),
    ("🟢 High Performers",hiperf,f"Grid 3B", "g", "", ""),
    ("🌱 Growth Employees",growth,f"High potential path","v","",""),
    ("🔴 At Risk (1A)",  at_risk,"Immediate action needed","r","",""),
    ("📊 Avg Score",     f"{avg_sc:.2f}" if not np.isnan(avg_sc) else "—","out of 5.0","t","",""),
    ("🎯 Avg Potential", f"{avg_pot:.2f}" if not np.isnan(avg_pot) else "—","tier 1–3","a","",""),
]
for col, (lbl, val, sub, cls, dlt, ddir) in zip([c1,c2,c3,c4,c5,c6,c7], cards):
    with col:
        st.markdown(kpi_card(lbl, val, sub, cls, dlt, ddir), unsafe_allow_html=True)


# ─── MAIN TABS ────────────────────────────────────────────────────────────────
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "🔲 9-Box Matrix",
    "📈 Year Comparison",
    "🏢 Department & Location",
    "🌟 Succession Planning",
    "📋 Employee Register",
])


# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — 9-BOX MATRIX
# ══════════════════════════════════════════════════════════════════════════════
with tab1:
    col_mat, col_side = st.columns([3, 1])

    with col_mat:
        st.markdown('<div class="stitle">Talent 9-Box Matrix</div>', unsafe_allow_html=True)
        plot_df = gdf.copy()
        rng = np.random.default_rng(seed=99)
        j = 0.20
        plot_df["jx"] = plot_df["perf_tier"].astype(float) + rng.uniform(-j, j, len(plot_df))
        plot_df["jy"] = plot_df["pot_tier"].astype(float)  + rng.uniform(-j, j, len(plot_df))

        fig = go.Figure()

        # cells
        for code, meta in GRID_META.items():
            px_c, py_c = meta["perf"], meta["pot"]
            fig.add_shape(type="rect",
                x0=px_c-.5, x1=px_c+.5, y0=py_c-.5, y1=py_c+.5,
                fillcolor=meta["bg"],
                line=dict(color="rgba(180,200,220,.4)", width=1), layer="below")
            fig.add_annotation(x=px_c-.44, y=py_c+.40,
                text=f"<b>{meta['label']}</b>", showarrow=False,
                font=dict(size=9, color=meta["color"]), xanchor="left", yanchor="top", opacity=.9)
            fig.add_annotation(x=px_c+.44, y=py_c+.40,
                text=f"<b>{code}</b>", showarrow=False,
                font=dict(size=8, color=meta["color"]), xanchor="right", yanchor="top", opacity=.6)
            cnt = int((plot_df["grid_zone"] == code).sum())
            fig.add_annotation(x=px_c+.42, y=py_c-.38,
                text=f"<b>n={cnt}</b>", showarrow=False,
                font=dict(size=9, color=meta["color"]), xanchor="right", yanchor="bottom", opacity=.85)

        # dots + names
        for code, meta in GRID_META.items():
            sub = plot_df[plot_df["grid_zone"] == code]
            if sub.empty:
                continue
            hover = ("<b>" + sub["name"] + "</b><br>"
                     + "Dept: "     + sub["dept"].fillna("—") + "<br>"
                     + "Location: " + sub["location"].fillna("—") + "<br>"
                     + "Role: "     + sub["position"].fillna("—") + "<br>"
                     + "Grid: " + code + " · " + meta["label"] + "<br>"
                     + "Score: "    + sub["overall_score"].round(2).astype(str))
            fig.add_trace(go.Scatter(
                x=sub["jx"], y=sub["jy"],
                mode="markers+text",
                name=meta["label"],
                marker=dict(color=meta["color"], size=9, opacity=.85,
                            line=dict(width=1, color="white")),
                text=sub["name"].str.split().str[0],
                textposition="top center",
                textfont=dict(size=7, color="#1E293B"),
                hovertext=hover, hoverinfo="text", showlegend=True,
            ))

        for v, lbl in [(1,"Poor"),(2,"Average"),(3,"High")]:
            fig.add_annotation(x=v, y=0.46, text=f"<b>{lbl}</b>",
                showarrow=False, font=dict(size=9, color="#64748B"), yanchor="top")
        for v, lbl in [(1,"Low"),(2,"Medium"),(3,"High")]:
            fig.add_annotation(x=0.47, y=v, text=f"<b>{lbl}</b>",
                showarrow=False, font=dict(size=9, color="#64748B"), xanchor="right")
        for v in [1.5, 2.5]:
            fig.add_vline(x=v, line=dict(color="rgba(148,163,184,.4)", width=1, dash="dot"))
            fig.add_hline(y=v, line=dict(color="rgba(148,163,184,.4)", width=1, dash="dot"))

        fig.update_layout(
            height=590,
            xaxis=dict(range=[.45,3.55], tickvals=[1,2,3], ticktext=["","",""],
                       showgrid=False, zeroline=False,
                       title="Performance →", title_font=dict(size=11,color="#475569")),
            yaxis=dict(range=[.45,3.55], tickvals=[1,2,3], ticktext=["","",""],
                       showgrid=False, zeroline=False,
                       title="← Potential", title_font=dict(size=11,color="#475569")),
            plot_bgcolor="#FAFBFC", paper_bgcolor="#FFFFFF",
            margin=dict(l=60,r=20,t=20,b=70),
            legend=dict(orientation="h", yanchor="bottom", y=-0.14,
                        xanchor="center", x=0.5, font=dict(size=9), itemsizing="constant"),
            font=dict(family="Inter, Arial, sans-serif"),
        )
        st.plotly_chart(fig, use_container_width=True)

    with col_side:
        st.markdown('<div class="stitle">Category Count</div>', unsafe_allow_html=True)
        cat_order = ["3C","3B","3A","2C","2B","2A","1C","1B","1A"]
        labels = [GRID_META[k]["label"] for k in cat_order]
        counts = [int((gdf["grid_zone"]==k).sum()) for k in cat_order]
        colors = [GRID_META[k]["color"] for k in cat_order]
        fig_b = go.Figure(go.Bar(
            x=counts, y=labels, orientation="h",
            marker_color=colors,
            text=counts, textposition="outside", textfont=dict(size=10),
        ))
        fig_b.update_layout(
            height=350, margin=dict(l=10,r=40,t=10,b=10),
            plot_bgcolor="#FFFFFF", paper_bgcolor="#FFFFFF",
            xaxis=dict(showgrid=True, gridcolor="#F1F5F9", zeroline=False),
            yaxis=dict(showgrid=False, autorange="reversed"),
            font=dict(size=10, family="Inter, Arial"), showlegend=False,
        )
        st.plotly_chart(fig_b, use_container_width=True)

        st.markdown('<div class="stitle" style="margin-top:4px">Employee Names</div>',
                    unsafe_allow_html=True)
        for code in cat_order:
            meta = GRID_META[code]
            sub  = gdf[gdf["grid_zone"] == code]
            if sub.empty:
                continue
            with st.expander(f"{code} {meta['label']} ({len(sub)})"):
                for name in sorted(sub["name"].tolist()):
                    st.markdown(
                        f"<span style='font-size:11px;color:{meta['color']}'>• {name}</span>",
                        unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — YEAR COMPARISON
# ══════════════════════════════════════════════════════════════════════════════
with tab2:
    st.markdown('<div class="stitle">Year-over-Year Talent Movement (2023 → 2024 → 2025)</div>',
                unsafe_allow_html=True)

    years = [2023, 2024, 2025]
    trend_rows = []
    for yr in years:
        ydf = df_all[
            (df_all["year"] == yr) &
            df_all["dept"].isin(sel_depts) &
            df_all["location"].isin(sel_locs) &
            df_all["grid_zone"].notna()
        ]
        trend_rows.append({
            "year":        yr,
            "total":       len(ydf),
            "Stars (3C)":  int((ydf["grid_zone"]=="3C").sum()),
            "High Perf 3B":int((ydf["grid_zone"]=="3B").sum()),
            "Growth 2C":   int((ydf["grid_zone"]=="2C").sum()),
            "Risk 1A":     int((ydf["grid_zone"]=="1A").sum()),
            "High Pot C":  int((ydf["pot_tier"]==3).sum()),
            "avg_score":   ydf["overall_score"].dropna().mean(),
        })
    trend = pd.DataFrame(trend_rows)

    ca, cb = st.columns(2)
    with ca:
        fig_tl = go.Figure()
        line_cfg = {
            "Stars (3C)":   ("#1F4E79","solid"),
            "High Perf 3B": ("#166534","dash"),
            "Growth 2C":    ("#6D28D9","dot"),
            "High Pot C":   ("#0369A1","dashdot"),
            "Risk 1A":      ("#991B1B","dot"),
        }
        for col, (clr, dsh) in line_cfg.items():
            fig_tl.add_trace(go.Scatter(
                x=trend["year"], y=trend[col], name=col,
                mode="lines+markers+text",
                line=dict(color=clr, width=2.5, dash=dsh),
                marker=dict(size=8, color=clr),
                text=trend[col].astype(int),
                textposition="top center", textfont=dict(size=9),
            ))
        fig_tl.update_layout(
            title="Talent Category Trend 2023–2025",
            title_font=dict(size=13,color="#1E293B"),
            height=340, plot_bgcolor="#FAFBFC", paper_bgcolor="#FFFFFF",
            xaxis=dict(tickvals=years, showgrid=False),
            yaxis=dict(showgrid=True, gridcolor="#F1F5F9"),
            legend=dict(orientation="h", y=-0.30, font=dict(size=9)),
            margin=dict(l=40,r=20,t=50,b=100),
            font=dict(family="Inter, Arial", size=10),
        )
        st.plotly_chart(fig_tl, use_container_width=True)

    with cb:
        stack_rows = []
        for yr in years:
            ydf = df_all[
                (df_all["year"]==yr) &
                df_all["dept"].isin(sel_depts) &
                df_all["location"].isin(sel_locs) &
                df_all["grid_zone"].notna()
            ]
            for code in ["3C","3B","3A","2C","2B","2A","1C","1B","1A"]:
                stack_rows.append({
                    "year": str(yr),
                    "category": GRID_META[code]["label"],
                    "count": int((ydf["grid_zone"]==code).sum()),
                })
        sdf = pd.DataFrame(stack_rows)
        fig_st = px.bar(sdf, x="year", y="count", color="category",
            color_discrete_map={GRID_META[k]["label"]: GRID_META[k]["color"] for k in GRID_META},
            title="Full Distribution by Category & Year",
            labels={"year":"Year","count":"Employees","category":"Category"},
            barmode="stack",
        )
        fig_st.update_layout(
            height=340, plot_bgcolor="#FAFBFC", paper_bgcolor="#FFFFFF",
            margin=dict(l=40,r=20,t=50,b=40),
            title_font=dict(size=13,color="#1E293B"),
            legend=dict(orientation="h", y=-0.40, font=dict(size=9)),
            font=dict(family="Inter, Arial", size=10),
        )
        st.plotly_chart(fig_st, use_container_width=True)

    # individual movement
    st.markdown('<div class="stitle">Individual Employee Grid Movement</div>',
                unsafe_allow_html=True)
    pivot = (df_all[
        df_all["dept"].isin(sel_depts) & df_all["location"].isin(sel_locs)]
        .pivot_table(index=["name","dept","location"],
                     columns="year", values="grid_zone", aggfunc="first")
        .reset_index())
    pivot.columns.name = None
    for yr in [2023,2024,2025]:
        if yr not in pivot.columns:
            pivot[yr] = np.nan

    def mvmt(g1, g2):
        if pd.isna(g1) or pd.isna(g2): return "🆕 New"
        s1 = GRID_META.get(g1,{}).get("perf",0)+GRID_META.get(g1,{}).get("pot",0)
        s2 = GRID_META.get(g2,{}).get("perf",0)+GRID_META.get(g2,{}).get("pot",0)
        return ("⬆️ Improved" if s2>s1 else "➡️ Stable" if s2==s1 else "⬇️ Declined")

    pivot["Movement 23→25"] = pivot.apply(lambda r: mvmt(r.get(2023), r.get(2025)), axis=1)
    pivot = pivot.rename(columns={"name":"Name","dept":"Dept","location":"Location",
                                   2023:"Grid 2023",2024:"Grid 2024",2025:"Grid 2025"})
    for c in ["Grid 2023","Grid 2024","Grid 2025"]:
        if c in pivot.columns:
            pivot[c] = pivot[c].fillna("N/A")
    st.dataframe(pivot.reset_index(drop=True), use_container_width=True, height=340)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 — DEPARTMENT & LOCATION
# ══════════════════════════════════════════════════════════════════════════════
with tab3:
    st.markdown('<div class="stitle">Department & Location Intelligence</div>',
                unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    with c1:
        if not gdf.empty:
            dc = gdf.groupby(["dept","grid_label"]).size().reset_index(name="n")
            fig_dc = px.bar(dc, x="dept", y="n", color="grid_label",
                color_discrete_map={GRID_META[k]["label"]:GRID_META[k]["color"] for k in GRID_META},
                title="Grid Category by Department",
                labels={"dept":"Department","n":"Employees","grid_label":"Category"})
            fig_dc.update_layout(
                height=380, plot_bgcolor="#FAFBFC", paper_bgcolor="#FFFFFF",
                xaxis_tickangle=-40,
                legend=dict(orientation="h", y=-0.45, font=dict(size=9)),
                margin=dict(l=40,r=20,t=50,b=150),
                title_font=dict(size=13,color="#1E293B"),
                font=dict(family="Inter, Arial", size=10))
            st.plotly_chart(fig_dc, use_container_width=True)

    with c2:
        da = (gdf.groupby("dept")["overall_score"].mean().dropna()
              .sort_values().reset_index())
        if not da.empty:
            hi = da["overall_score"].quantile(.67)
            lo = da["overall_score"].quantile(.33)
            clrs = ["#1F4E79" if v>=hi else "#3B82F6" if v>=lo else "#93C5FD"
                    for v in da["overall_score"]]
            fig_da = go.Figure(go.Bar(
                x=da["overall_score"], y=da["dept"], orientation="h",
                marker_color=clrs,
                text=da["overall_score"].round(2), textposition="outside"))
            fig_da.update_layout(
                title="Avg Performance Score by Department",
                title_font=dict(size=13,color="#1E293B"),
                height=380, plot_bgcolor="#FAFBFC", paper_bgcolor="#FFFFFF",
                xaxis=dict(range=[0,5.6], showgrid=True, gridcolor="#F1F5F9"),
                yaxis=dict(showgrid=False),
                margin=dict(l=10,r=60,t=50,b=40),
                font=dict(family="Inter, Arial", size=10), showlegend=False)
            st.plotly_chart(fig_da, use_container_width=True)

    c3, c4 = st.columns(2)
    with c3:
        lc = gdf.groupby("location").size().reset_index(name="n")
        if not lc.empty:
            fig_loc = go.Figure(go.Pie(
                labels=lc["location"], values=lc["n"], hole=0.55,
                marker_colors=["#1F4E79","#166534","#B45309","#6D28D9"][:len(lc)],
                textinfo="label+percent", textfont=dict(size=10)))
            fig_loc.add_annotation(
                text=f"<b>{lc['n'].sum()}</b><br>Total",
                x=0.5, y=0.5, font_size=14, showarrow=False,
                font=dict(color="#1E293B"))
            fig_loc.update_layout(
                title="Distribution by Location",
                title_font=dict(size=13,color="#1E293B"),
                height=310, paper_bgcolor="#FFFFFF",
                margin=dict(l=20,r=20,t=50,b=20),
                legend=dict(orientation="h", y=-0.1),
                font=dict(family="Inter, Arial"))
            st.plotly_chart(fig_loc, use_container_width=True)

    with c4:
        if not gdf.empty:
            dh = (gdf.groupby("dept")
                  .apply(lambda g: pd.Series({
                      "total": len(g), "hp": (g["pot_tier"]==3).sum()}))
                  .reset_index())
            dh["hp%"] = dh["hp"] / dh["total"] * 100
            dh = dh.sort_values("hp%", ascending=True)
            fig_hp = go.Figure(go.Bar(
                x=dh["hp%"], y=dh["dept"], orientation="h",
                marker_color="#6D28D9",
                text=dh["hp%"].round(1).astype(str)+"%", textposition="outside"))
            fig_hp.update_layout(
                title="High Potential % by Department",
                title_font=dict(size=13,color="#1E293B"),
                height=310, plot_bgcolor="#FAFBFC", paper_bgcolor="#FFFFFF",
                xaxis=dict(range=[0,115], showgrid=True,
                           gridcolor="#F1F5F9", ticksuffix="%"),
                yaxis=dict(showgrid=False),
                margin=dict(l=10,r=70,t=50,b=40),
                font=dict(family="Inter, Arial", size=10), showlegend=False)
            st.plotly_chart(fig_hp, use_container_width=True)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 4 — SUCCESSION PLANNING
# ══════════════════════════════════════════════════════════════════════════════
with tab4:
    st.markdown('<div class="stitle">Succession Planning — Best Potential Successor by Department & Location</div>',
                unsafe_allow_html=True)
    st.caption(
        "Ranking methodology: Potential tier (×3) + 2025 Score (×2) − Score variability (σ). "
        "Only employees with a valid 2025 Grid Zone are considered.")

    # 3-year consistency
    emp_hist = (df_all[df_all["overall_score"].notna()]
                .groupby("name")["overall_score"]
                .agg(mean="mean", std="std", count="count")
                .reset_index())
    emp_hist["std"] = emp_hist["std"].fillna(0)
    emp_hist.columns = ["name","avg_3yr","std_3yr","yr_count"]

    base = df_all[
        (df_all["year"] == 2025) &
        df_all["grid_zone"].notna() &
        df_all["dept"].isin(sel_depts) &
        df_all["location"].isin(sel_locs)
    ].copy()
    base = base.merge(emp_hist, on="name", how="left")
    base["succ_score"] = (
        base["pot_tier"].fillna(1) * 3 +
        base["overall_score"].fillna(0) * 2 -
        base["std_3yr"].fillna(0))

    def readiness(row):
        if row["pot_tier"] == 3 and row["overall_score"] >= 4.0:
            return "High Ready", "ready-high", "🟢"
        if row["pot_tier"] >= 2 and row["overall_score"] >= 3.5:
            return "Ready", "ready-medium", "🟡"
        return "Developing", "ready-low", "🔴"

    base[["readiness","ready_cls","ready_icon"]] = base.apply(
        lambda r: pd.Series(readiness(r)), axis=1)

    # best successor per dept × location
    succ = (base.sort_values("succ_score", ascending=False)
            .groupby(["dept","location"]).first().reset_index())

    for loc in sorted(succ["location"].unique()):
        st.markdown(f"### 📍 {loc}")
        loc_df = succ[succ["location"] == loc].sort_values("dept")
        n_cols = min(3, len(loc_df))
        if n_cols == 0:
            continue
        cols = st.columns(n_cols)
        for i, (_, row) in enumerate(loc_df.iterrows()):
            code = row["grid_zone"]
            meta = GRID_META.get(code, {"label":"—","color":"#64748B"})
            sc3 = f"{row['avg_3yr']:.2f}" if not np.isnan(row["avg_3yr"]) else "—"
            with cols[i % n_cols]:
                st.markdown(f"""
                <div class="succ-card">
                  <div class="succ-name">{row['name']}</div>
                  <div class="succ-dept">{row['dept']} · {row['position'] if pd.notna(row.get('position')) else ''}</div>
                  <div style="font-size:11px;color:#475569;margin-top:5px">
                    <b>Grid:</b>
                    <span style="color:{meta['color']};font-weight:600">{code} — {meta['label']}</span><br>
                    <b>2025 Score:</b> {round(row['overall_score'],2) if not np.isnan(row['overall_score']) else '—'}
                    &nbsp;|&nbsp; <b>3-yr Avg:</b> {sc3}<br>
                    <b>Potential:</b>
                    {'High' if row['pot_tier']==3 else 'Medium' if row['pot_tier']==2 else 'Low'}
                    &nbsp;|&nbsp; <b>Consistency σ:</b>
                    {round(row['std_3yr'],2) if not np.isnan(row['std_3yr']) else '—'}
                  </div>
                  <span class="succ-badge {row['ready_cls']}">{row['ready_icon']} {row['readiness']}</span>
                </div>
                """, unsafe_allow_html=True)
        st.markdown("")

    # full table
    st.markdown('<div class="stitle">Full Succession Table</div>', unsafe_allow_html=True)
    t = succ[["dept","location","name","position","grid_zone","grid_label",
              "overall_score","pot_tier","avg_3yr","std_3yr","readiness"]].copy()
    t.columns = ["Department","Location","Successor Name","Role","Grid Code",
                 "Category","2025 Score","Potential (1-3)",
                 "3-yr Avg Score","Score Stability (σ)","Readiness"]
    for c in ["2025 Score","3-yr Avg Score","Score Stability (σ)"]:
        t[c] = pd.to_numeric(t[c], errors="coerce").round(2)
    st.dataframe(t.reset_index(drop=True), use_container_width=True, height=400)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 5 — EMPLOYEE REGISTER
# ══════════════════════════════════════════════════════════════════════════════
with tab5:
    st.markdown('<div class="stitle">Full Employee Register — All Years</div>',
                unsafe_allow_html=True)

    reg = (df_all[df_all["dept"].isin(sel_depts) & df_all["location"].isin(sel_locs)]
           .pivot_table(
               index=["name","dept","location","position"],
               columns="year",
               values=["grid_zone","overall_score"],
               aggfunc="first")
           .reset_index())
    reg.columns = [f"{b}_{a}" if b else a for a, b in reg.columns]
    disp = reg.rename(columns={
        "name":"Name","dept":"Department","location":"Location","position":"Role",
        "grid_zone_2023":"Grid 2023","grid_zone_2024":"Grid 2024","grid_zone_2025":"Grid 2025",
        "overall_score_2023":"Score 2023","overall_score_2024":"Score 2024",
        "overall_score_2025":"Score 2025",
    })
    for c in ["Score 2023","Score 2024","Score 2025"]:
        if c in disp.columns:
            disp[c] = pd.to_numeric(disp[c], errors="coerce").round(2)
    for c in ["Grid 2023","Grid 2024","Grid 2025"]:
        if c in disp.columns:
            disp[c] = disp[c].fillna("—")

    st.dataframe(disp.reset_index(drop=True), use_container_width=True, height=560)
    st.download_button(
        "⬇️  Download CSV",
        disp.to_csv(index=False).encode("utf-8"),
        "talent_register.csv", "text/csv")

# ─── FOOTER ───────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown(
    "<p style='text-align:center;color:#94A3B8;font-size:10px'>"
    "MBC Media Solutions · Talent 9-Grid Dashboard · "
    "Source: Evalutaion 25/24/23 sheets · Built with Streamlit & Plotly"
    "</p>", unsafe_allow_html=True)
