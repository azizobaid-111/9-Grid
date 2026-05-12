"""
Talent 9-Box Dashboard  —  Streamlit + Plotly
Source: 9Grid_Final.xlsx  (sheet: Overall Evaluation)

Column name map (edit these constants if your file differs):
  COL_EMP_ID      = "EMP ID"
  COL_NAME        = "Name"
  COL_DEPT        = "Depatrment"          ← typo is intentional (matches file)
  COL_LOCATION    = "Employee Location "  ← trailing space is intentional
  COL_JOIN        = "Joining date "
  COL_PERF_25     = "overall 25"
  COL_PERF_24     = "overall 24"
  COL_PERF_23     = "overall 23"
  COL_GRID_25     = "Grid Location"       ← 2025 grid cell (e.g. "3C", "2B")
  COL_GRID_24     = "Grid Location 24"
  COL_GRID_23     = "Grid Location 23"
  COL_PERF_ZONE   = "Perforamnce ZONE 25" ← typo is intentional
  COL_PERF_AVG    = "Performance Average"
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import warnings
warnings.filterwarnings("ignore")

# ─── PAGE CONFIG ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Talent 9-Grid Dashboard",
    page_icon="🎯",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── COLUMN NAME MAP ─────────────────────────────────────────────────────────
# Edit ONLY the right-hand strings if your Excel file uses different names.
COL_EMP_ID    = "EMP ID"
COL_NAME      = "Name"
COL_DEPT      = "Depatrment"           # intentional typo from source file
COL_LOCATION  = "Employee Location "   # trailing space matches source
COL_JOIN      = "Joining date "
COL_PERF_25   = "overall 25"
COL_PERF_24   = "overall 24"
COL_PERF_23   = "overall 23"
COL_GRID_25   = "Grid Location"
COL_GRID_24   = "Grid Location 24"
COL_GRID_23   = "Grid Location 23"
COL_PERF_ZONE = "Perforamnce ZONE 25"
COL_PERF_AVG  = "Performance Average"

# ─── GRID METADATA ───────────────────────────────────────────────────────────
GRID_META = {
    "3C": {"label": "Star",             "perf": 3, "pot": 3, "color": "#1F4E79"},
    "3B": {"label": "High Performer",   "perf": 3, "pot": 2, "color": "#166534"},
    "3A": {"label": "Solid Performer",  "perf": 3, "pot": 1, "color": "#0F766E"},
    "2C": {"label": "High Potential",   "perf": 2, "pot": 3, "color": "#6D28D9"},
    "2B": {"label": "Core Player",      "perf": 2, "pot": 2, "color": "#374151"},
    "2A": {"label": "Avg. Performer",   "perf": 2, "pot": 1, "color": "#B45309"},
    "1C": {"label": "Enigma",           "perf": 1, "pot": 3, "color": "#9D174D"},
    "1B": {"label": "Underperformer",   "perf": 1, "pot": 2, "color": "#9A3412"},
    "1A": {"label": "Risk",             "perf": 1, "pot": 1, "color": "#991B1B"},
}

CELL_BG = {
    "3C": "rgba(31,78,121,0.13)",
    "3B": "rgba(22,101,52,0.13)",
    "3A": "rgba(15,118,110,0.13)",
    "2C": "rgba(109,40,217,0.13)",
    "2B": "rgba(55,65,81,0.10)",
    "2A": "rgba(180,83,9,0.13)",
    "1C": "rgba(157,23,77,0.13)",
    "1B": "rgba(154,52,18,0.13)",
    "1A": "rgba(153,27,27,0.13)",
}

YEAR_COLS = {2023: COL_PERF_23, 2024: COL_PERF_24, 2025: COL_PERF_25}
GRID_COLS = {2023: COL_GRID_23, 2024: COL_GRID_24, 2025: COL_GRID_25}

# ─── CUSTOM CSS ──────────────────────────────────────────────────────────────
st.markdown("""
<style>
/* ── global ── */
[data-testid="stAppViewContainer"] { background: #F0F4F8; }
[data-testid="stSidebar"] { background: #0F172A; }
[data-testid="stSidebar"] * { color: #CBD5E1 !important; }
[data-testid="stSidebar"] .stSelectbox label,
[data-testid="stSidebar"] .stMultiSelect label { color: #94A3B8 !important; font-size:12px !important; }
[data-testid="stSidebar"] h1,
[data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3 { color: #F1F5F9 !important; }

/* ── KPI cards ── */
.kpi-card {
    background: #FFFFFF;
    border-radius: 12px;
    padding: 18px 20px;
    border-left: 4px solid #2E75B6;
    box-shadow: 0 1px 4px rgba(0,0,0,0.07);
    margin-bottom: 8px;
}
.kpi-card.green  { border-left-color: #166534; }
.kpi-card.amber  { border-left-color: #B45309; }
.kpi-card.red    { border-left-color: #991B1B; }
.kpi-card.violet { border-left-color: #6D28D9; }
.kpi-card.teal   { border-left-color: #0F766E; }

.kpi-label { font-size:11px; font-weight:600; color:#64748B;
             text-transform:uppercase; letter-spacing:.06em; margin-bottom:4px; }
.kpi-value { font-size:28px; font-weight:700; color:#0F172A; line-height:1.1; }
.kpi-sub   { font-size:11px; color:#94A3B8; margin-top:3px; }
.kpi-delta { font-size:12px; font-weight:600; }
.kpi-delta.up   { color:#166534; }
.kpi-delta.down { color:#991B1B; }

/* ── section titles ── */
.section-title {
    font-size:15px; font-weight:700; color:#1E3A5F;
    border-left:3px solid #2E75B6; padding-left:10px;
    margin: 24px 0 12px 0;
}
/* ── page header ── */
.page-header {
    background: linear-gradient(135deg, #0F172A 0%, #1E3A5F 100%);
    border-radius:14px; padding:28px 32px; margin-bottom:24px;
    display:flex; align-items:center; justify-content:space-between;
}
.page-header h1 { color:#F1F5F9; font-size:26px; margin:0; font-weight:700; }
.page-header p  { color:#94A3B8; margin:4px 0 0 0; font-size:13px; }
.badge { display:inline-block; background:rgba(46,117,182,0.25);
         color:#93C5FD; border-radius:999px; padding:3px 12px;
         font-size:11px; font-weight:600; }
</style>
""", unsafe_allow_html=True)


# ─── DATA LOADING ─────────────────────────────────────────────────────────────
@st.cache_data(show_spinner="Loading talent data…")
def load_data(path: str) -> pd.DataFrame:
    """
    Load and clean the Overall Evaluation sheet from 9Grid_Final.xlsx.
    Returns a tidy DataFrame with all three years' scores and grid locations.
    """
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb["Overall Evaluation"]
    rows = list(ws.iter_rows(values_only=True))

    # Row index 10 (0-based) holds the headers; data starts at index 11
    raw_headers = list(rows[10])

    # Fix duplicate column names that appear in the wide layout
    seen = {}
    headers = []
    for h in raw_headers:
        h = str(h) if h is not None else "unnamed"
        count = seen.get(h, 0)
        if count:
            headers.append(f"{h}_dup{count}")
        else:
            headers.append(h)
        seen[h] = count + 1

    data = []
    for r in rows[11:]:
        if r[0] is None:
            break
        data.append(r)

    df = pd.DataFrame(data, columns=headers)

    # ── clean score columns ──────────────────────────────────────────────────
    NULL_TOKENS = {"NA", "NR", "-", "", "nan", "None", None}

    def clean_score(col):
        def _parse(v):
            if v in NULL_TOKENS or (isinstance(v, float) and np.isnan(v)):
                return np.nan
            try:
                val = float(v)
                return val if val > 0 else np.nan   # 0 = not rated
            except Exception:
                return np.nan
        return df[col].apply(_parse)

    df[COL_PERF_23] = clean_score(COL_PERF_23)
    df[COL_PERF_24] = clean_score(COL_PERF_24)
    df[COL_PERF_25] = clean_score(COL_PERF_25)

    # ── clean grid location columns ──────────────────────────────────────────
    VALID_CODES = set(GRID_META.keys())

    def clean_grid(col):
        def _parse(v):
            if v in NULL_TOKENS or (isinstance(v, float) and np.isnan(v)):
                return np.nan
            return str(v).strip() if str(v).strip() in VALID_CODES else np.nan
        return df[col].apply(_parse)

    df[COL_GRID_25] = clean_grid(COL_GRID_25)
    df[COL_GRID_24] = clean_grid(COL_GRID_24)
    df[COL_GRID_23] = clean_grid(COL_GRID_23)

    # ── derived columns ──────────────────────────────────────────────────────
    def perf_tier(code):
        try:
            return int(str(code)[0])
        except Exception:
            return np.nan

    def pot_tier(code):
        mapping = {"A": 1, "B": 2, "C": 3}
        try:
            return mapping.get(str(code)[1].upper(), np.nan)
        except Exception:
            return np.nan

    df["perf_tier_25"]  = df[COL_GRID_25].apply(perf_tier)
    df["pot_tier_25"]   = df[COL_GRID_25].apply(pot_tier)
    df["grid_label_25"] = df[COL_GRID_25].map(
        {k: v["label"] for k, v in GRID_META.items()})

    df[COL_NAME]     = df[COL_NAME].astype(str).str.strip()
    df[COL_DEPT]     = df[COL_DEPT].astype(str).str.strip()
    df[COL_LOCATION] = df[COL_LOCATION].astype(str).str.strip()

    # Tenure (years)
    try:
        df["tenure_years"] = (
            pd.Timestamp.now() - pd.to_datetime(df[COL_JOIN], errors="coerce")
        ).dt.days / 365.25
    except Exception:
        df["tenure_years"] = np.nan

    return df


# ─── LOAD ─────────────────────────────────────────────────────────────────────
DATA_PATH = "9Grid_Final.xlsx"

try:
    df_raw = load_data(DATA_PATH)
except FileNotFoundError:
    st.error(f"❌  Excel file not found at `{DATA_PATH}`. "
             "Place `9Grid_Final.xlsx` in the same folder as `app.py`.")
    st.stop()


# ─── SIDEBAR FILTERS ─────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 🎯 Talent Dashboard")
    st.markdown("---")

    # Year selector
    sel_year = st.selectbox(
        "📅 Evaluation Year",
        options=[2025, 2024, 2023],
        index=0,
    )

    perf_col = YEAR_COLS[sel_year]
    grid_col = GRID_COLS[sel_year]

    # Compute pot/perf tiers for selected year on the fly
    def _perf_tier(code):
        try: return int(str(code)[0])
        except: return np.nan

    def _pot_tier(code):
        m = {"A": 1, "B": 2, "C": 3}
        try: return m.get(str(code)[1].upper(), np.nan)
        except: return np.nan

    df_raw["_sel_grid"]  = df_raw[grid_col]
    df_raw["_sel_perf"]  = df_raw["_sel_grid"].apply(_perf_tier)
    df_raw["_sel_pot"]   = df_raw["_sel_grid"].apply(_pot_tier)
    df_raw["_sel_score"] = df_raw[perf_col]
    df_raw["_sel_label"] = df_raw["_sel_grid"].map(
        {k: v["label"] for k, v in GRID_META.items()})

    # Department filter
    depts = sorted(df_raw[COL_DEPT].dropna().unique())
    sel_depts = st.multiselect("🏢 Department", depts, default=depts)

    # Location filter
    locs = sorted(df_raw[COL_LOCATION].dropna().unique())
    sel_locs = st.multiselect("📍 Location", locs, default=locs)

    # Employee name search
    names = sorted(df_raw[COL_NAME].dropna().unique())
    sel_names = st.multiselect("👤 Employee (optional)", names, default=[])

    # Grid category filter
    categories = sorted(
        [v["label"] for v in GRID_META.values()])
    sel_cats = st.multiselect(
        "📊 Grid Category", categories, default=categories)

    st.markdown("---")
    st.caption("Data source: 9Grid_Final.xlsx")

# ─── FILTER DATAFRAME ─────────────────────────────────────────────────────────
mask = (
    df_raw[COL_DEPT].isin(sel_depts) &
    df_raw[COL_LOCATION].isin(sel_locs) &
    df_raw["_sel_grid"].notna()
)
if sel_names:
    mask &= df_raw[COL_NAME].isin(sel_names)
if sel_cats:
    mask &= df_raw["_sel_label"].isin(sel_cats)

df = df_raw[mask].copy()
df_all = df_raw[
    df_raw[COL_DEPT].isin(sel_depts) &
    df_raw[COL_LOCATION].isin(sel_locs) &
    df_raw["_sel_grid"].notna()
].copy()  # unfiltered by name/category — for trend charts


# ─── HEADER ──────────────────────────────────────────────────────────────────
st.markdown(f"""
<div class="page-header">
  <div>
    <h1>🎯 Talent 9-Grid Dashboard</h1>
    <p>Executive Talent Analytics · Year: <strong style="color:#93C5FD">{sel_year}</strong>
       · Filtered: <strong style="color:#93C5FD">{len(df)}</strong> employees</p>
  </div>
  <div>
    <span class="badge">9-Box Framework</span>
  </div>
</div>
""", unsafe_allow_html=True)


# ─── KPI CARDS ────────────────────────────────────────────────────────────────
def kpi_html(label, value, sub="", color_cls="", delta_txt="", delta_dir=""):
    delta_html = ""
    if delta_txt:
        arrow = "▲" if delta_dir == "up" else "▼"
        delta_html = f'<div class="kpi-delta {delta_dir}">{arrow} {delta_txt}</div>'
    return f"""
    <div class="kpi-card {color_cls}">
      <div class="kpi-label">{label}</div>
      <div class="kpi-value">{value}</div>
      <div class="kpi-sub">{sub}</div>
      {delta_html}
    </div>"""

total_emp      = len(df)
stars          = int((df["_sel_grid"] == "3C").sum())
high_pot       = int((df["_sel_pot"]  == 3).sum())
at_risk        = int((df["_sel_grid"] == "1A").sum())
avg_perf       = df["_sel_score"].dropna().mean()
avg_pot        = df["_sel_pot"].dropna().mean()
top_dept       = (df.groupby(COL_DEPT)["_sel_score"]
                    .mean().dropna().idxmax()
                 if total_emp > 0 else "—")
high_pot_pct   = high_pot / total_emp * 100 if total_emp else 0
star_pct       = stars / total_emp * 100 if total_emp else 0

c1, c2, c3, c4, c5, c6 = st.columns(6)
with c1:
    st.markdown(kpi_html("Total Employees", total_emp,
                          f"{sel_year} evaluation cycle"), unsafe_allow_html=True)
with c2:
    st.markdown(kpi_html("Stars (3C)", stars,
                          f"{star_pct:.1f}% of workforce",
                          "green"), unsafe_allow_html=True)
with c3:
    st.markdown(kpi_html("High Potentials", high_pot,
                          f"{high_pot_pct:.1f}% density",
                          "violet"), unsafe_allow_html=True)
with c4:
    st.markdown(kpi_html("Avg Performance", f"{avg_perf:.2f}" if not np.isnan(avg_perf) else "—",
                          "Score out of 5.0",
                          "teal"), unsafe_allow_html=True)
with c5:
    st.markdown(kpi_html("Top Department", top_dept,
                          "by avg performance score",
                          ""), unsafe_allow_html=True)
with c6:
    st.markdown(kpi_html("At Risk (1A)", at_risk,
                          "Needs immediate action",
                          "red"), unsafe_allow_html=True)


# ─── TAB LAYOUT ──────────────────────────────────────────────────────────────
tab1, tab2, tab3, tab4 = st.tabs([
    "🔲 9-Box Matrix",
    "📈 YoY Trends",
    "🏢 Department Analysis",
    "📋 Employee Detail",
])


# ══════════════════════════════════════════════════════════════════════════════
# TAB 1  —  9-BOX MATRIX
# ══════════════════════════════════════════════════════════════════════════════
with tab1:
    col_matrix, col_dist = st.columns([3, 1])

    with col_matrix:
        st.markdown('<div class="section-title">Talent 9-Box Matrix</div>',
                    unsafe_allow_html=True)

        grid_df = df[df["_sel_grid"].notna() &
                     df["_sel_perf"].notna() &
                     df["_sel_pot"].notna()].copy()

        rng = np.random.default_rng(seed=42)
        jitter = 0.22
        grid_df["jx"] = grid_df["_sel_perf"].astype(float) + rng.uniform(
            -jitter, jitter, size=len(grid_df))
        grid_df["jy"] = grid_df["_sel_pot"].astype(float) + rng.uniform(
            -jitter, jitter, size=len(grid_df))

        fig = go.Figure()

        # ── cell backgrounds ─────────────────────────────────────────────────
        for code, meta in GRID_META.items():
            px_c = meta["perf"]
            py_c = meta["pot"]
            fig.add_shape(
                type="rect",
                x0=px_c - 0.5, x1=px_c + 0.5,
                y0=py_c - 0.5, y1=py_c + 0.5,
                fillcolor=CELL_BG[code],
                line=dict(color="rgba(200,210,220,0.4)", width=1),
                layer="below",
            )
            # Cell label (top-left)
            fig.add_annotation(
                x=px_c - 0.42, y=py_c + 0.40,
                text=f"<b>{meta['label']}</b>",
                showarrow=False,
                font=dict(size=9, color=meta["color"]),
                xanchor="left", yanchor="top",
                opacity=0.85,
            )
            # Code badge (top-right)
            fig.add_annotation(
                x=px_c + 0.42, y=py_c + 0.40,
                text=f"<b>{code}</b>",
                showarrow=False,
                font=dict(size=8, color=meta["color"]),
                xanchor="right", yanchor="top",
                opacity=0.60,
            )

        # ── scatter dots ─────────────────────────────────────────────────────
        for code, meta in GRID_META.items():
            subset = grid_df[grid_df["_sel_grid"] == code]
            if subset.empty:
                continue
            first_name = subset[COL_NAME].str.split().str[0]
            hover_text = (
                "<b>" + subset[COL_NAME] + "</b><br>" +
                "Dept: " + subset[COL_DEPT] + "<br>" +
                "Location: " + subset[COL_LOCATION] + "<br>" +
                "Grid: " + code + " (" + meta["label"] + ")<br>" +
                "Perf score: " + subset["_sel_score"].round(2).astype(str)
            )
            fig.add_trace(go.Scatter(
                x=subset["jx"], y=subset["jy"],
                mode="markers+text",
                name=meta["label"],
                marker=dict(
                    color=meta["color"], size=9,
                    line=dict(width=1, color="white"),
                    opacity=0.88,
                ),
                text=first_name,
                textposition="top center",
                textfont=dict(size=7, color="#1E293B"),
                hovertext=hover_text,
                hoverinfo="text",
                showlegend=True,
            ))

        # ── axis labels ───────────────────────────────────────────────────────
        for val, lbl in [(1, "Poor"), (2, "Average"), (3, "High")]:
            fig.add_annotation(x=val, y=0.45,
                text=f"<b>{lbl}</b>", showarrow=False,
                font=dict(size=9, color="#64748B"), yanchor="top")
            fig.add_annotation(x=0.45, y=val,
                text=f"<b>{lbl}</b>", showarrow=False,
                font=dict(size=9, color="#64748B"), xanchor="right")

        fig.update_layout(
            height=560,
            xaxis=dict(
                range=[0.45, 3.55], tickvals=[1, 2, 3],
                ticktext=["", "", ""], showgrid=False,
                zeroline=False, title="Performance →",
                title_font=dict(size=11, color="#475569"),
            ),
            yaxis=dict(
                range=[0.45, 3.55], tickvals=[1, 2, 3],
                ticktext=["", "", ""], showgrid=False,
                zeroline=False, title="← Potential",
                title_font=dict(size=11, color="#475569"),
            ),
            plot_bgcolor="#FAFBFC",
            paper_bgcolor="#FFFFFF",
            margin=dict(l=60, r=20, t=20, b=60),
            legend=dict(
                orientation="h", yanchor="bottom", y=-0.18,
                xanchor="center", x=0.5,
                font=dict(size=9), itemsizing="constant",
            ),
            font=dict(family="Inter, Arial, sans-serif"),
        )
        # Grid lines at 1.5 and 2.5
        for v in [1.5, 2.5]:
            fig.add_vline(x=v, line=dict(color="rgba(148,163,184,0.5)", width=1, dash="dot"))
            fig.add_hline(y=v, line=dict(color="rgba(148,163,184,0.5)", width=1, dash="dot"))

        st.plotly_chart(fig, use_container_width=True)

    with col_dist:
        st.markdown('<div class="section-title">Category Breakdown</div>',
                    unsafe_allow_html=True)

        cat_counts = (grid_df.groupby("_sel_label")
                              .size().reset_index(name="count")
                              .sort_values("count", ascending=True))
        colors_bar = [
            GRID_META.get(
                next((k for k, v in GRID_META.items()
                      if v["label"] == lbl), None), {}).get("color", "#64748B")
            for lbl in cat_counts["_sel_label"]
        ]
        fig_bar = go.Figure(go.Bar(
            x=cat_counts["count"], y=cat_counts["_sel_label"],
            orientation="h",
            marker_color=colors_bar,
            text=cat_counts["count"],
            textposition="outside",
            textfont=dict(size=10),
        ))
        fig_bar.update_layout(
            height=560, margin=dict(l=10, r=30, t=10, b=30),
            plot_bgcolor="#FFFFFF", paper_bgcolor="#FFFFFF",
            xaxis=dict(showgrid=True, gridcolor="#F1F5F9", zeroline=False),
            yaxis=dict(showgrid=False),
            font=dict(size=10, family="Inter, Arial"),
        )
        st.plotly_chart(fig_bar, use_container_width=True)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 2  —  YoY TRENDS
# ══════════════════════════════════════════════════════════════════════════════
with tab2:
    st.markdown('<div class="section-title">Year-over-Year Movement Analysis</div>',
                unsafe_allow_html=True)

    # Build long-format for all 3 years (unfiltered by name/category)
    frames = []
    for yr, p_col, g_col in [
        (2023, COL_PERF_23, COL_GRID_23),
        (2024, COL_PERF_24, COL_GRID_24),
        (2025, COL_PERF_25, COL_GRID_25),
    ]:
        tmp = df_all[[COL_NAME, COL_DEPT, COL_LOCATION, p_col, g_col]].copy()
        tmp.columns = [COL_NAME, COL_DEPT, COL_LOCATION, "perf_score", "grid"]
        tmp["year"] = yr
        tmp = tmp[tmp["grid"].notna() & tmp["grid"].str.match(r"^[123][ABC]$", na=False)]
        tmp["perf_tier"] = tmp["grid"].apply(_perf_tier)
        tmp["pot_tier"]  = tmp["grid"].apply(_pot_tier)
        tmp["label"]     = tmp["grid"].map({k: v["label"] for k, v in GRID_META.items()})
        frames.append(tmp)

    df_long = pd.concat(frames, ignore_index=True)

    # Stars & high-potentials per year
    trend_df = (df_long.groupby("year")
                .apply(lambda g: pd.Series({
                    "Total":          len(g),
                    "Stars":          (g["grid"] == "3C").sum(),
                    "High Potentials":(g["pot_tier"] == 3).sum(),
                    "At Risk":        (g["grid"] == "1A").sum(),
                    "Avg Perf":       g["perf_score"].mean(),
                }))
                .reset_index())

    c_trend1, c_trend2 = st.columns(2)

    with c_trend1:
        fig_trend = go.Figure()
        fig_trend.add_trace(go.Scatter(
            x=trend_df["year"], y=trend_df["Stars"],
            name="Stars (3C)", mode="lines+markers+text",
            line=dict(color="#1F4E79", width=2.5),
            marker=dict(size=8, color="#1F4E79"),
            text=trend_df["Stars"].astype(int),
            textposition="top center",
        ))
        fig_trend.add_trace(go.Scatter(
            x=trend_df["year"], y=trend_df["High Potentials"],
            name="High Potentials", mode="lines+markers+text",
            line=dict(color="#6D28D9", width=2.5, dash="dash"),
            marker=dict(size=8, color="#6D28D9"),
            text=trend_df["High Potentials"].astype(int),
            textposition="top center",
        ))
        fig_trend.add_trace(go.Scatter(
            x=trend_df["year"], y=trend_df["At Risk"],
            name="At Risk (1A)", mode="lines+markers+text",
            line=dict(color="#991B1B", width=2.5, dash="dot"),
            marker=dict(size=8, color="#991B1B"),
            text=trend_df["At Risk"].astype(int),
            textposition="bottom center",
        ))
        fig_trend.update_layout(
            title="Talent Tier Movement 2023 → 2025",
            title_font=dict(size=13, color="#1E293B"),
            height=320, plot_bgcolor="#FAFBFC", paper_bgcolor="#FFFFFF",
            xaxis=dict(tickvals=[2023, 2024, 2025], showgrid=False),
            yaxis=dict(showgrid=True, gridcolor="#F1F5F9"),
            legend=dict(orientation="h", y=-0.2),
            margin=dict(l=40, r=20, t=40, b=60),
            font=dict(family="Inter, Arial"),
        )
        st.plotly_chart(fig_trend, use_container_width=True)

    with c_trend2:
        fig_avg = go.Figure()
        fig_avg.add_trace(go.Bar(
            x=trend_df["year"].astype(str),
            y=trend_df["Avg Perf"].round(2),
            marker_color=["#93C5FD", "#3B82F6", "#1F4E79"],
            text=trend_df["Avg Perf"].round(2),
            textposition="outside",
        ))
        fig_avg.update_layout(
            title="Average Performance Score by Year",
            title_font=dict(size=13, color="#1E293B"),
            height=320, plot_bgcolor="#FAFBFC", paper_bgcolor="#FFFFFF",
            xaxis=dict(showgrid=False),
            yaxis=dict(range=[0, 5.3], showgrid=True, gridcolor="#F1F5F9"),
            margin=dict(l=40, r=20, t=40, b=40),
            font=dict(family="Inter, Arial"),
            showlegend=False,
        )
        st.plotly_chart(fig_avg, use_container_width=True)

    # YoY movement table per employee
    st.markdown('<div class="section-title">Employee Grid Movement (2023 → 2025)</div>',
                unsafe_allow_html=True)

    movement = df_all[[COL_NAME, COL_DEPT, COL_LOCATION,
                        COL_GRID_23, COL_GRID_24, COL_GRID_25]].copy()
    movement.columns = [COL_NAME, COL_DEPT, COL_LOCATION, "Grid 2023", "Grid 2024", "Grid 2025"]
    movement = movement[movement["Grid 2025"].notna()]

    def movement_indicator(g_from, g_to):
        if pd.isna(g_from) or pd.isna(g_to):
            return "🆕 New"
        s_from = GRID_META.get(g_from, {}).get("perf", 0) + GRID_META.get(g_from, {}).get("pot", 0)
        s_to   = GRID_META.get(g_to, {}).get("perf", 0)   + GRID_META.get(g_to, {}).get("pot", 0)
        if s_to > s_from:   return "⬆️ Improved"
        if s_to == s_from:  return "➡️ Stable"
        return "⬇️ Declined"

    movement["YoY 23→25"] = movement.apply(
        lambda r: movement_indicator(r["Grid 2023"], r["Grid 2025"]), axis=1)

    movement["Grid 2023"] = movement["Grid 2023"].fillna("N/A")
    movement["Grid 2024"] = movement["Grid 2024"].fillna("N/A")
    movement["Grid 2025"] = movement["Grid 2025"].fillna("N/A")

    st.dataframe(
        movement.reset_index(drop=True),
        use_container_width=True,
        height=320,
        column_config={
            COL_NAME:     st.column_config.TextColumn("Employee Name", width="large"),
            COL_DEPT:     st.column_config.TextColumn("Department"),
            COL_LOCATION: st.column_config.TextColumn("Location"),
            "Grid 2023":  st.column_config.TextColumn("2023"),
            "Grid 2024":  st.column_config.TextColumn("2024"),
            "Grid 2025":  st.column_config.TextColumn("2025"),
            "YoY 23→25":  st.column_config.TextColumn("Movement"),
        }
    )


# ══════════════════════════════════════════════════════════════════════════════
# TAB 3  —  DEPARTMENT ANALYSIS
# ══════════════════════════════════════════════════════════════════════════════
with tab3:
    st.markdown('<div class="section-title">Department & Location Intelligence</div>',
                unsafe_allow_html=True)

    c_d1, c_d2 = st.columns(2)

    with c_d1:
        # Dept breakdown by grid category (stacked bar)
        dept_grid = (df.groupby([COL_DEPT, "_sel_label"])
                       .size().reset_index(name="count"))
        category_order = [v["label"] for v in GRID_META.values()]
        color_map = {v["label"]: v["color"] for v in GRID_META.values()}

        fig_dept = px.bar(
            dept_grid, x=COL_DEPT, y="count", color="_sel_label",
            color_discrete_map=color_map,
            category_orders={"_sel_label": category_order},
            title="Employee Distribution by Department & Grid Category",
            labels={COL_DEPT: "Department", "count": "Employees", "_sel_label": "Category"},
        )
        fig_dept.update_layout(
            height=380, plot_bgcolor="#FAFBFC", paper_bgcolor="#FFFFFF",
            xaxis_tickangle=-35,
            legend=dict(orientation="h", y=-0.30, font=dict(size=9)),
            margin=dict(l=40, r=20, t=50, b=120),
            font=dict(family="Inter, Arial", size=10),
            title_font=dict(size=13, color="#1E293B"),
        )
        st.plotly_chart(fig_dept, use_container_width=True)

    with c_d2:
        # Avg performance by dept (horizontal bar)
        dept_avg = (df.groupby(COL_DEPT)["_sel_score"]
                      .mean().dropna().sort_values(ascending=True)
                      .reset_index())
        dept_avg.columns = ["Department", "Avg Score"]
        bar_colors = ["#1F4E79" if v >= dept_avg["Avg Score"].quantile(0.67)
                      else "#3B82F6" if v >= dept_avg["Avg Score"].quantile(0.33)
                      else "#93C5FD"
                      for v in dept_avg["Avg Score"]]
        fig_avg_dept = go.Figure(go.Bar(
            x=dept_avg["Avg Score"], y=dept_avg["Department"],
            orientation="h",
            marker_color=bar_colors,
            text=dept_avg["Avg Score"].round(2),
            textposition="outside",
        ))
        fig_avg_dept.update_layout(
            title="Average Performance Score by Department",
            title_font=dict(size=13, color="#1E293B"),
            height=380, plot_bgcolor="#FAFBFC", paper_bgcolor="#FFFFFF",
            xaxis=dict(range=[0, 5.5], showgrid=True, gridcolor="#F1F5F9"),
            yaxis=dict(showgrid=False),
            margin=dict(l=10, r=60, t=50, b=40),
            font=dict(family="Inter, Arial", size=10),
            showlegend=False,
        )
        st.plotly_chart(fig_avg_dept, use_container_width=True)

    c_d3, c_d4 = st.columns(2)

    with c_d3:
        # Location distribution (donut)
        loc_counts = df.groupby(COL_LOCATION).size().reset_index(name="count")
        loc_colors = ["#1F4E79", "#2E75B6", "#166534", "#B45309"]
        fig_loc = go.Figure(go.Pie(
            labels=loc_counts[COL_LOCATION], values=loc_counts["count"],
            hole=0.55,
            marker_colors=loc_colors[:len(loc_counts)],
            textinfo="label+percent",
            textfont=dict(size=10),
        ))
        fig_loc.add_annotation(
            text=f"<b>{len(df)}</b><br>Employees",
            x=0.5, y=0.5, font_size=14, showarrow=False,
            font=dict(color="#1E293B"),
        )
        fig_loc.update_layout(
            title="Talent Distribution by Location",
            title_font=dict(size=13, color="#1E293B"),
            height=320, paper_bgcolor="#FFFFFF",
            margin=dict(l=20, r=20, t=50, b=20),
            showlegend=True,
            legend=dict(orientation="h", y=-0.1),
            font=dict(family="Inter, Arial"),
        )
        st.plotly_chart(fig_loc, use_container_width=True)

    with c_d4:
        # High-potential % by department
        dept_hp = (df.groupby(COL_DEPT)
                     .apply(lambda g: pd.Series({
                         "total":    len(g),
                         "high_pot": (g["_sel_pot"] == 3).sum(),
                     }))
                     .reset_index())
        dept_hp["hp_pct"] = dept_hp["high_pot"] / dept_hp["total"] * 100
        dept_hp = dept_hp.sort_values("hp_pct", ascending=True)

        fig_hp = go.Figure(go.Bar(
            x=dept_hp["hp_pct"], y=dept_hp[COL_DEPT],
            orientation="h",
            marker_color="#6D28D9",
            text=dept_hp["hp_pct"].round(1).astype(str) + "%",
            textposition="outside",
        ))
        fig_hp.update_layout(
            title="High Potential % by Department",
            title_font=dict(size=13, color="#1E293B"),
            height=320, plot_bgcolor="#FAFBFC", paper_bgcolor="#FFFFFF",
            xaxis=dict(range=[0, 110], showgrid=True, gridcolor="#F1F5F9",
                       ticksuffix="%"),
            yaxis=dict(showgrid=False),
            margin=dict(l=10, r=70, t=50, b=40),
            font=dict(family="Inter, Arial", size=10),
            showlegend=False,
        )
        st.plotly_chart(fig_hp, use_container_width=True)

    # Performance vs Potential scatter
    st.markdown('<div class="section-title">Performance Score vs Potential Tier</div>',
                unsafe_allow_html=True)

    scatter_df = df[df["_sel_score"].notna() & df["_sel_pot"].notna()].copy()
    scatter_df["pot_label"] = scatter_df["_sel_pot"].map(
        {1: "Low", 2: "Medium", 3: "High"})

    fig_scatter = px.scatter(
        scatter_df,
        x="_sel_score", y="_sel_pot",
        color=COL_DEPT,
        hover_name=COL_NAME,
        hover_data={COL_DEPT: True, COL_LOCATION: True,
                    "_sel_score": ":.2f", "_sel_pot": False},
        size_max=10,
        labels={"_sel_score": "Performance Score (0–5)",
                "_sel_pot": "Potential Tier",
                COL_DEPT: "Department"},
        title=f"Performance Score vs Potential Tier ({sel_year})",
    )
    fig_scatter.update_traces(marker=dict(size=9, opacity=0.75,
                                          line=dict(width=1, color="white")))
    fig_scatter.update_layout(
        height=380, plot_bgcolor="#FAFBFC", paper_bgcolor="#FFFFFF",
        yaxis=dict(tickvals=[1, 2, 3],
                   ticktext=["Low (1)", "Medium (2)", "High (3)"],
                   showgrid=True, gridcolor="#F1F5F9"),
        xaxis=dict(range=[0, 5.3], showgrid=True, gridcolor="#F1F5F9"),
        margin=dict(l=60, r=20, t=50, b=40),
        title_font=dict(size=13, color="#1E293B"),
        font=dict(family="Inter, Arial", size=10),
        legend=dict(font=dict(size=9)),
    )
    st.plotly_chart(fig_scatter, use_container_width=True)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 4  —  EMPLOYEE DETAIL TABLE
# ══════════════════════════════════════════════════════════════════════════════
with tab4:
    st.markdown('<div class="section-title">Employee Detail — Full Register</div>',
                unsafe_allow_html=True)

    display_cols = [
        COL_EMP_ID, COL_NAME, COL_DEPT, COL_LOCATION,
        COL_GRID_23, COL_GRID_24, COL_GRID_25,
        "_sel_score", "_sel_label",
    ]
    display_df = df[display_cols].copy()
    display_df.columns = [
        "EMP ID", "Name", "Department", "Location",
        "Grid 2023", "Grid 2024", "Grid 2025",
        f"Score {sel_year}", f"Category {sel_year}",
    ]
    display_df[[f"Score {sel_year}"]] = display_df[[f"Score {sel_year}"]].round(2)
    display_df = display_df.fillna("—").reset_index(drop=True)

    st.dataframe(
        display_df,
        use_container_width=True,
        height=520,
        column_config={
            "EMP ID":     st.column_config.NumberColumn("EMP ID", width="small"),
            "Name":       st.column_config.TextColumn("Employee Name", width="large"),
            "Department": st.column_config.TextColumn("Department"),
            "Location":   st.column_config.TextColumn("Location"),
            "Grid 2023":  st.column_config.TextColumn("2023", width="small"),
            "Grid 2024":  st.column_config.TextColumn("2024", width="small"),
            "Grid 2025":  st.column_config.TextColumn("2025", width="small"),
            f"Score {sel_year}": st.column_config.NumberColumn(
                f"Score {sel_year}", format="%.2f", width="small"),
            f"Category {sel_year}": st.column_config.TextColumn(
                f"Category {sel_year}"),
        }
    )

    # Download button
    csv = display_df.to_csv(index=False).encode("utf-8")
    st.download_button(
        label="⬇️  Download as CSV",
        data=csv,
        file_name=f"talent_data_{sel_year}.csv",
        mime="text/csv",
    )

# ─── FOOTER ──────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown(
    "<p style='text-align:center;color:#94A3B8;font-size:11px;'>"
    "Talent 9-Grid Dashboard · Data source: 9Grid_Final.xlsx · "
    "Built with Streamlit & Plotly"
    "</p>",
    unsafe_allow_html=True
)
