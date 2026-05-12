# Talent 9-Grid Dashboard
### Streamlit · Plotly · Executive HR Analytics

---

## Folder structure

```
talent_dashboard/
├── app.py                    ← main Streamlit application
├── 9Grid_Final.xlsx          ← your Excel data file  ← MUST BE HERE
├── requirements.txt
├── .streamlit/
│   └── config.toml           ← theme + server settings
└── README.md
```

> **Important:** `9Grid_Final.xlsx` must sit in the same folder as `app.py`.

---

## Field mapping confirmed from your file

| Dashboard field | Excel column | Notes |
|---|---|---|
| Employee Name | `Name` | Exact match, trimmed |
| Year | Derived from columns | `overall 23`, `overall 24`, `overall 25` |
| Department | `Depatrment` | Intentional typo in source file |
| Location | `Employee Location ` | Trailing space is part of the column name |
| Performance Score | `overall 25` / `overall 24` / `overall 23` | Float 0–5; NA/NR/0 → excluded |
| Potential Score | Derived from `Grid Location` | Letter: A=Low, B=Medium, C=High |
| Grid Category | `Grid Location` (2025), `Grid Location 24`, `Grid Location 23` | e.g. "3C" = Star |

### Data quality notes (from audit)
- `Depatrment` and `Perforamnce ZONE 25` contain typos — code uses exact spellings
- NA / NR scores are treated as "not evaluated" and excluded
- Score = 0 is treated as "not rated" and excluded
- No standalone Potential column exists; potential is inferred from the grid code letter

---

## Run locally

```bash
# 1. Clone or copy this folder to your machine
cd talent_dashboard

# 2. Create a virtual environment (recommended)
python -m venv venv
source venv/bin/activate        # Mac/Linux
venv\Scripts\activate           # Windows

# 3. Install dependencies
pip install -r requirements.txt

# 4. Run the app
streamlit run app.py
```

The dashboard opens at **http://localhost:8501**

---

## Deploy to Streamlit Cloud (public URL — free)

### Step 1 — Push to GitHub

```bash
# initialise git in the project folder
git init
git add .
git commit -m "initial commit: talent dashboard"

# create a repo on github.com (e.g. talent-9grid-dashboard)
# then push:
git remote add origin https://github.com/YOUR_USERNAME/talent-9grid-dashboard.git
git branch -M main
git push -u origin main
```

### Step 2 — Connect to Streamlit Cloud

1. Go to **https://share.streamlit.io** → sign in with GitHub
2. Click **"New app"**
3. Select your repository: `YOUR_USERNAME/talent-9grid-dashboard`
4. Branch: `main`
5. Main file path: `app.py`
6. Click **"Deploy!"**

### Step 3 — Add the Excel file

Because `9Grid_Final.xlsx` contains private employee data, **do not commit it to a public repo**.

**Option A — Private repo (recommended)**
Make your GitHub repo **private**. Streamlit Cloud still deploys it.
Push the file: `git add 9Grid_Final.xlsx && git commit -m "add data" && git push`

**Option B — Streamlit secrets (advanced)**
Convert the Excel to base64, store in `st.secrets`, and load it in `app.py`. Ask if you need this approach.

### Step 4 — Get your public URL

After deployment (takes ~60 seconds), Streamlit Cloud gives you a URL:
```
https://YOUR_USERNAME-talent-9grid-dashboard-app-XXXXXX.streamlit.app
```
Share this link with anyone. No Power BI license needed.

---

## Customisation guide

### Rename a column
Edit the constant at the top of `app.py` (section labelled `COLUMN NAME MAP`):
```python
COL_DEPT = "Depatrment"   # ← change this string to match your column
```

### Change the performance score thresholds
The grid cell (e.g. "3C") is already pre-computed in the Excel file.
If you want to recompute from raw scores, edit the `_perf_tier` lambda and
the `_pot_tier` lambda inside `app.py`.

### Add more years
Add entries to `YEAR_COLS` and `GRID_COLS` dictionaries, and add
corresponding columns to the Excel file.

---

## Architecture

```
9Grid_Final.xlsx
        │
        ▼
  load_data()          ← openpyxl, pandas, cached with @st.cache_data
        │
        ▼
  Sidebar filters      ← Year, Department, Location, Name, Category
        │
        ▼
  Filtered DataFrame   ── Tab 1: 9-Box Matrix (Plotly scatter)
                        ── Tab 2: YoY Trends (line + bar)
                        ── Tab 3: Dept / Location analysis (bar, donut, scatter)
                        ── Tab 4: Full employee table + CSV export
```
