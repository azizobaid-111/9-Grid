# MBC Media Solutions — Talent 9-Grid Dashboard  v3

**Streamlit · Plotly · Executive HR Analytics**  
Brand colours: `#666EFF` Purple · `#30BFA6` Green · `#EDEDEE` Light Gray · `#2B2B2B` Dark Gray

---

## Folder structure

```
talent_dashboard/
├── app.py                  ← main Streamlit app  (this file)
├── 9Grid_Final.xlsx        ← Excel source  ← MUST be here
├── logo.png                ← MBC Media Solutions logo
├── requirements.txt
├── .streamlit/
│   └── config.toml         ← brand theme
└── README.md
```

---

## Data source & field mapping

| Dashboard field | Excel sheet | Exact column | Notes |
|---|---|---|---|
| Employee Name | Evalutaion 23/24/25 | `Employee Name` | Exact match, trimmed |
| Year | Sheet name | 2023 / 2024 / 2025 | Derived |
| Department | All sheets | `Department` | Trimmed (has trailing spaces in file) |
| Location | All sheets | `Location` | MMS KSA→Saudi Arabia, MMS UAE→UAE, MMS Egypt→Egypt |
| Position / Role | All sheets | `Position` | Exact match |
| Performance Score | All sheets | `Objectives and Competency Score` | Float 0–5; NA/NR/0 → excluded |
| Potential | All sheets | `Potential zone` | High/Medium/Low → 3/2/1 |
| Grid Category | All sheets | `Grid Zone ` / `Grid Zone` / `Gird Zone` | e.g. "3C" = Star |

### Validated counts (from Evalutaion sheets — exact source of truth)

| Code | Label | 2023 | 2024 | 2025 |
|---|---|---:|---:|---:|
| 3C | Star | 23 | 36 | **56** |
| 3B | High Performer | 3 | 1 | 3 |
| 3A | Solid Performer | 3 | 0 | 0 |
| 2C | Growth Employee | 5 | 2 | 11 |
| 2B | Core Player | 52 | 84 | 97 |
| 2A | Average Performer | 30 | 16 | 21 |
| 1C | Potential Gem | 0 | 0 | 0 |
| 1B | Inconsistent Player | 0 | 0 | 0 |
| 1A | Risk | 2 | 3 | 1 |
| **Total rated** | | **118** | **142** | **189** |

---

## Features

| Tab | Contents |
|---|---|
| 🔲 9-Box Matrix | Interactive Plotly grid with jittered names, employee count per cell, expandable name lists |
| 🏆 Top Performers | Ranked cards (3C → 3B → score) with name, role, dept, location, score, category |
| 📈 Year Comparison | Bell curve, talent movement line chart, stacked bar, avg score bar, individual movement table |
| 📋 Employee Register | Full pivoted table (all 3 years) + CSV download |

**Filters:** Year · Department (All + individual) · Location (All + individual) · Employee Name (All + individual)  
All filters update KPI cards, 9-box, top performers, charts, and register simultaneously.

**Removed:** Succession Planning section (separate project) · Role/Position filter

---

## Run locally

```bash
cd talent_dashboard

# create virtual env (recommended)
python -m venv venv
source venv/bin/activate      # Mac/Linux
venv\Scripts\activate         # Windows

pip install -r requirements.txt
streamlit run app.py
# Opens at http://localhost:8501
```

---

## Deploy to Streamlit Cloud (free public URL)

1. Push this folder to a **private** GitHub repository  
   ```bash
   git init && git add . && git commit -m "MMS Talent Dashboard v3"
   git remote add origin https://github.com/YOUR_USER/mms-talent-dashboard.git
   git branch -M main && git push -u origin main
   ```

2. Go to **https://share.streamlit.io** → sign in with GitHub → **New app**

3. Select repo, branch `main`, file `app.py` → **Deploy**

4. Your public URL will be:  
   `https://YOUR_USER-mms-talent-dashboard-app-XXXX.streamlit.app`

---

## Files to replace in GitHub (v2 → v3)

| File | Action |
|---|---|
| `app.py` | **Replace** — full rewrite with brand colours, new layout, bell curve, top performers |
| `requirements.txt` | **Replace** — same 5 packages, updated pin |
| `.streamlit/config.toml` | **Replace** — brand purple `#666EFF` primary colour |
| `README.md` | **Replace** — updated docs |
| `logo.png` | **Keep** — unchanged |
| `9Grid_Final.xlsx` | **Keep** — unchanged |
