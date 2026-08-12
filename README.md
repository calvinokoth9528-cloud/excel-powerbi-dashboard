# 📊 Excel → Power BI Analytics Dashboard

**An interactive, self-contained HTML dashboard built from two real-world Excel
datasets — the Microsoft Power BI Financial Sample and a chocolate company's
shipment ledger.** Built during the data-analytics attachment at **KEMRI**.

[![HTML](https://img.shields.io/badge/Dashboard-HTML-orange?logo=html5&logoColor=white&labelColor=gray)](dashboard.html)
[![Python](https://img.shields.io/badge/Python-3.x-3776AB?logo=python&logoColor=white&labelColor=gray)](build_dashboard.py)
[![Chart.js](https://img.shields.io/badge/Chart.js-4.x-FF6384?logo=chartdotjs&logoColor=white&labelColor=gray)](https://www.chartjs.org/)
[![Excel](https://img.shields.io/badge/Excel-2%20workbooks-217346?logo=microsoftexcel&logoColor=white&labelColor=gray)](financial.xlsx)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Status: Complete](https://img.shields.io/badge/Status-Complete-brightgreen)](dashboard.html)

**Author:** Calvin Okoth · Data Analytics Attachment, Kenya Medical Research Institute (KEMRI)

| Quick links | |
|---|---|
| 🖥️ Live dashboard | [`dashboard.html`](dashboard.html) — double-click to open (works offline) |
| 🌐 Hosted demo | `https://<your-username>.github.io/excel-powerbi-dashboard/dashboard.html` — once Pages is enabled (see [Live Demo](#live-demo)) |
| ⚙️ Generator | [`build_dashboard.py`](build_dashboard.py) |
| 🗂️ Data · Power BI sample | [`financial.xlsx`](financial.xlsx) |
| 🗂️ Data · Shipments | [`ac-sample-data.xlsx`](ac-sample-data.xlsx) |
| ⚖️ License | [`LICENSE`](LICENSE) |

---

## Table of Contents

- [What This Is](#what-this-is)
- [Live Demo](#live-demo)
- [Preview](#preview)
- [The Two Datasets](#the-two-datasets)
- [Dashboard Features](#dashboard-features)
- [Key Insights](#key-insights)
- [How to Run](#how-to-run)
- [How It's Built](#how-its-built)
- [Repository Structure](#repository-structure)
- [Skills Demonstrated](#skills-demonstrated)
- [Assessment Talking Points](#assessment-talking-points)
- [License & Data Provenance](#license--data-provenance)
- [Citation](#citation)

---

## What This Is

This project turns two everyday Excel workbooks into a polished, interactive
dashboard — the kind of deliverable produced in a real business-intelligence
workflow:

1. **Read** the Excel files programmatically (Python + `openpyxl`).
2. **Clean & join** them (the shipment sheet's header sits at row 8; sales
   people and teams live in a separate dimension table).
3. **Aggregate** into KPIs, monthly trends, and breakdowns by segment, country,
   product, region, team, and sales person.
4. **Render** an interactive dashboard (Chart.js) with the data embedded — no
   server, no internet connection required.

It demonstrates the complete **Excel → analysis → presentation** pipeline that
feeds directly into Power BI reporting.

## Preview

Open **`dashboard.html`** in any browser (double-click the file). Everything is
embedded in that single file — the Chart.js library *and* all data — so it
works fully **offline**, which makes it safe to demo in front of an assessor
even with no internet in the room.

> 💡 **Presenting on a projector:** open the file, press **F11** for fullscreen,
> and use the three tabs (Overview · Financial · Shipments) to walk through the
> results. Hover any chart for exact values.

## Live Demo

Once this repository is pushed to GitHub, the dashboard can be hosted for free
on **GitHub Pages** — no server, no cost:

**Live URL (after enabling Pages):**

```
https://<your-username>.github.io/excel-powerbi-dashboard/dashboard.html
```

**Enable it in ~30 seconds:**

1. Push the repo to GitHub (see [How to Publish to GitHub](#how-to-publish-to-github)).
2. Open the repository on GitHub → **Settings** → **Pages** (left sidebar).
3. Under *Build and deployment*, set **Source** to **Deploy from a branch**.
4. Set **Branch** to `main` and folder to **/ (root)** → **Save**.
5. Wait about a minute — the dashboard is live at the URL above.

> The repo ships with a `.nojekyll` file, so GitHub Pages serves
> `dashboard.html` exactly as-is (no Jekyll processing).

**Instant alternative** (works as soon as the repo is *public*, no Pages setup
needed at all):

```
https://htmlpreview.github.io/?https://github.com/<your-username>/excel-powerbi-dashboard/main/dashboard.html
```

## The Two Datasets

### 1. `financial.xlsx` — Power BI Financial Sample
The classic Microsoft Power BI sample workbook (single sheet, **700 rows**,
2013–2014). Fields: Segment, Country, Product, Discount Band, Units Sold,
Manufacturing/Sale Price, Gross Sales, Discounts, Sales, COGS, Profit, Date.

| Metric | Value |
|---|---|
| Total sales | **$118.7M** |
| Total profit | **$16.9M** |
| Profit margin | **14.2%** |
| Units sold | **1.1M** |
| Segments | Government · Enterprise · Small Business · Midmarket · Channel Partners |
| Countries | Canada · Germany · France · Mexico · United States of America |
| Products | Paseo · VTT · Velo · Amarilla · Montana · Carretera |

### 2. `ac-sample-data.xlsx` — Chocolate Co. Shipment Ledger
A three-sheet workbook from the "Data for Power BI Sale" sample: a **Shipment
Data** fact table (header at row 8: Sales Person, Geography, Product, Date,
Sales, Boxes), a **Dimension Data** lookup table (products → category/cost,
geographies → region, sales people → team), and a **Calendar Table**.

| Metric | Value |
|---|---|
| Shipments | **6,113** (Feb 2023 – Feb 2024) |
| Total sales | **$34.0M** |
| Boxes shipped | **2.08M** |
| Products | 22 |
| Sales people | 25 (4 teams: Jucies · Delish · Yummies · Tempo) |
| Geographies | 6 |

## Dashboard Features

- **Three views** — an *Overview* report card, the *Power BI Financial* tab,
  and the *Chocolate Co. Shipments* tab.
- **KPI cards** — total sales, profit, margin, units, boxes, and scale numbers
  at a glance.
- **Interactive year filter** — toggle the financial tab between *All*,
  *2013*, and *2014*; every chart and KPI re-aggregates instantly.
- **12 charts** — dual-axis monthly sales & profit trends, segment/country/
  region doughnuts, product & geography bars, discount-band margin analysis,
  team performance.
- **Ranking tables** — top products and top 10 sales people with team and
  volumes.
- **Hover tooltips** everywhere, with human-readable $ formatting.

## Key Insights

- **Power BI sample:** *Government* is the largest segment at **$52.5M (44.2%)**
  of sales; *Paseo* is the top product (**$33.0M**); overall margin is **14.2%**.
- **Shipment ledger:** *Jucies* is the top team (**$9.8M**); *Kelci Walkden* is
  the top sales person (**$1.5M**); *APAC* leads regions (**$17.2M**), with
  *Bars* the top category (**$17.1M**).

## How to Run

**Option A — just view it (no tools needed):**

```bash
# Double-click this file in any browser:
dashboard.html
```

**Option B — regenerate the dashboard from the Excel files:**

```bash
pip install openpyxl
python build_dashboard.py        # -> writes dashboard.html
```

The generator reads both workbooks, joins the dimension table, computes every
aggregation, and rebuilds the self-contained HTML.

## How to Publish to GitHub

The repository is already initialized and committed locally (branch `main`).
To put it on GitHub:

**Option 1 — GitHub CLI (fastest):**

```bash
winget install --id GitHub.cli      # only if you don't have gh yet
gh auth login
cd excel-powerbi-dashboard
gh repo create excel-powerbi-dashboard --public --source . --push
```

**Option 2 — via the website:**

1. Create an empty repository on github.com named `excel-powerbi-dashboard`
   (leave "Add a README" unticked).
2. Then run:

```bash
cd excel-powerbi-dashboard
git remote add origin https://github.com/<your-username>/excel-powerbi-dashboard.git
git push -u origin main
```

After pushing, follow the [Live Demo](#live-demo) steps to turn on GitHub Pages.

## How It's Built

| Step | Tool | What happens |
|------|------|-------------|
| 1. Extract | `openpyxl` | Read both workbooks; locate the shipment header at row 8 |
| 2. Clean & join | Python | Strip whitespace, parse dates, join products → category, geography → region, sales person → team |
| 3. Aggregate | Python | KPIs, monthly series, segment/country/product/band/team/region breakdowns |
| 4. Embed | JSON | Aggregated results embedded in the HTML as data |
| 5. Visualize | Chart.js (v4, inlined) | 12 interactive charts + ranking tables, dark professional theme |

## Repository Structure

```
excel-powerbi-dashboard/
├── dashboard.html          # The interactive dashboard — open this
├── build_dashboard.py      # Regenerates dashboard.html from the Excel files
├── chart.umd.min.js        # Vendored Chart.js v4 (offline capability)
├── financial.xlsx          # Power BI Financial Sample (700 rows)
├── ac-sample-data.xlsx     # Chocolate Co. shipments + dimension + calendar
├── CITATION.cff            # How to cite this repository
├── LICENSE                 # MIT license
├── .gitignore
└── README.md               # This file
```

## Skills Demonstrated

- **Excel data engineering:** reading workbooks programmatically, handling
  non-standard headers, multi-sheet lookups and joins
- **Data cleaning:** whitespace/typo handling, date parsing, dimension joins
- **Data aggregation:** KPI computation, time-series and categorical roll-ups
- **Front-end data visualization:** Chart.js, dark UI theming, responsive layout
- **Self-contained delivery:** a single HTML file that runs offline anywhere
- **Reproducible pipeline:** one command rebuilds the dashboard from source data

## Assessment Talking Points

- *"I built an interactive dashboard from two raw Excel files — joining the
  shipment ledger to its dimension tables, aggregating 6,800+ rows, and
  shipping it as a single offline-ready HTML file."*
- *"The dashboard answers real business questions: which segment drives
  revenue (Government, 44%), which product (Paseo, $33M), and which team and
  sales person perform best (Jucies / Kelci Walkden)."*
- *"It mirrors the Power BI workflow — the same data, the same star-schema
  thinking (facts + dimensions + calendar), delivered as an interactive
  report."*

## License & Data Provenance

- **Code and documentation:** [MIT License](LICENSE) © 2026 Calvin Okoth.
- **`financial.xlsx`:** Microsoft's public *Financial Sample* workbook,
  distributed for Power BI training purposes.
- **`ac-sample-data.xlsx`:** the *Data for Power BI Sale* sample workbook by
  **Chandoo.org**, used here for learning/portfolio purposes. Both files remain
  the property of their respective creators.

## Citation

If you use or build on this project, cite it via the
[`CITATION.cff`](CITATION.cff) file (GitHub's **"Cite this repository"**
button), for example:

> Okoth, C. (2026). *Excel → Power BI Analytics Dashboard: An Interactive HTML
> Dashboard Built from Excel Data*. GitHub repository.
