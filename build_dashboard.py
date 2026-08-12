#!/usr/bin/env python3
"""
Builds dashboard.html — an interactive, self-contained HTML dashboard from
financial.xlsx (Power BI Financial Sample) and ac-sample-data.xlsx (Chocolate Co. shipments).

Run:  python build_dashboard.py
Output: dashboard.html  (fully offline — Chart.js is inlined)
"""
import datetime
import json
import os
import openpyxl

OUT = "dashboard.html"
CHARTJS_FILE = "chart.umd.min.js"

MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
          "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


# --------------------------------------------------------------------------
# 1. Load + aggregate the Power BI Financial Sample (financial.xlsx)
# --------------------------------------------------------------------------
def load_financial():
    wb = openpyxl.load_workbook("financial.xlsx", read_only=True, data_only=True)
    ws = wb["Sheet1"]
    recs = []
    for r in ws.iter_rows(values_only=True):
        if r[0] is None:
            continue
        try:
            units = float(r[4])
        except (TypeError, ValueError):
            continue
        d = r[12]
        recs.append({
            "segment": str(r[0]).strip(),
            "country": str(r[1]).strip(),
            "product": str(r[2]).strip(),
            "band": str(r[3]).strip(),
            "units": units,
            "sales": float(r[9]),
            "profit": float(r[11]),
            "year": int(r[15]),
            "month": d.month if isinstance(d, (datetime.datetime, datetime.date)) else 1,
        })
    wb.close()
    return recs


def agg_financial(recs):
    monthly = {}
    seg, cnt, prod, band = {}, {}, {}, {}
    for r in recs:
        key = (r["year"], r["month"])
        m = monthly.setdefault(key, {"sales": 0.0, "profit": 0.0, "units": 0.0})
        m["sales"] += r["sales"]
        m["profit"] += r["profit"]
        m["units"] += r["units"]

        s = seg.setdefault(r["segment"], {"sales": 0.0, "profit": 0.0, "units": 0.0})
        s["sales"] += r["sales"]; s["profit"] += r["profit"]; s["units"] += r["units"]

        c = cnt.setdefault(r["country"], {"sales": 0.0, "profit": 0.0, "units": 0.0})
        c["sales"] += r["sales"]; c["profit"] += r["profit"]; c["units"] += r["units"]

        p = prod.setdefault(r["product"], {"sales": 0.0, "profit": 0.0, "units": 0.0})
        p["sales"] += r["sales"]; p["profit"] += r["profit"]; p["units"] += r["units"]

        b = band.setdefault(r["band"], {"sales": 0.0, "profit": 0.0, "count": 0})
        b["sales"] += r["sales"]; b["profit"] += r["profit"]; b["count"] += 1

    def to_list(d, extra=None):
        out = []
        for name, v in d.items():
            item = {"name": name, "sales": round(v["sales"], 2),
                    "profit": round(v["profit"], 2), "units": round(v["units"], 1)}
            item["margin"] = round(item["profit"] / item["sales"], 4) if item["sales"] else 0
            out.append(item)
        out.sort(key=lambda x: x["sales"], reverse=True)
        return out

    monthly_list = []
    for (y, m), v in sorted(monthly.items()):
        monthly_list.append({
            "label": f"{MONTHS[m - 1]} {y % 100}",
            "sales": round(v["sales"], 2), "profit": round(v["profit"], 2),
            "units": round(v["units"], 1),
        })

    band_list = []
    for name, v in sorted(band.items(), key=lambda kv: kv[1]["sales"], reverse=True):
        band_list.append({
            "name": name, "sales": round(v["sales"], 2), "profit": round(v["profit"], 2),
            "count": v["count"],
            "margin": round(v["profit"] / v["sales"], 4) if v["sales"] else 0,
        })

    ts = sum(r["sales"] for r in recs)
    tp = sum(r["profit"] for r in recs)
    tu = sum(r["units"] for r in recs)
    return {
        "kpis": {"sales": round(ts, 2), "profit": round(tp, 2), "units": round(tu, 1),
                 "margin": round(tp / ts, 4) if ts else 0, "rows": len(recs)},
        "monthly": monthly_list,
        "segments": to_list(seg),
        "countries": to_list(cnt),
        "products": to_list(prod),
        "bands": band_list,
    }


# --------------------------------------------------------------------------
# 2. Load + aggregate Chocolate Co. shipments (ac-sample-data.xlsx)
# --------------------------------------------------------------------------
def load_shipments():
    wb = openpyxl.load_workbook("ac-sample-data.xlsx", read_only=True, data_only=True)

    # --- Dimension Data (lookups: product -> category/cost, geo -> region, person -> team)
    wsd = wb["Dimension Data"]
    drows = list(wsd.iter_rows(values_only=True))
    prod_cat, prod_cost, geo_region, person_team = {}, {}, {}, {}
    for r in drows[3:]:
        # Each lookup column is parsed independently: some trailing rows list
        # sales people without a product, so never skip a row wholesale.
        if r[1] is not None:
            p = str(r[1]).strip()
            prod_cat[p] = str(r[2]).strip() if r[2] else ""
            try:
                prod_cost[p] = float(r[3])
            except (TypeError, ValueError):
                prod_cost[p] = None
        g = str(r[7]).strip() if r[7] else ""
        if g:
            geo_region[g] = str(r[8]).strip() if r[8] else ""
        person = str(r[12]).strip() if r[12] else ""
        if person:
            person_team[person] = str(r[13]).strip() if r[13] else ""

    # --- Shipment Data (header at row 8: Sales Person, Geography, Product, Date, Sales, Boxes)
    ws = wb["Shipment Data"]
    rows = list(ws.iter_rows(values_only=True))
    recs = []
    for r in rows[8:]:
        if r[6] is None:
            continue
        geo = str(r[3]).strip()
        prod = str(r[4]).strip()
        person = str(r[2]).strip()
        d = r[5]
        recs.append({
            "person": person,
            "geo": geo,
            "region": geo_region.get(geo, ""),
            "product": prod,
            "cat": prod_cat.get(prod, ""),
            "team": person_team.get(person, "Unknown"),
            "date": d,
            "sales": float(r[6]),
            "boxes": float(r[7]) if r[7] else 0.0,
            "cost": prod_cost.get(prod),
        })
    wb.close()
    return recs


def agg_shipments(recs):
    monthly = {}
    geo, region, cat, team, prod, person = {}, {}, {}, {}, {}, {}
    est_profit_total = 0.0
    for r in recs:
        d = r["date"]
        y, m = (d.year, d.month) if isinstance(d, (datetime.datetime, datetime.date)) else (0, 0)
        key = (y, m)
        mm = monthly.setdefault(key, {"sales": 0.0, "boxes": 0.0})
        mm["sales"] += r["sales"]
        mm["boxes"] += r["boxes"]

        # Estimated gross profit = sales - (boxes x cost per box from the dimension table)
        est_profit = r["sales"] - r["boxes"] * (r["cost"] or 0.0)
        est_profit_total += est_profit

        g = geo.setdefault(r["geo"], {"sales": 0.0, "boxes": 0.0, "region": r["region"]})
        g["sales"] += r["sales"]; g["boxes"] += r["boxes"]

        if r["region"]:
            rg = region.setdefault(r["region"], {"sales": 0.0, "boxes": 0.0})
            rg["sales"] += r["sales"]; rg["boxes"] += r["boxes"]

        if r["cat"]:
            c = cat.setdefault(r["cat"], {"sales": 0.0, "boxes": 0.0, "profit": 0.0})
            c["sales"] += r["sales"]; c["boxes"] += r["boxes"]; c["profit"] += est_profit

        t = team.setdefault(r["team"], {"sales": 0.0, "boxes": 0.0})
        t["sales"] += r["sales"]; t["boxes"] += r["boxes"]

        p = prod.setdefault(r["product"], {"sales": 0.0, "boxes": 0.0, "cat": r["cat"], "profit": 0.0})
        p["sales"] += r["sales"]; p["boxes"] += r["boxes"]; p["profit"] += est_profit

        pe = person.setdefault(r["person"], {"sales": 0.0, "boxes": 0.0, "team": r["team"]})
        pe["sales"] += r["sales"]; pe["boxes"] += r["boxes"]

    def to_list(d, keys=("sales",)):
        out = []
        for name, v in d.items():
            item = {"name": name}
            for k in keys:
                item[k] = round(v[k], 2) if k != "boxes" else round(v[k], 1)
            out.append(item)
        out.sort(key=lambda x: x[keys[0]], reverse=True)
        return out

    monthly_list = []
    for (y, m), v in sorted(monthly.items()):
        monthly_list.append({"label": f"{MONTHS[m - 1]} {y % 100}",
                             "sales": round(v["sales"], 2), "boxes": round(v["boxes"], 1)})

    products = to_list(prod, ("sales", "boxes", "profit"))
    for p in products:
        p["cat"] = prod[p["name"]]["cat"]

    top_people = to_list(person, ("sales", "boxes"))
    for p in top_people:
        p["team"] = person[p["name"]]["team"]

    ts = sum(r["sales"] for r in recs)
    tb = sum(r["boxes"] for r in recs)
    return {
        "kpis": {"sales": round(ts, 2),
                 "boxes": round(tb, 1),
                 "shipments": len(recs),
                 "products": len(prod), "people": len(person), "geos": len(geo),
                 "estProfit": round(est_profit_total, 2),
                 "estMargin": round(est_profit_total / ts, 4) if ts else 0},
        "monthly": monthly_list,
        "regions": to_list(region, ("sales", "boxes")),
        "geos": to_list(geo, ("sales", "boxes"))[:8],
        "categories": to_list(cat, ("sales", "boxes", "profit")),
        "teams": to_list(team, ("sales", "boxes")),
        "products": products[:10],
        "topPeople": top_people[:10],
    }


def build_insights(fin_all, ship):
    def top(lst):
        return lst[0] if lst else {"name": "-", "sales": 0, "margin": 0}

    seg = top(fin_all["segments"])
    cnt = top(fin_all["countries"])
    prd = top(fin_all["products"])
    team = top(ship["teams"])
    person = top(ship["topPeople"])
    region = top(ship["regions"])
    cat = top(ship["categories"])

    return {
        "topSegment": {"name": seg["name"], "value": seg["sales"],
                       "share": round(seg["sales"] / fin_all["kpis"]["sales"], 3)},
        "topCountry": {"name": cnt["name"], "value": cnt["sales"],
                       "share": round(cnt["sales"] / fin_all["kpis"]["sales"], 3)},
        "topProduct": {"name": prd["name"], "value": prd["sales"],
                       "share": round(prd["sales"] / fin_all["kpis"]["sales"], 3)},
        "topTeam": {"name": team["name"], "value": team["sales"]},
        "topPerson": {"name": person["name"], "value": person["sales"]},
        "topRegion": {"name": region["name"], "value": region["sales"]},
        "topCategory": {"name": cat["name"], "value": cat["sales"]},
    }


def build_combined(fin_all, ship):
    """Cross-dataset view: combined top products, indexed monthly trend, totals."""
    products = []
    for p in fin_all["products"]:
        products.append({"name": p["name"], "dataset": "Power BI",
                         "sales": p["sales"], "profit": p["profit"],
                         "volume": p["units"]})
    for p in ship["products"]:
        products.append({"name": p["name"], "dataset": "Chocolate",
                         "sales": p["sales"], "profit": p["profit"],
                         "volume": p["boxes"]})
    products.sort(key=lambda x: x["sales"], reverse=True)
    products = products[:12]

    fin_sales = fin_all["kpis"]["sales"]
    choc_sales = ship["kpis"]["sales"]
    choc_profit = ship["kpis"]["estProfit"]
    total_sales = fin_sales + choc_sales
    total_profit = fin_all["kpis"]["profit"] + choc_profit
    return {
        "products": products,
        "totals": {
            "sales": round(total_sales, 2),
            "profit": round(total_profit, 2),
            "margin": round(total_profit / total_sales, 4) if total_sales else 0,
            "volume": round(fin_all["kpis"]["units"] + ship["kpis"]["boxes"], 1),
            "finProfit": round(fin_all["kpis"]["profit"], 2),
            "finMargin": fin_all["kpis"]["margin"],
            "chocProfit": round(choc_profit, 2),
            "chocMargin": ship["kpis"]["estMargin"],
        },
    }


# --------------------------------------------------------------------------
# 3. HTML template (placeholders: __CHARTJS__, __DATA__)
# --------------------------------------------------------------------------
TEMPLATE = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Data Analytics Dashboard — Attachment Portfolio</title>
<style>
  :root {
    --bg: #0a1428;
    --panel: #101d36;
    --panel-2: #142444;
    --line: #22365c;
    --text: #e8eefc;
    --muted: #9fb0cc;
    --faint: #6b7f9f;
    --teal: #2dd4bf;
    --gold: #f5b942;
    --blue: #60a5fa;
  }
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body {
    background:
      radial-gradient(1200px 600px at 85% -10%, rgba(45, 212, 191, 0.10), transparent 60%),
      radial-gradient(1000px 500px at -10% 110%, rgba(245, 185, 66, 0.08), transparent 60%),
      var(--bg);
    color: var(--text);
    font-family: "Segoe UI", system-ui, -apple-system, "Helvetica Neue", Arial, sans-serif;
    min-height: 100vh;
  }
  header {
    position: sticky; top: 0; z-index: 20;
    background: rgba(10, 20, 40, 0.92);
    backdrop-filter: blur(8px);
    border-bottom: 1px solid var(--line);
  }
  .bar {
    max-width: 1240px; margin: 0 auto; padding: 16px 28px;
    display: flex; align-items: center; justify-content: space-between; gap: 16px; flex-wrap: wrap;
  }
  .brand { display: flex; align-items: center; gap: 12px; }
  .logo {
    width: 38px; height: 38px; border-radius: 10px; flex: none;
    background: linear-gradient(135deg, var(--teal), var(--gold));
    display: flex; align-items: center; justify-content: center;
    font-weight: 800; color: #081222; font-size: 17px;
  }
  .brand h1 { font-size: 18px; font-weight: 700; letter-spacing: 0.2px; }
  .brand p { font-size: 12.5px; color: var(--muted); margin-top: 2px; }
  nav { display: flex; gap: 6px; background: rgba(255,255,255,0.04); padding: 5px; border-radius: 12px; border: 1px solid var(--line); }
  .tab-btn {
    background: none; border: none; color: var(--muted); cursor: pointer;
    font-size: 13.5px; font-weight: 600; padding: 9px 16px; border-radius: 8px; transition: all .15s;
    font-family: inherit;
  }
  .tab-btn:hover { color: var(--text); }
  .tab-btn.active { background: linear-gradient(135deg, rgba(45,212,191,.16), rgba(245,185,66,.12)); color: var(--teal); box-shadow: inset 0 0 0 1px rgba(45,212,191,.35); }
  main { max-width: 1240px; margin: 0 auto; padding: 26px 28px 40px; }
  .tab { display: none; }
  .tab.active { display: block; }
  .kpi-row { display: grid; grid-template-columns: repeat(auto-fit, minmax(210px, 1fr)); gap: 14px; margin-bottom: 20px; }
  .kpi {
    background: linear-gradient(160deg, var(--panel), var(--panel-2));
    border: 1px solid var(--line); border-radius: 14px; padding: 16px 18px; position: relative; overflow: hidden;
  }
  .kpi::after {
    content: ""; position: absolute; inset: 0 0 auto 0; height: 3px;
    background: linear-gradient(90deg, var(--teal), var(--gold));
  }
  .kpi .label { font-size: 11.5px; text-transform: uppercase; letter-spacing: 1px; color: var(--faint); font-weight: 600; }
  .kpi .value { font-size: 26px; font-weight: 700; margin-top: 6px; letter-spacing: -0.5px; }
  .kpi .sub { font-size: 12px; color: var(--muted); margin-top: 4px; }
  .kpi .value.gold { color: var(--gold); }
  .kpi .value.teal { color: var(--teal); }
  .grid { display: grid; grid-template-columns: 1fr 1fr; gap: 18px; margin-bottom: 18px; }
  .span2 { grid-column: span 2; }
  .card {
    background: linear-gradient(165deg, var(--panel), rgba(16, 29, 54, 0.6));
    border: 1px solid var(--line); border-radius: 14px; padding: 18px 20px;
  }
  .card h3 { font-size: 13.5px; font-weight: 600; color: var(--text); margin-bottom: 14px; letter-spacing: 0.3px; display: flex; align-items: center; gap: 8px; }
  .card h3::before { content: ""; width: 9px; height: 9px; border-radius: 3px; background: linear-gradient(135deg, var(--teal), var(--gold)); flex: none; }
  .chart-wrap { position: relative; height: 300px; }
  .chart-wrap.tall { height: 340px; }
  .seg { display: flex; gap: 8px; margin: 0 0 16px; flex-wrap: wrap; }
  .seg-btn {
    background: rgba(255,255,255,0.05); border: 1px solid var(--line); color: var(--muted);
    font-size: 12.5px; font-weight: 600; padding: 7px 14px; border-radius: 999px; cursor: pointer; transition: all .15s; font-family: inherit;
  }
  .seg-btn:hover { color: var(--text); }
  .seg-btn.active { background: linear-gradient(135deg, rgba(45,212,191,.18), rgba(245,185,66,.14)); color: var(--teal); border-color: rgba(45,212,191,.4); }
  table { width: 100%; border-collapse: collapse; font-size: 13px; }
  th { text-align: left; font-size: 11px; text-transform: uppercase; letter-spacing: 0.8px; color: var(--faint); font-weight: 600; padding: 8px 10px; border-bottom: 1px solid var(--line); }
  td { padding: 9px 10px; border-bottom: 1px solid rgba(148,168,205,0.08); color: var(--text); }
  tr:last-child td { border-bottom: none; }
  td.num, th.num { text-align: right; font-variant-numeric: tabular-nums; }
  .rank { display: inline-flex; align-items: center; justify-content: center; width: 22px; height: 22px; border-radius: 7px; background: rgba(255,255,255,0.06); font-size: 11.5px; font-weight: 700; color: var(--muted); }
  tr.first .rank { background: linear-gradient(135deg, rgba(245,185,66,.25), rgba(245,185,66,.12)); color: var(--gold); }
  tr.first td:first-child { color: var(--gold); }
  .insights { grid-template-columns: repeat(auto-fit, minmax(240px, 1fr)); }
  .insight { padding: 16px 18px; }
  .insight .i-label { font-size: 11.5px; text-transform: uppercase; letter-spacing: 1px; color: var(--faint); font-weight: 600; }
  .insight .i-value { font-size: 21px; font-weight: 700; margin-top: 7px; color: var(--teal); }
  .insight .i-sub { font-size: 12px; color: var(--muted); margin-top: 4px; }
  footer { max-width: 1240px; margin: 0 auto; padding: 0 28px 34px; color: var(--faint); font-size: 12px; line-height: 1.6; }
  footer b { color: var(--muted); font-weight: 600; }
  @media (max-width: 900px) {
    .grid { grid-template-columns: 1fr; }
    .span2 { grid-column: span 1; }
    .bar { padding: 14px 18px; }
    main { padding: 20px 18px 30px; }
  }
</style>
</head>
<body>
<header>
  <div class="bar">
    <div class="brand">
      <div class="logo">CO</div>
      <div>
        <h1>Data Analytics Dashboard</h1>
        <p>Attachment portfolio · Calvin Okoth · Excel → Power BI workflow</p>
      </div>
    </div>
    <nav>
      <button class="tab-btn active" data-tab="overview">Overview</button>
      <button class="tab-btn" data-tab="financial">Power BI · Financial</button>
      <button class="tab-btn" data-tab="shipments">Chocolate Co. · Shipments</button>
      <button class="tab-btn" data-tab="combined">Combined Report</button>
    </nav>
  </div>
</header>

<main>
  <!-- ================= OVERVIEW ================= -->
  <section class="tab active" id="tab-overview">
    <div class="kpi-row" id="ov-kpis"></div>
    <div class="grid">
      <div class="card span2"><h3>Financial Sample — Monthly Sales &amp; Profit (2013–2014)</h3><div class="chart-wrap" id="wrap-ov-fin"></div></div>
      <div class="card"><h3>Financial Sales Share by Segment</h3><div class="chart-wrap" id="wrap-ov-seg"></div></div>
      <div class="card"><h3>Financial Sales by Country</h3><div class="chart-wrap" id="wrap-ov-cnt"></div></div>
      <div class="card span2"><h3>Chocolate Co. — Monthly Shipment Sales (Feb 2023 – Feb 2024)</h3><div class="chart-wrap" id="wrap-ov-ship"></div></div>
      <div class="card"><h3>Shipment Sales by Region</h3><div class="chart-wrap" id="wrap-ov-reg"></div></div>
      <div class="card"><h3>Key Findings</h3><div id="ov-findings"></div></div>
    </div>
  </section>

  <!-- ================= FINANCIAL ================= -->
  <section class="tab" id="tab-financial">
    <div class="seg" id="fin-year-seg">
      <button class="seg-btn active" data-year="all">All years</button>
      <button class="seg-btn" data-year="2013">2013</button>
      <button class="seg-btn" data-year="2014">2014</button>
    </div>
    <div class="kpi-row" id="fin-kpis"></div>
    <div class="grid">
      <div class="card span2"><h3>Monthly Sales &amp; Profit</h3><div class="chart-wrap" id="wrap-fin-month"></div></div>
      <div class="card"><h3>Sales Share by Segment</h3><div class="chart-wrap" id="wrap-fin-seg"></div></div>
      <div class="card"><h3>Sales by Country</h3><div class="chart-wrap" id="wrap-fin-cnt"></div></div>
      <div class="card"><h3>Sales by Product</h3><div class="chart-wrap" id="wrap-fin-prod"></div></div>
      <div class="card"><h3>Discount Bands — Profit Margin &amp; Share</h3><div class="chart-wrap" id="wrap-fin-band"></div></div>
      <div class="card span2"><h3>Top 5 Products by Sales</h3><div id="fin-top-products"></div></div>
    </div>
  </section>

  <!-- ================= SHIPMENTS ================= -->
  <section class="tab" id="tab-shipments">
    <div class="kpi-row" id="ship-kpis"></div>
    <div class="grid">
      <div class="card span2"><h3>Monthly Sales &amp; Boxes Shipped</h3><div class="chart-wrap" id="wrap-ship-month"></div></div>
      <div class="card"><h3>Sales by Region</h3><div class="chart-wrap" id="wrap-ship-reg"></div></div>
      <div class="card"><h3>Sales by Geography</h3><div class="chart-wrap" id="wrap-ship-geo"></div></div>
      <div class="card"><h3>Sales by Product Category</h3><div class="chart-wrap" id="wrap-ship-cat"></div></div>
      <div class="card"><h3>Sales by Team</h3><div class="chart-wrap" id="wrap-ship-team"></div></div>
      <div class="card span2"><h3>Top 10 Sales People</h3><div id="ship-top-people"></div></div>
    </div>
  </section>

  <!-- ================= COMBINED REPORT ================= -->
  <section class="tab" id="tab-combined">
    <div class="seg" id="comb-metric-seg">
      <button class="seg-btn active" data-metric="sales">Sales</button>
      <button class="seg-btn" data-metric="profit">Profit</button>
      <button class="seg-btn" data-metric="volume">Volume</button>
    </div>
    <div class="kpi-row" id="comb-kpis"></div>
    <div class="grid">
      <div class="card"><h3>Monthly Sales — Power BI Sample (2013–14)</h3><div class="chart-wrap" id="wrap-comb-trend-fin"></div></div>
      <div class="card"><h3>Monthly Sales — Chocolate Co. (2023–24)</h3><div class="chart-wrap" id="wrap-comb-trend-choc"></div></div>
      <div class="card"><h3>Revenue by Segment — Power BI Sample</h3><div class="chart-wrap" id="wrap-comb-seg"></div></div>
      <div class="card"><h3>Revenue by Region — Chocolate Co.</h3><div class="chart-wrap" id="wrap-comb-reg"></div></div>
      <div class="card span2"><h3>Top 12 Products Across Both Datasets</h3><div class="chart-wrap tall" id="wrap-comb-prod"></div></div>
      <div class="card"><h3>Profit Margin by Segment — Power BI Sample</h3><div class="chart-wrap" id="wrap-comb-margin"></div></div>
      <div class="card"><h3>Est. Profit by Category — Chocolate Co.</h3><div class="chart-wrap" id="wrap-comb-cat"></div></div>
      <div class="card span2"><h3>Both Datasets at a Glance</h3><div id="comb-table"></div></div>
    </div>
  </section>
</main>

<footer>
  <b>Sources:</b> <b>financial.xlsx</b> — Power BI Financial Sample (700 rows · 2013–2014) &nbsp;·&nbsp; <b>ac-sample-data.xlsx</b> — Chocolate Co. Shipment Data (6,113 rows · Feb 2023 – Feb 2024) plus Dimension &amp; Calendar tables.
  Built with Python + Chart.js (v4, embedded — works fully offline). Hover any chart for details.
</footer>

<script>
/*__CHARTJS__*/
</script>
<script>
"use strict";
const DATA = /*__DATA__*/;

const TEAL = "#2dd4bf", GOLD = "#f5b942", BLUE = "#60a5fa", PURPLE = "#a78bfa",
      ROSE = "#fb7185", GREEN = "#4ade80", ORANGE = "#fb923c", CYAN = "#22d3ee",
      PINK = "#f472b6", LIME = "#a3e635";
const PALETTE = [TEAL, GOLD, BLUE, PURPLE, ROSE, GREEN, ORANGE, CYAN, PINK, LIME];
const F = new Intl.NumberFormat("en-US");

function money(n) {
  if (n >= 1e9) return "$" + (n / 1e9).toFixed(2) + "B";
  if (n >= 1e6) return "$" + (n / 1e6).toFixed(1) + "M";
  if (n >= 1e3) return "$" + (n / 1e3).toFixed(0) + "K";
  return "$" + Math.round(n).toLocaleString();
}
function num(n) {
  if (n >= 1e6) return (n / 1e6).toFixed(1) + "M";
  if (n >= 1e3) return (n / 1e3).toFixed(0) + "K";
  return Math.round(n).toLocaleString();
}
function pct(x) { return (x * 100).toFixed(1) + "%"; }

Chart.defaults.color = "#9fb0cc";
Chart.defaults.borderColor = "rgba(148, 168, 205, 0.12)";
Chart.defaults.font.family = "'Segoe UI', system-ui, -apple-system, sans-serif";
Chart.defaults.plugins.legend.labels.usePointStyle = true;
Chart.defaults.plugins.legend.labels.boxWidth = 8;
Chart.defaults.plugins.legend.labels.padding = 14;
Chart.defaults.plugins.tooltip.backgroundColor = "rgba(8, 16, 32, 0.95)";
Chart.defaults.plugins.tooltip.borderColor = "rgba(45, 212, 191, 0.35)";
Chart.defaults.plugins.tooltip.borderWidth = 1;
Chart.defaults.plugins.tooltip.titleColor = "#eef4ff";
Chart.defaults.plugins.tooltip.bodyColor = "#c6d2e8";
Chart.defaults.plugins.tooltip.padding = 10;
Chart.defaults.plugins.tooltip.cornerRadius = 8;
Chart.defaults.animation.duration = 550;

const charts = {};
function makeChart(id, cfg) {
  if (charts[id]) charts[id].destroy();
  const el = canvas(id);
  if (!el) return;
  charts[id] = new Chart(el, cfg);
}

function tooltipFor(fmt) {
  return {
    callbacks: {
      label(ctx) {
        const v = ctx.parsed && typeof ctx.parsed === "object"
          ? (ctx.parsed.y ?? ctx.parsed.x ?? ctx.parsed.r) : ctx.parsed;
        const lbl = ctx.dataset.label ? ctx.dataset.label + ": " : "";
        return lbl + (fmt === "pct" ? pct(v) : fmt === "int" ? num(v) : money(v));
      }
    }
  };
}
function moneyTicks(fmt) {
  return { ticks: { callback: v => (fmt === "pct" ? pct(v) : fmt === "int" ? num(v) : money(v)) } };
}

function lineCfg(id, labels, datasets, extra) {
  const ds = datasets.map((d, i) => ({
    label: d.label, data: d.data,
    borderColor: d.color, backgroundColor: d.color,
    yAxisID: d.axis || "y",
    tension: 0.35, fill: false, borderWidth: 2.5, pointRadius: 2, pointHoverRadius: 5,
  }));
  const scales = {
    x: { grid: { display: false } },
    y: { position: "left", ...moneyTicks(extra.fmt || "money") },
  };
  if (extra.second) {
    scales.y1 = { position: "right", grid: { drawOnChartArea: false }, ...moneyTicks(extra.fmt2 || "money") };
  }
  return {
    type: "line",
    data: { labels, datasets: ds },
    options: {
      responsive: true, maintainAspectRatio: false,
      interaction: { mode: "index", intersect: false },
      plugins: { legend: { position: "top", align: "end" }, tooltip: tooltipFor(extra.fmt || "money") },
      scales,
    },
  };
}

function doughnutCfg(labels, values, fmt) {
  return {
    type: "doughnut",
    data: {
      labels,
      datasets: [{ data: values, backgroundColor: labels.map((_, i) => PALETTE[i % PALETTE.length]),
                   borderColor: "#0a1428", borderWidth: 2 }],
    },
    options: {
      responsive: true, maintainAspectRatio: false, cutout: "62%",
      plugins: { legend: { position: "bottom" }, tooltip: tooltipFor(fmt || "money") },
    },
  };
}

function barCfg(labels, values, opts) {
  const horizontal = opts && opts.horizontal;
  const fmt = (opts && opts.fmt) || "money";
  const colors = opts && opts.colors ? opts.colors : labels.map((_, i) => PALETTE[i % PALETTE.length]);
  const data = { labels, datasets: [{ data: values, backgroundColor: colors, borderRadius: 5, borderSkipped: false, maxBarThickness: 34 }] };
  const scales = horizontal
    ? { x: { grid: { display: false }, ...moneyTicks(fmt) }, y: { grid: { display: false } } }
    : { x: { grid: { display: false } }, y: { grid: { color: "rgba(148,168,205,0.08)" }, ...moneyTicks(fmt) } };
  return {
    type: "bar",
    data,
    options: {
      indexAxis: horizontal ? "y" : "x",
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false }, tooltip: tooltipFor(fmt) },
      scales,
    },
  };
}

function canvas(id) {
  const wrap = document.getElementById("wrap-" + id);
  if (!wrap) return null;
  wrap.innerHTML = '<canvas id="' + id + '"></canvas>';
  return wrap.querySelector("canvas");
}

/* ---------- KPI card builder ---------- */
function kpi(label, value, sub, cls) {
  return '<div class="kpi"><div class="label">' + label + '</div><div class="value ' + (cls || "") + '">' + value + '</div>' + (sub ? '<div class="sub">' + sub + "</div>" : "") + "</div>";
}

/* ---------- Overview ---------- */
function renderOverview() {
  const fin = DATA.financial.all, ship = DATA.shipments;
  document.getElementById("ov-kpis").innerHTML =
    kpi("Power BI · Total Sales", money(fin.kpis.sales), "700 rows · 2013–2014", "teal") +
    kpi("Power BI · Total Profit", money(fin.kpis.profit), "Margin " + pct(fin.kpis.margin), "gold") +
    kpi("Chocolate Co. · Total Sales", money(ship.kpis.sales), "6,113 shipments · Feb 2023 – Feb 2024", "teal") +
    kpi("Boxes Shipped", num(ship.kpis.boxes), ship.kpis.products + " products · " + ship.kpis.people + " sales people", "gold");

  makeChart("ov-fin", lineCfg("ov-fin",
    fin.monthly.map(m => m.label),
    [{ label: "Sales", data: fin.monthly.map(m => m.sales), color: TEAL },
     { label: "Profit", data: fin.monthly.map(m => m.profit), color: GOLD, axis: "y1" }],
    { second: true }));

  makeChart("ov-seg", doughnutCfg(fin.segments.map(s => s.name), fin.segments.map(s => s.sales)));

  makeChart("ov-cnt", barCfg(fin.countries.map(c => c.name), fin.countries.map(c => c.sales)));

  makeChart("ov-ship", lineCfg("ov-ship",
    ship.monthly.map(m => m.label),
    [{ label: "Sales", data: ship.monthly.map(m => m.sales), color: TEAL }],
    {}));

  makeChart("ov-reg", doughnutCfg(ship.regions.map(r => r.name), ship.regions.map(r => r.sales)));

  const ins = DATA.insights;
  const share = (v, s) => v + "  <span style='color:#6b7f9f'>· " + pct(s) + " of total</span>";
  document.getElementById("ov-findings").innerHTML =
    "<table><tbody>" +
    "<tr><td><b>Top segment</b></td><td class='num'>" + ins.topSegment.name + " · " + money(ins.topSegment.value) + "</td></tr>" +
    "<tr><td><b>Top country</b></td><td class='num'>" + ins.topCountry.name + " · " + money(ins.topCountry.value) + "</td></tr>" +
    "<tr><td><b>Top product (PBI)</b></td><td class='num'>" + ins.topProduct.name + " · " + money(ins.topProduct.value) + "</td></tr>" +
    "<tr><td><b>Top category</b></td><td class='num'>" + ins.topCategory.name + " · " + money(ins.topCategory.value) + "</td></tr>" +
    "<tr><td><b>Top region</b></td><td class='num'>" + ins.topRegion.name + " · " + money(ins.topRegion.value) + "</td></tr>" +
    "<tr><td><b>Top team</b></td><td class='num'>" + ins.topTeam.name + " · " + money(ins.topTeam.value) + "</td></tr>" +
    "<tr><td><b>Top sales person</b></td><td class='num'>" + ins.topPerson.name + " · " + money(ins.topPerson.value) + "</td></tr>" +
    "</tbody></table>";
}

/* ---------- Financial (year-filterable) ---------- */
let finYear = "all";
function finData() { return DATA.financial[finYear] || DATA.financial.all; }

function renderFinancial() {
  const d = finData();
  document.getElementById("fin-kpis").innerHTML =
    kpi("Total Sales", money(d.kpis.sales), d.kpis.rows + " rows · " + (finYear === "all" ? "2013–2014" : finYear), "teal") +
    kpi("Total Profit", money(d.kpis.profit), "Margin " + pct(d.kpis.margin), "gold") +
    kpi("Units Sold", num(d.kpis.units), "across all products") +
    kpi("Profit Margin", pct(d.kpis.margin), "profit ÷ sales", "gold");

  makeChart("fin-month", lineCfg("fin-month",
    d.monthly.map(m => m.label),
    [{ label: "Sales", data: d.monthly.map(m => m.sales), color: TEAL },
     { label: "Profit", data: d.monthly.map(m => m.profit), color: GOLD, axis: "y1" }],
    { second: true }));

  makeChart("fin-seg", doughnutCfg(d.segments.map(s => s.name), d.segments.map(s => s.sales)));

  makeChart("fin-cnt", barCfg(d.countries.map(c => c.name), d.countries.map(c => c.sales)));

  makeChart("fin-prod", barCfg(d.products.map(p => p.name), d.products.map(p => p.sales), { horizontal: true }));

  makeChart("fin-band", barCfg(
    d.bands.map(b => b.name + " (" + b.count + ")"),
    d.bands.map(b => b.margin),
    { fmt: "pct", colors: d.bands.map((_, i) => PALETTE[i % PALETTE.length]) }));

  const top = d.products.slice(0, 5);
  document.getElementById("fin-top-products").innerHTML =
    "<table><thead><tr><th>Product</th><th class='num'>Sales</th><th class='num'>Profit</th><th class='num'>Margin</th></tr></thead><tbody>" +
    top.map((p, i) => "<tr class='" + (i === 0 ? "first" : "") + "'><td><span class='rank'>" + (i + 1) + "</span> " + p.name + "</td><td class='num'>" + money(p.sales) + "</td><td class='num'>" + money(p.profit) + "</td><td class='num'>" + pct(p.margin) + "</td></tr>").join("") +
    "</tbody></table>";
}

/* ---------- Shipments ---------- */
function renderShipments() {
  const d = DATA.shipments;
  document.getElementById("ship-kpis").innerHTML =
    kpi("Total Sales", money(d.kpis.sales), d.kpis.shipments + " shipments", "teal") +
    kpi("Boxes Shipped", num(d.kpis.boxes), "≈ " + Math.round(d.kpis.boxes / d.kpis.shipments) + " boxes / shipment", "gold") +
    kpi("Products Sold", d.kpis.products, d.kpis.geos + " countries / geographies") +
    kpi("Sales People", d.kpis.people, "across " + d.teams.length + " teams", "gold");

  makeChart("ship-month", lineCfg("ship-month",
    d.monthly.map(m => m.label),
    [{ label: "Sales", data: d.monthly.map(m => m.sales), color: TEAL },
     { label: "Boxes", data: d.monthly.map(m => m.boxes), color: GOLD, axis: "y1" }],
    { second: true, fmt2: "int" }));

  makeChart("ship-reg", doughnutCfg(d.regions.map(r => r.name), d.regions.map(r => r.sales)));

  makeChart("ship-geo", barCfg(d.geos.map(g => g.name), d.geos.map(g => g.sales), { horizontal: true }));

  makeChart("ship-cat", barCfg(d.categories.map(c => c.name), d.categories.map(c => c.sales), { horizontal: true }));

  makeChart("ship-team", barCfg(d.teams.map(t => t.name), d.teams.map(t => t.sales), { horizontal: true, colors: d.teams.map((_, i) => PALETTE[i % PALETTE.length]) }));

  document.getElementById("ship-top-people").innerHTML =
    "<table><thead><tr><th>#</th><th>Sales Person</th><th>Team</th><th class='num'>Sales</th><th class='num'>Boxes</th></tr></thead><tbody>" +
    d.topPeople.map((p, i) => "<tr class='" + (i === 0 ? "first" : "") + "'><td><span class='rank'>" + (i + 1) + "</span></td><td><b>" + p.name + "</b></td><td>" + p.team + "</td><td class='num'>" + money(p.sales) + "</td><td class='num'>" + num(p.boxes) + "</td></tr>").join("") +
    "</tbody></table>";
}

/* ---------- Combined Report ---------- */
let combMetric = "sales";
function combRow(label, a, b) {
  return "<tr><td><b>" + label + "</b></td><td>" + a + "</td><td>" + b + "</td></tr>";
}
function renderCombined() {
  const c = DATA.combined, fin = DATA.financial.all, ship = DATA.shipments;
  const t = c.totals;
  document.getElementById("comb-kpis").innerHTML =
    kpi("Combined Revenue", money(t.sales), "both datasets · 2013–14 & 2023–24", "teal") +
    kpi("Combined Est. Profit", money(t.profit), "combined margin " + pct(t.margin), "gold") +
    kpi("Financial Profit", money(t.finProfit), "margin " + pct(t.finMargin)) +
    kpi("Chocolate Est. Profit", money(t.chocProfit), "margin " + pct(t.chocMargin), "gold") +
    kpi("Combined Volume", num(t.volume), "units + boxes shipped") +
    kpi("Data Rows Analysed", F.format(fin.kpis.rows + ship.kpis.shipments), "700 + 6,113");

  makeChart("comb-trend-fin", lineCfg("comb-trend-fin",
    fin.monthly.map(m => m.label),
    [{ label: "Sales", data: fin.monthly.map(m => m.sales), color: TEAL }],
    {}));

  makeChart("comb-trend-choc", lineCfg("comb-trend-choc",
    ship.monthly.map(m => m.label),
    [{ label: "Sales", data: ship.monthly.map(m => m.sales), color: GOLD }],
    {}));

  makeChart("comb-seg", barCfg(fin.segments.map(s => s.name), fin.segments.map(s => s.sales), { horizontal: true }));

  makeChart("comb-reg", doughnutCfg(ship.regions.map(r => r.name), ship.regions.map(r => r.sales)));

  const sorted = [...c.products].sort((a, b) => b[combMetric] - a[combMetric]);
  makeChart("comb-prod", barCfg(
    sorted.map(p => p.name),
    sorted.map(p => p[combMetric]),
    { horizontal: true, fmt: combMetric === "volume" ? "int" : "money",
      colors: sorted.map(p => p.dataset === "Power BI" ? TEAL : GOLD) }));

  makeChart("comb-margin", barCfg(fin.segments.map(s => s.name), fin.segments.map(s => s.margin), { fmt: "pct" }));

  makeChart("comb-cat", barCfg(ship.categories.map(x => x.name), ship.categories.map(x => x.profit), { horizontal: true }));

  const topFin = fin.products[0], topChoc = ship.products[0];
  document.getElementById("comb-table").innerHTML =
    "<table><thead><tr><th>Metric</th><th>Power BI Sample</th><th>Chocolate Co.</th></tr></thead><tbody>" +
    combRow("Rows analysed", F.format(fin.kpis.rows), F.format(ship.kpis.shipments)) +
    combRow("Period", "2013 – 2014", "Feb 2023 – Feb 2024") +
    combRow("Total sales", money(fin.kpis.sales), money(ship.kpis.sales)) +
    combRow("Profit / est. profit", money(fin.kpis.profit), money(ship.kpis.estProfit)) +
    combRow("Margin", pct(fin.kpis.margin), pct(ship.kpis.estMargin)) +
    combRow("Top product", topFin.name + " · " + money(topFin.sales), topChoc.name + " · " + money(topChoc.sales)) +
    combRow("Top contributor", "Segment: " + fin.segments[0].name, "Region: " + ship.regions[0].name) +
    combRow("Scale", fin.products.length + " products · 5 countries", ship.kpis.people + " people · 4 teams") +
    "</tbody></table>";
}

/* ---------- tab switching ---------- */
function showTab(name) {
  document.querySelectorAll(".tab").forEach(s => s.classList.toggle("active", s.id === "tab-" + name));
  document.querySelectorAll(".tab-btn").forEach(b => b.classList.toggle("active", b.dataset.tab === name));
  if (name === "overview") renderOverview();
  if (name === "financial") renderFinancial();
  if (name === "shipments") renderShipments();
  if (name === "combined") renderCombined();
}
document.querySelectorAll(".tab-btn").forEach(b => b.addEventListener("click", () => showTab(b.dataset.tab)));

document.querySelectorAll("#fin-year-seg .seg-btn").forEach(b => b.addEventListener("click", () => {
  finYear = b.dataset.year;
  document.querySelectorAll("#fin-year-seg .seg-btn").forEach(x => x.classList.toggle("active", x === b));
  renderFinancial();
}));

document.querySelectorAll("#comb-metric-seg .seg-btn").forEach(b => b.addEventListener("click", () => {
  combMetric = b.dataset.metric;
  document.querySelectorAll("#comb-metric-seg .seg-btn").forEach(x => x.classList.toggle("active", x === b));
  renderCombined();
}));

showTab("overview");
</script>
</body>
</html>
"""


# --------------------------------------------------------------------------
# 4. Assemble everything
# --------------------------------------------------------------------------
def main():
    fin_all = agg_financial(load_financial())
    fin_2013 = agg_financial([r for r in load_financial() if r["year"] == 2013])
    fin_2014 = agg_financial([r for r in load_financial() if r["year"] == 2014])
    ship = agg_shipments(load_shipments())
    insights = build_insights(fin_all, ship)
    combined = build_combined(fin_all, ship)

    data = {
        "financial": {"all": fin_all, "2013": fin_2013, "2014": fin_2014},
        "shipments": ship,
        "insights": insights,
        "combined": combined,
    }

    with open(CHARTJS_FILE, "r", encoding="utf-8") as f:
        chartjs = f.read()
    # Never let the library accidentally terminate the inline <script> block.
    chartjs = chartjs.replace("</script>", "<\\/script>")

    html = TEMPLATE.replace("/*__CHARTJS__*/", chartjs)
    html = html.replace("/*__DATA__*/", json.dumps(data))

    with open(OUT, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"dashboard.html written ({os.path.getsize(OUT) / 1024:.0f} KB)")
    print("Financial KPIs :", fin_all["kpis"])
    print("Shipment KPIs  :", ship["kpis"])
    print("Combined totals:", combined["totals"])
    print("Top segment    :", insights["topSegment"])
    print("Top person     :", insights["topPerson"])
    print("Teams          :", [t["name"] for t in ship["teams"]])


if __name__ == "__main__":
    main()
