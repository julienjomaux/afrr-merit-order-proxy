"""
aFRR Germany - Combined Website (Heatmaps + Daily View + Merit Order)
------------------------------------------------------------------------
A single Flask app that serves:

  GET /                          -> the interactive dashboard (HTML page)
  GET /merit-order?date=...      -> live merit order data (JSON), used by
                                     the dashboard's "Show merit order" button

This is meant to be deployed once (e.g. on Render.com) so you get ONE
public URL to link from your Substack post. Visitors open that URL in
their browser and get the full interactive experience directly -- no
download, no PDF, no local Python required on their end.

How it works:
  - At startup, this app reads all "RESULT_OVERVIEW_CAPACITY_MARKET_aFRR_*.xlsx"
    files placed in the same folder (or DATA_DIR, see below), builds the
    dashboard data payload once, and keeps it in memory.
  - Every request to "/" renders the same dashboard HTML using that cached
    payload (fast, no recomputation per request).
  - Every request to "/merit-order" downloads + parses the live bid list
    from regelleistung.net for the requested date, exactly like the
    earlier standalone merit_order_proxy.py did.

To update the heatmap data (e.g. add a new month of capacity-market
data), upload the new/updated .xlsx file(s) to the repo and redeploy;
the app re-reads them on the next startup.

Run locally:
    pip install -r requirements.txt
    python app.py
    -> open http://127.0.0.1:5000 in your browser

Deploy: see DEPLOY.md in this folder.
"""

import os
import io
import glob
import json
import re

import numpy as np
import pandas as pd
import requests
from flask import Flask, request, jsonify, Response

# ----------------------------------------------------------------------
# CONFIG
# ----------------------------------------------------------------------
DATA_DIR = os.environ.get("DATA_DIR", os.path.dirname(os.path.abspath(__file__)))
FILE_PATTERN = "RESULT_OVERVIEW_CAPACITY_MARKET_aFRR_*.xlsx"
SHEET_NAME = "001"

AVERAGE_COL = "GERMANY_AVERAGE_CAPACITY_PRICE_[(EUR/MW)/h]"
MARGINAL_COL = "GERMANY_MARGINAL_CAPACITY_PRICE_[(EUR/MW)/h]"

BANNER_LINE_1 = "gemenergyanalytics.substack.com"
BANNER_LINE_2 = "Julien Jomaux"

MERIT_ORDER_BASE_URL = (
    "https://www.regelleistung.net/apps/crds/api/v2/tenders/results/anonymous"
)

BLOCKS = ["00_04", "04_08", "08_12", "12_16", "16_20", "20_24"]
BLOCK_LABELS = ["00-04", "04-08", "08-12", "12-16", "16-20", "20-24"]

MONTH_LABELS = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
]

ALL_PRODUCTS = [f"POS_{b}" for b in BLOCKS] + [f"NEG_{b}" for b in BLOCKS]

METRICS = {
    "avg_of_average": {"column": AVERAGE_COL, "agg": "mean", "title": "Average of Average Capacity Price"},
    "avg_of_marginal": {"column": MARGINAL_COL, "agg": "mean", "title": "Average of Marginal Capacity Price"},
    "max_of_average": {"column": AVERAGE_COL, "agg": "max", "title": "Maximum of Average Capacity Price"},
    "max_of_marginal": {"column": MARGINAL_COL, "agg": "max", "title": "Maximum of Marginal Capacity Price"},
}
METRIC_ORDER = ["avg_of_average", "avg_of_marginal", "max_of_average", "max_of_marginal"]


# ----------------------------------------------------------------------
# DATA LOADING (same logic as the standalone script)
# ----------------------------------------------------------------------
def discover_files(data_dir: str) -> list:
    files = sorted(glob.glob(os.path.join(data_dir, FILE_PATTERN)))
    if not files:
        raise FileNotFoundError(
            f"No files matching '{FILE_PATTERN}' found in {data_dir}. "
            "Upload your RESULT_OVERVIEW_CAPACITY_MARKET_aFRR_*.xlsx files "
            "to this folder (same repo as app.py) and redeploy."
        )
    return files


def load_all_data(files: list) -> pd.DataFrame:
    frames = []
    for f in files:
        df = pd.read_excel(f, sheet_name=SHEET_NAME)
        df["DATE_FROM"] = pd.to_datetime(df["DATE_FROM"])
        frames.append(df[["DATE_FROM", "PRODUCT", AVERAGE_COL, MARGINAL_COL]])
    full = pd.concat(frames, ignore_index=True)
    full = full.drop_duplicates(subset=["DATE_FROM", "PRODUCT"])

    full["YEAR"] = full["DATE_FROM"].dt.year
    full["MONTH"] = full["DATE_FROM"].dt.month
    full["DAY"] = full["DATE_FROM"].dt.date.astype(str)
    full["DIRECTION"] = full["PRODUCT"].str.split("_").str[0]
    full["BLOCK"] = full["PRODUCT"].str.split("_", n=1).str[1]

    return full


def build_year_matrix_and_days(df_year: pd.DataFrame, column: str, agg: str):
    matrix = np.full((12, 13), np.nan)
    day_matrix = [["" for _ in range(13)] for _ in range(12)]

    for month in range(1, 13):
        df_month = df_year[df_year["MONTH"] == month]
        if df_month.empty:
            continue
        first_day_of_month = df_month["DAY"].min()

        for i, block in enumerate(BLOCKS):
            for direction, col_idx in (("POS", i), ("NEG", i + 7)):
                sub = df_month[
                    (df_month["DIRECTION"] == direction)
                    & (df_month["BLOCK"] == block)
                ]
                if sub.empty:
                    continue
                if agg == "mean":
                    val = sub[column].mean()
                    matrix[month - 1, col_idx] = val
                    day_matrix[month - 1][col_idx] = first_day_of_month
                else:
                    idx_max = sub[column].idxmax()
                    val = sub.loc[idx_max, column]
                    matrix[month - 1, col_idx] = val
                    day_matrix[month - 1][col_idx] = sub.loc[idx_max, "DAY"]

    return matrix, day_matrix


def matrix_to_json(matrix: np.ndarray) -> list:
    return [
        [None if np.isnan(v) else round(float(v), 2) for v in row]
        for row in matrix
    ]


def build_dashboard_payload(df: pd.DataFrame) -> dict:
    years = sorted(df["YEAR"].unique().tolist())

    heatmaps = {metric_key: {} for metric_key in METRIC_ORDER}
    heatmap_days = {metric_key: {} for metric_key in METRIC_ORDER}
    for year in years:
        df_year = df[df["YEAR"] == year]
        for metric_key in METRIC_ORDER:
            spec = METRICS[metric_key]
            matrix, day_matrix = build_year_matrix_and_days(df_year, spec["column"], spec["agg"])
            heatmaps[metric_key][str(year)] = matrix_to_json(matrix)
            heatmap_days[metric_key][str(year)] = day_matrix

    daily = {}
    grp_avg_df = df.groupby(["DAY", "DIRECTION", "BLOCK"], as_index=False)[AVERAGE_COL].mean()
    grp_marg_df = df.groupby(["DAY", "DIRECTION", "BLOCK"], as_index=False)[MARGINAL_COL].mean()

    # Build plain dict lookups: (day, direction, block) -> value.
    # This avoids pandas MultiIndex .loc edge cases that behave
    # differently across pandas versions when a group has a single row.
    avg_lookup = {
        (d, dirn, b): v
        for d, dirn, b, v in zip(
            grp_avg_df["DAY"], grp_avg_df["DIRECTION"], grp_avg_df["BLOCK"], grp_avg_df[AVERAGE_COL]
        )
    }
    marg_lookup = {
        (d, dirn, b): v
        for d, dirn, b, v in zip(
            grp_marg_df["DAY"], grp_marg_df["DIRECTION"], grp_marg_df["BLOCK"], grp_marg_df[MARGINAL_COL]
        )
    }

    for day in sorted(df["DAY"].unique()):
        def get_vals(lookup, direction):
            vals = []
            for block in BLOCKS:
                val = lookup.get((day, direction, block))
                vals.append(round(float(val), 2) if val is not None else None)
            return vals

        daily[day] = {
            "up_avg": get_vals(avg_lookup, "POS"),
            "down_avg": get_vals(avg_lookup, "NEG"),
            "up_marg": get_vals(marg_lookup, "POS"),
            "down_marg": get_vals(marg_lookup, "NEG"),
        }

    days_by_year = {}
    for year in years:
        days_by_year[str(year)] = sorted(
            df[df["YEAR"] == year]["DAY"].unique().tolist()
        )

    metric_labels = {k: METRICS[k]["title"] for k in METRIC_ORDER}

    return {
        "years": [str(y) for y in years],
        "metric_order": METRIC_ORDER,
        "metric_labels": metric_labels,
        "month_labels": MONTH_LABELS,
        "block_labels": BLOCK_LABELS,
        "all_products": ALL_PRODUCTS,
        "heatmaps": heatmaps,
        "heatmap_days": heatmap_days,
        "daily": daily,
        "days_by_year": days_by_year,
        "banner_line_1": BANNER_LINE_1,
        "banner_line_2": BANNER_LINE_2,
        # merit-order requests go to a relative, same-origin path since
        # this app serves both the page and the API.
        "merit_order_proxy_url": "",
    }


# ----------------------------------------------------------------------
# MERIT ORDER (same logic as the standalone proxy)
# ----------------------------------------------------------------------
def find_column(columns, *keyword_groups):
    upper_cols = {c: str(c).upper() for c in columns}
    for group in keyword_groups:
        for col, upper in upper_cols.items():
            if all(kw.upper() in upper for kw in group):
                return col
    return None


def download_and_parse_merit_order(delivery_date: str) -> dict:
    params = {
        "productType": "aFRR",
        "market": "CAPACITY",
        "exportFormat": "xlsx",
        "deliveryDate": delivery_date,
    }
    try:
        resp = requests.get(MERIT_ORDER_BASE_URL, params=params, timeout=30)
        resp.raise_for_status()
    except Exception as exc:
        raise RuntimeError(
            f"Could not download bid list from regelleistung.net for "
            f"{delivery_date}: {exc}"
        )

    try:
        raw = pd.read_excel(io.BytesIO(resp.content))
    except Exception as exc:
        raise RuntimeError(
            f"Downloaded file for {delivery_date} could not be read as "
            f"Excel (the API may have returned an error page instead of "
            f"data, e.g. if no auction ran that day): {exc}"
        )

    columns = list(raw.columns)

    product_col = find_column(columns, ["PRODUCT"])
    country_col = find_column(columns, ["COUNTRY"])
    price_col = find_column(columns, ["PRICE"])
    volume_col = find_column(columns, ["OFFERED", "VOLUME"])
    if volume_col is None:
        volume_col = find_column(columns, ["OFFERED", "CAPACITY"])
    if volume_col is None:
        volume_col = find_column(columns, ["CAPACITY", "MW"])
    if volume_col is None:
        volume_col = find_column(columns, ["VOLUME"])

    missing = [
        name for name, val in
        [("PRODUCT", product_col), ("COUNTRY", country_col),
         ("PRICE", price_col), ("VOLUME/OFFERED", volume_col)]
        if val is None
    ]
    if missing:
        raise RuntimeError(
            f"Could not auto-detect required column(s) {missing} in the "
            f"downloaded file. Actual columns found: {columns}."
        )

    df = raw[[product_col, country_col, price_col, volume_col]].copy()
    df.columns = ["product", "country", "price", "volume"]
    df = df.dropna(subset=["product", "country", "price", "volume"])
    df["price"] = pd.to_numeric(df["price"], errors="coerce")
    df["volume"] = pd.to_numeric(df["volume"], errors="coerce")
    df = df.dropna(subset=["price", "volume"])

    bids = df.to_dict(orient="records")
    countries = sorted(df["country"].astype(str).unique().tolist())

    return {"bids": bids, "countries": countries}


# ----------------------------------------------------------------------
# HTML TEMPLATE (dashboard page, same UI as the standalone version)
# ----------------------------------------------------------------------
HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Germany aFRR Capacity Price Dashboard</title>
<style>
  :root {
    --bg: #0f1117;
    --panel: #161922;
    --panel-border: #262b38;
    --text: #e6e8ee;
    --text-dim: #8b93a7;
    --accent: #e3733b;
    --accent-soft: rgba(227, 115, 59, 0.15);
    --grid-line: #1f2330;
    --up-color: #4f8ff7;
    --down-color: #e3733b;
  }
  * { box-sizing: border-box; }
  body {
    margin: 0;
    background: var(--bg);
    color: var(--text);
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial, sans-serif;
    padding: 24px;
  }
  h1 {
    font-size: 22px;
    font-weight: 700;
    margin: 0 0 4px 0;
    letter-spacing: -0.01em;
  }
  .subtitle {
    color: var(--text-dim);
    font-size: 13px;
    margin-bottom: 24px;
  }
  .controls {
    display: flex;
    gap: 16px;
    align-items: center;
    margin-bottom: 20px;
    flex-wrap: wrap;
  }
  .control-group {
    display: flex;
    align-items: center;
    gap: 8px;
  }
  label {
    font-size: 13px;
    color: var(--text-dim);
  }
  select, input[type="date"] {
    background: var(--panel);
    color: var(--text);
    border: 1px solid var(--panel-border);
    border-radius: 6px;
    padding: 7px 10px;
    font-size: 13px;
    font-family: inherit;
  }
  select:focus, input:focus { outline: 1px solid var(--accent); }
  button.action-btn {
    background: var(--accent);
    color: #fff;
    border: none;
    border-radius: 6px;
    padding: 8px 14px;
    font-size: 13px;
    font-weight: 600;
    cursor: pointer;
  }
  button.action-btn:hover { background: #c95f2c; }
  button.action-btn:disabled {
    background: var(--panel-border);
    color: var(--text-dim);
    cursor: not-allowed;
  }
  .panel {
    background: var(--panel);
    border: 1px solid var(--panel-border);
    border-radius: 10px;
    padding: 20px;
    margin-bottom: 24px;
    position: relative;
    overflow: hidden;
  }
  .panel-title {
    font-size: 14px;
    font-weight: 600;
    margin-bottom: 14px;
    color: var(--text);
  }
  .panel-hint {
    font-size: 12px;
    color: var(--text-dim);
    margin-bottom: 14px;
  }
  table.heatmap {
    border-collapse: collapse;
    width: 100%;
    font-size: 12px;
    position: relative;
    z-index: 1;
  }
  table.heatmap th {
    color: var(--text-dim);
    font-weight: 500;
    font-size: 11px;
    padding: 6px 4px;
    text-align: center;
  }
  table.heatmap td {
    text-align: center;
    padding: 10px 4px;
    border: 2px solid var(--bg);
    font-weight: 700;
    cursor: pointer;
    transition: outline 0.1s ease;
    border-radius: 4px;
  }
  table.heatmap td.gap-col {
    background: transparent !important;
    cursor: default;
  }
  table.heatmap td.empty-cell {
    background: var(--panel) !important;
    cursor: default;
    color: var(--text-dim);
  }
  table.heatmap td:hover:not(.gap-col):not(.empty-cell) {
    outline: 2px solid var(--accent);
    outline-offset: -2px;
  }
  table.heatmap th.month-label {
    text-align: right;
    padding-right: 10px;
    color: var(--text);
    font-weight: 600;
  }
  .group-header {
    text-align: center;
    font-size: 12px;
    font-weight: 700;
    color: var(--text-dim);
    letter-spacing: 0.03em;
    padding-bottom: 4px;
  }
  .legend {
    display: flex;
    align-items: center;
    gap: 10px;
    margin-top: 16px;
    font-size: 11px;
    color: var(--text-dim);
  }
  .legend-gradient {
    height: 10px;
    flex: 1;
    max-width: 260px;
    border-radius: 4px;
    background: linear-gradient(to right, #fff7ec, #fdbb84, #e34a33, #7f0000);
  }
  .banner-bg {
    position: absolute;
    top: 50%;
    left: 50%;
    transform: translate(-50%, -50%) rotate(-18deg);
    font-size: 30px;
    font-weight: 800;
    color: rgba(230, 232, 238, 0.04);
    white-space: nowrap;
    text-align: center;
    line-height: 1.4;
    pointer-events: none;
    z-index: 0;
    user-select: none;
  }
  .footer-credit {
    position: absolute;
    bottom: 10px;
    right: 14px;
    font-size: 10px;
    color: var(--text-dim);
    opacity: 0.6;
    text-align: right;
    line-height: 1.3;
    z-index: 1;
  }
  #dayChart, #meritChart {
    width: 100%;
    height: 380px;
    display: block;
  }
  .bar-tooltip {
    position: absolute;
    background: #1f2330;
    border: 1px solid var(--panel-border);
    border-radius: 6px;
    padding: 6px 10px;
    font-size: 12px;
    pointer-events: none;
    opacity: 0;
    transition: opacity 0.1s ease;
    z-index: 10;
  }
  .day-nav {
    display: flex;
    gap: 8px;
    align-items: center;
  }
  .day-nav button {
    background: var(--panel);
    color: var(--text);
    border: 1px solid var(--panel-border);
    border-radius: 6px;
    padding: 6px 10px;
    font-size: 13px;
    cursor: pointer;
  }
  .day-nav button:hover { border-color: var(--accent); }
  .current-day-label {
    font-size: 13px;
    color: var(--text);
    font-weight: 600;
    min-width: 110px;
  }
  .chart-legend {
    display: flex;
    gap: 18px;
    align-items: center;
    margin-top: 10px;
    font-size: 11px;
    color: var(--text-dim);
    flex-wrap: wrap;
  }
  .chart-legend-item {
    display: flex;
    align-items: center;
    gap: 6px;
  }
  .legend-swatch-bar {
    width: 14px;
    height: 10px;
    border-radius: 2px;
  }
  .legend-swatch-hatch {
    width: 14px;
    height: 10px;
    border-radius: 2px;
    border: 2px dashed currentColor;
    background: transparent;
  }
  .merit-status {
    font-size: 12px;
    color: var(--text-dim);
    margin-top: 10px;
  }
  .merit-status.error { color: #ff8080; }
  .merit-status.ok { color: #7fd99a; }
  .country-badge {
    display: inline-block;
    background: var(--accent-soft);
    color: var(--accent);
    border-radius: 4px;
    padding: 2px 8px;
    font-size: 11px;
    font-weight: 700;
    margin-left: 8px;
  }
  .site-footer {
    text-align: center;
    font-size: 12px;
    color: var(--text-dim);
    padding: 20px 0 0;
  }
  .site-footer a { color: var(--accent); text-decoration: none; }
</style>
</head>
<body>

  <h1>Germany aFRR Capacity Price</h1>
  <div class="subtitle">[(EUR/MW)/h] &middot; auction-level capacity market data</div>

  <div class="controls">
    <div class="control-group">
      <label for="yearSelect">Year</label>
      <select id="yearSelect"></select>
    </div>
    <div class="control-group">
      <label for="metricSelect">Metric</label>
      <select id="metricSelect"></select>
    </div>
  </div>

  <div class="panel" id="heatmapPanel">
    <div class="banner-bg" id="bannerHeatmap"></div>
    <div class="panel-title" id="heatmapTitle">Capacity price by month and 4h block</div>
    <div class="panel-hint" id="heatmapHint">Click a cell to load that month's first day in the daily view below.</div>
    <table class="heatmap" id="heatmapTable"></table>
    <div class="legend">
      <span id="legendMin">0</span>
      <div class="legend-gradient"></div>
      <span id="legendMax">100</span>
      <span style="margin-left:6px;">EUR/MW/h</span>
    </div>
    <div class="footer-credit" id="footerCreditHeatmap"></div>
  </div>

  <div class="panel" id="dayPanel">
    <div class="banner-bg" id="bannerDay"></div>
    <div class="panel-title">Daily auction prices (12 auctions: 6 up + 6 down)</div>
    <div class="panel-hint">Solid bar = average capacity price &middot; dashed outline bar on top = marginal capacity price.</div>
    <div class="controls" style="margin-bottom: 8px;">
      <div class="control-group">
        <label for="dateInput">Date</label>
        <input type="date" id="dateInput">
      </div>
      <div class="day-nav">
        <button id="prevDay">&larr; prev day</button>
        <span class="current-day-label" id="currentDayLabel"></span>
        <button id="nextDay">next day &rarr;</button>
      </div>
      <button class="action-btn" id="showMeritBtn">Show merit order</button>
    </div>
    <canvas id="dayChart"></canvas>
    <div class="chart-legend">
      <div class="chart-legend-item"><div class="legend-swatch-bar" style="background:#4f8ff7;"></div> Upward avg</div>
      <div class="chart-legend-item"><div class="legend-swatch-bar" style="background:#e3733b;"></div> Downward avg</div>
      <div class="chart-legend-item"><div class="legend-swatch-hatch" style="color:#9fd9ff;"></div> Upward marginal</div>
      <div class="chart-legend-item"><div class="legend-swatch-hatch" style="color:#ffd0a8;"></div> Downward marginal</div>
    </div>
    <div class="bar-tooltip" id="barTooltip"></div>
    <div class="footer-credit" id="footerCreditDay"></div>
  </div>

  <div class="panel" id="meritPanel" style="display:none;">
    <div class="banner-bg" id="bannerMerit"></div>
    <div class="panel-title">Merit order &mdash; anonymised list of bids<span class="country-badge" id="meritCountryBadge"></span></div>
    <div class="panel-hint">Cumulative offered volume (x-axis) vs bid price (y-axis), for the product and country selected below. Data is downloaded live from regelleistung.net for the date shown above.</div>
    <canvas id="meritChart"></canvas>
    <div class="controls" style="margin-top: 14px;">
      <div class="control-group">
        <label for="meritProductSelect">Product (12 per day)</label>
        <select id="meritProductSelect"></select>
      </div>
      <div class="control-group">
        <label for="meritCountrySelect">Country</label>
        <select id="meritCountrySelect"></select>
      </div>
    </div>
    <div class="merit-status" id="meritStatus"></div>
    <div class="footer-credit" id="footerCreditMerit"></div>
  </div>

  <div class="site-footer">
    <a href="https://gemenergyanalytics.substack.com" target="_blank" rel="noopener">gemenergyanalytics.substack.com</a> &middot; Julien Jomaux
  </div>

<script>
const DATA = __DATA_JSON__;

const bannerText1 = DATA.banner_line_1;
const bannerText2 = DATA.banner_line_2;
document.querySelectorAll('.banner-bg').forEach(el => {
  el.innerHTML = bannerText1 + '<br>' + bannerText2;
});
document.querySelectorAll('.footer-credit').forEach(el => {
  el.innerHTML = bannerText1 + '<br>' + bannerText2;
});

const yearSelect = document.getElementById('yearSelect');
DATA.years.forEach(y => {
  const opt = document.createElement('option');
  opt.value = y;
  opt.textContent = y;
  yearSelect.appendChild(opt);
});
yearSelect.value = DATA.years[DATA.years.length - 1];

const metricSelect = document.getElementById('metricSelect');
DATA.metric_order.forEach(m => {
  const opt = document.createElement('option');
  opt.value = m;
  opt.textContent = DATA.metric_labels[m];
  metricSelect.appendChild(opt);
});
metricSelect.value = DATA.metric_order[0];

const dateInput = document.getElementById('dateInput');
const currentDayLabel = document.getElementById('currentDayLabel');
const heatmapTitle = document.getElementById('heatmapTitle');
const heatmapHint = document.getElementById('heatmapHint');

function colorForValue(val, vmin, vmax) {
  if (val === null || val === undefined) return null;
  const stops = [
    [0.00, [255, 247, 236]],
    [0.33, [253, 187, 132]],
    [0.66, [227, 74, 51]],
    [1.00, [127, 0, 0]],
  ];
  let t = vmax > vmin ? (val - vmin) / (vmax - vmin) : 0.5;
  t = Math.max(0, Math.min(1, t));
  for (let i = 0; i < stops.length - 1; i++) {
    const [t0, c0] = stops[i];
    const [t1, c1] = stops[i + 1];
    if (t >= t0 && t <= t1) {
      const f = (t - t0) / (t1 - t0);
      const r = Math.round(c0[0] + f * (c1[0] - c0[0]));
      const g = Math.round(c0[1] + f * (c1[1] - c0[1]));
      const b = Math.round(c0[2] + f * (c1[2] - c0[2]));
      return [r, g, b, t];
    }
  }
  return [127, 0, 0, 1];
}

function isMaxMetric(metric) {
  return metric === 'max_of_average' || metric === 'max_of_marginal';
}

function renderHeatmap(year, metric) {
  const matrix = DATA.heatmaps[metric] ? DATA.heatmaps[metric][year] : null;
  if (!matrix) return;

  heatmapTitle.textContent = DATA.metric_labels[metric] + ' by month and 4h block';
  heatmapHint.textContent = isMaxMetric(metric)
    ? "Click a cell to jump the daily view to the exact day that maximum occurred."
    : "Click a cell to load that month's first day in the daily view below.";

  let allVals = [];
  matrix.forEach(row => row.forEach(v => { if (v !== null) allVals.push(v); }));
  const vmin = allVals.length ? Math.min(...allVals) : 0;
  const vmax = allVals.length ? Math.max(...allVals) : 1;
  document.getElementById('legendMin').textContent = Math.round(vmin);
  document.getElementById('legendMax').textContent = Math.round(vmax);

  const table = document.getElementById('heatmapTable');
  table.innerHTML = '';

  const groupRow = document.createElement('tr');
  groupRow.appendChild(document.createElement('th'));
  const upHeader = document.createElement('th');
  upHeader.colSpan = 6;
  upHeader.className = 'group-header';
  groupRow.appendChild(upHeader);
  const gapHeader = document.createElement('th');
  groupRow.appendChild(gapHeader);
  const downHeader = document.createElement('th');
  downHeader.colSpan = 6;
  downHeader.className = 'group-header';
  groupRow.appendChild(downHeader);
  table.appendChild(groupRow);

  const headerRow = document.createElement('tr');
  headerRow.appendChild(document.createElement('th'));
  DATA.block_labels.forEach(b => {
    const th = document.createElement('th');
    th.textContent = b;
    headerRow.appendChild(th);
  });
  const gapTh = document.createElement('th');
  headerRow.appendChild(gapTh);
  DATA.block_labels.forEach(b => {
    const th = document.createElement('th');
    th.textContent = b;
    headerRow.appendChild(th);
  });
  table.appendChild(headerRow);

  const dayMatrix = DATA.heatmap_days[metric] ? DATA.heatmap_days[metric][year] : null;

  matrix.forEach((row, monthIdx) => {
    const tr = document.createElement('tr');
    const monthTh = document.createElement('th');
    monthTh.className = 'month-label';
    monthTh.textContent = DATA.month_labels[monthIdx];
    tr.appendChild(monthTh);

    row.forEach((val, colIdx) => {
      const td = document.createElement('td');
      if (colIdx === 6) {
        td.className = 'gap-col';
      } else if (val === null) {
        td.className = 'empty-cell';
        td.textContent = '';
      } else {
        const c = colorForValue(val, vmin, vmax);
        td.style.background = `rgb(${c[0]}, ${c[1]}, ${c[2]})`;
        td.style.color = c[3] > 0.55 ? '#fff' : '#1a1a1a';
        td.textContent = Math.round(val);
        const targetDay = dayMatrix ? dayMatrix[monthIdx][colIdx] : null;
        if (targetDay) {
          td.title = targetDay;
          td.addEventListener('click', () => {
            dateInput.value = targetDay;
            renderDay(targetDay);
            document.getElementById('meritPanel').style.display = 'none';
          });
        }
      }
      tr.appendChild(td);
    });
    table.appendChild(tr);
  });
}

function renderDay(dayStr) {
  const dayData = DATA.daily[dayStr];
  const canvas = document.getElementById('dayChart');
  const ctx = canvas.getContext('2d');

  const dpr = window.devicePixelRatio || 1;
  const rect = canvas.getBoundingClientRect();
  const H = 380;
  canvas.width = rect.width * dpr;
  canvas.height = H * dpr;
  ctx.setTransform(1, 0, 0, 1, 0, 0);
  ctx.scale(dpr, dpr);
  const W = rect.width;

  ctx.clearRect(0, 0, W, H);

  currentDayLabel.textContent = dayStr;

  if (!dayData) {
    ctx.fillStyle = '#8b93a7';
    ctx.font = '13px sans-serif';
    ctx.fillText('No data for this date', 20, 40);
    return;
  }

  const avgValues = [...dayData.up_avg, null, ...dayData.down_avg];
  const margValues = [...dayData.up_marg, null, ...dayData.down_marg];
  const labels = [...DATA.block_labels.map(b => b + ' UP'), '', ...DATA.block_labels.map(b => b + ' DOWN')];
  const barColors = [...DATA.block_labels.map(() => '#4f8ff7'), null, ...DATA.block_labels.map(() => '#e3733b')];
  const lineColors = [...DATA.block_labels.map(() => '#9fd9ff'), null, ...DATA.block_labels.map(() => '#ffd0a8')];

  const allVals = [...avgValues, ...margValues].filter(v => v !== null);
  const maxVal = allVals.length ? Math.max(...allVals, 1) : 1;

  const marginLeft = 45;
  const marginBottom = 50;
  const marginTop = 20;
  const marginRight = 15;
  const chartW = W - marginLeft - marginRight;
  const chartH = H - marginTop - marginBottom;

  const n = avgValues.length;
  const slotW = chartW / n;
  const barWidthRatio = 0.6;

  ctx.strokeStyle = '#1f2330';
  ctx.lineWidth = 1;
  const ySteps = 5;
  ctx.fillStyle = '#8b93a7';
  ctx.font = '11px sans-serif';
  ctx.textAlign = 'right';
  for (let i = 0; i <= ySteps; i++) {
    const y = marginTop + chartH - (chartH * i / ySteps);
    ctx.beginPath();
    ctx.moveTo(marginLeft, y);
    ctx.lineTo(W - marginRight, y);
    ctx.stroke();
    const val = (maxVal * i / ySteps);
    ctx.fillText(Math.round(val), marginLeft - 8, y + 3);
  }

  function xCenter(i) {
    return marginLeft + i * slotW + slotW / 2;
  }
  function yForVal(val) {
    return marginTop + chartH - (val / maxVal) * chartH;
  }

  const barRects = [];

  avgValues.forEach((val, i) => {
    if (val === null) return;
    const barW = slotW * barWidthRatio;
    const x = xCenter(i) - barW / 2;

    const avgBarH = (val / maxVal) * chartH;
    const avgY = marginTop + chartH - avgBarH;
    ctx.fillStyle = barColors[i];
    ctx.fillRect(x, avgY, barW, avgBarH);

    ctx.fillStyle = '#e6e8ee';
    ctx.font = 'bold 11px sans-serif';
    ctx.textAlign = 'center';
    ctx.fillText(Math.round(val), x + barW / 2, avgY - 6 - 14);

    barRects.push({x, y: avgY, w: barW, h: avgBarH, val, label: labels[i] + ' (avg)'});

    const margVal = margValues[i];
    if (margVal !== null && margVal !== undefined) {
      const margBarH = (margVal / maxVal) * chartH;
      const margY = marginTop + chartH - margBarH;
      ctx.save();
      ctx.strokeStyle = lineColors[i];
      ctx.lineWidth = 2;
      ctx.setLineDash([5, 4]);
      ctx.strokeRect(x, margY, barW, margBarH);
      ctx.restore();

      ctx.fillStyle = lineColors[i];
      ctx.font = 'bold 10px sans-serif';
      ctx.textAlign = 'center';
      ctx.fillText(Math.round(margVal), x + barW / 2, margY - 6);

      barRects.push({x, y: margY, w: barW, h: Math.max(margBarH, 1), val: margVal, label: labels[i] + ' (marginal)'});
    }
  });

  ctx.fillStyle = '#8b93a7';
  ctx.font = '10px sans-serif';
  ctx.textAlign = 'center';
  avgValues.forEach((val, i) => {
    if (labels[i] === '') return;
    const x = xCenter(i);
    ctx.save();
    ctx.translate(x, marginTop + chartH + 14);
    ctx.fillText(labels[i], 0, 0);
    ctx.restore();
  });

  canvas.onmousemove = (e) => {
    const bx = e.offsetX;
    const by = e.offsetY;
    const tooltip = document.getElementById('barTooltip');
    const hit = barRects.find(r => bx >= r.x && bx <= r.x + r.w && by >= r.y && by <= r.y + r.h);
    if (hit) {
      tooltip.style.opacity = 1;
      tooltip.style.left = (e.pageX - canvas.getBoundingClientRect().left + 10) + 'px';
      tooltip.style.top = (e.pageY - canvas.getBoundingClientRect().top - 30) + 'px';
      tooltip.textContent = `${hit.label}: ${hit.val} EUR/MW/h`;
    } else {
      tooltip.style.opacity = 0;
    }
  };
  canvas.onmouseleave = () => {
    document.getElementById('barTooltip').style.opacity = 0;
  };
}

function getAllDaysSorted() {
  let allDays = [];
  Object.values(DATA.days_by_year).forEach(arr => { allDays = allDays.concat(arr); });
  return allDays.sort();
}

function shiftDay(deltaDays) {
  if (!dateInput.value) return;
  const allDays = getAllDaysSorted();
  const idx = allDays.indexOf(dateInput.value);
  if (idx === -1) return;
  const newIdx = idx + deltaDays;
  if (newIdx >= 0 && newIdx < allDays.length) {
    dateInput.value = allDays[newIdx];
    renderDay(allDays[newIdx]);
  }
}

document.getElementById('prevDay').addEventListener('click', () => shiftDay(-1));
document.getElementById('nextDay').addEventListener('click', () => shiftDay(1));
dateInput.addEventListener('change', () => renderDay(dateInput.value));

yearSelect.addEventListener('change', () => {
  renderHeatmap(yearSelect.value, metricSelect.value);
});
metricSelect.addEventListener('change', () => {
  renderHeatmap(yearSelect.value, metricSelect.value);
});

window.addEventListener('resize', () => {
  if (dateInput.value) renderDay(dateInput.value);
  if (document.getElementById('meritPanel').style.display !== 'none') {
    drawMeritChart();
  }
});

const meritProductSelect = document.getElementById('meritProductSelect');
DATA.all_products.forEach(p => {
  const opt = document.createElement('option');
  opt.value = p;
  opt.textContent = p;
  meritProductSelect.appendChild(opt);
});

const meritCountrySelect = document.getElementById('meritCountrySelect');
const meritStatus = document.getElementById('meritStatus');
const meritPanel = document.getElementById('meritPanel');
const meritCountryBadge = document.getElementById('meritCountryBadge');

let meritRawBids = null;
let meritCountriesAvailable = [];

function meritOrderUrl(path) {
  // Same-origin: this app serves both the page and the API, so a
  // relative path is all that's needed (works on localhost AND once
  // deployed, with no URL to configure).
  return path;
}

async function fetchMeritData(dayStr) {
  meritStatus.textContent = 'Downloading bid list for ' + dayStr + ' from regelleistung.net ...';
  meritStatus.className = 'merit-status';
  try {
    const resp = await fetch(meritOrderUrl(`/merit-order?date=${dayStr}`));
    if (!resp.ok) {
      const errText = await resp.text();
      throw new Error(errText || ('HTTP ' + resp.status));
    }
    const json = await resp.json();
    if (json.error) {
      throw new Error(json.error);
    }
    meritRawBids = json.bids;
    meritCountriesAvailable = json.countries || [];
    meritCountrySelect.innerHTML = '';
    const allOpt = document.createElement('option');
    allOpt.value = '__ALL__';
    allOpt.textContent = 'All countries';
    meritCountrySelect.appendChild(allOpt);
    meritCountriesAvailable.forEach(c => {
      const opt = document.createElement('option');
      opt.value = c;
      opt.textContent = c;
      meritCountrySelect.appendChild(opt);
    });
    meritStatus.textContent = `Loaded ${json.bids.length} bids for ${dayStr}.`;
    meritStatus.className = 'merit-status ok';
    drawMeritChart();
  } catch (err) {
    meritStatus.textContent = 'Could not load merit order: ' + err.message +
      '. The auction data for this date may not be published yet, or regelleistung.net may be temporarily unreachable.';
    meritStatus.className = 'merit-status error';
    meritRawBids = null;
  }
}

function drawMeritChart() {
  const canvas = document.getElementById('meritChart');
  const ctx = canvas.getContext('2d');
  const dpr = window.devicePixelRatio || 1;
  const rect = canvas.getBoundingClientRect();
  const H = 380;
  canvas.width = rect.width * dpr;
  canvas.height = H * dpr;
  ctx.setTransform(1, 0, 0, 1, 0, 0);
  ctx.scale(dpr, dpr);
  const W = rect.width;
  ctx.clearRect(0, 0, W, H);

  if (!meritRawBids) {
    ctx.fillStyle = '#8b93a7';
    ctx.font = '13px sans-serif';
    ctx.fillText('No merit order data loaded yet.', 20, 40);
    return;
  }

  const product = meritProductSelect.value;
  const country = meritCountrySelect.value;
  meritCountryBadge.textContent = country === '__ALL__' ? 'ALL COUNTRIES' : country;

  let bids = meritRawBids.filter(b => b.product === product);
  if (country !== '__ALL__') {
    bids = bids.filter(b => b.country === country);
  }
  bids = bids.slice().sort((a, b) => a.price - b.price);

  if (bids.length === 0) {
    ctx.fillStyle = '#8b93a7';
    ctx.font = '13px sans-serif';
    ctx.fillText('No bids for this product/country combination.', 20, 40);
    return;
  }

  let cumVol = 0;
  const points = [{x: 0, y: bids[0].price}];
  bids.forEach(b => {
    points.push({x: cumVol, y: b.price});
    cumVol += b.volume;
    points.push({x: cumVol, y: b.price});
  });

  const maxVol = cumVol;
  const maxPrice = Math.max(...bids.map(b => b.price));
  const minPrice = Math.min(0, Math.min(...bids.map(b => b.price)));

  const marginLeft = 60;
  const marginBottom = 45;
  const marginTop = 20;
  const marginRight = 20;
  const chartW = W - marginLeft - marginRight;
  const chartH = H - marginTop - marginBottom;

  function xFor(vol) { return marginLeft + (vol / maxVol) * chartW; }
  function yFor(price) {
    return marginTop + chartH - ((price - minPrice) / (maxPrice - minPrice || 1)) * chartH;
  }

  ctx.strokeStyle = '#1f2330';
  ctx.lineWidth = 1;
  ctx.fillStyle = '#8b93a7';
  ctx.font = '11px sans-serif';
  ctx.textAlign = 'right';
  const ySteps = 5;
  for (let i = 0; i <= ySteps; i++) {
    const price = minPrice + (maxPrice - minPrice) * i / ySteps;
    const y = yFor(price);
    ctx.beginPath();
    ctx.moveTo(marginLeft, y);
    ctx.lineTo(W - marginRight, y);
    ctx.stroke();
    ctx.fillText(Math.round(price), marginLeft - 8, y + 3);
  }

  ctx.textAlign = 'center';
  const xSteps = 5;
  for (let i = 0; i <= xSteps; i++) {
    const vol = maxVol * i / xSteps;
    const x = xFor(vol);
    ctx.fillText(Math.round(vol), x, marginTop + chartH + 16);
  }
  ctx.fillStyle = '#e6e8ee';
  ctx.font = '12px sans-serif';
  ctx.fillText('Cumulative offered volume (MW)', marginLeft + chartW / 2, H - 6);

  ctx.save();
  ctx.translate(14, marginTop + chartH / 2);
  ctx.rotate(-Math.PI / 2);
  ctx.fillStyle = '#e6e8ee';
  ctx.font = '12px sans-serif';
  ctx.textAlign = 'center';
  ctx.fillText('Bid price [EUR/MW]', 0, 0);
  ctx.restore();

  ctx.strokeStyle = '#e3733b';
  ctx.lineWidth = 2.5;
  ctx.beginPath();
  points.forEach((p, idx) => {
    const x = xFor(p.x);
    const y = yFor(p.y);
    if (idx === 0) ctx.moveTo(x, y);
    else ctx.lineTo(x, y);
  });
  ctx.stroke();

  ctx.save();
  ctx.fillStyle = 'rgba(227, 115, 59, 0.12)';
  ctx.beginPath();
  ctx.moveTo(xFor(0), yFor(minPrice));
  points.forEach(p => ctx.lineTo(xFor(p.x), yFor(p.y)));
  ctx.lineTo(xFor(maxVol), yFor(minPrice));
  ctx.closePath();
  ctx.fill();
  ctx.restore();

  canvas.onmousemove = (e) => {
    const bx = e.offsetX;
    const tooltip = document.getElementById('barTooltip');
    if (bx < marginLeft || bx > W - marginRight) {
      tooltip.style.opacity = 0;
      return;
    }
    const vol = ((bx - marginLeft) / chartW) * maxVol;
    let cum = 0;
    let matched = bids[0];
    for (const b of bids) {
      cum += b.volume;
      if (vol <= cum) { matched = b; break; }
    }
    tooltip.style.opacity = 1;
    tooltip.style.left = (e.pageX - canvas.getBoundingClientRect().left + 10) + 'px';
    tooltip.style.top = (e.pageY - canvas.getBoundingClientRect().top - 10) + 'px';
    tooltip.textContent = `~${Math.round(vol)} MW cumulative @ ${matched.price} EUR/MW (${matched.country})`;
  };
  canvas.onmouseleave = () => {
    document.getElementById('barTooltip').style.opacity = 0;
  };
}

document.getElementById('showMeritBtn').addEventListener('click', () => {
  const day = dateInput.value;
  if (!day) return;
  meritPanel.style.display = 'block';
  meritPanel.scrollIntoView({behavior: 'smooth', block: 'start'});
  fetchMeritData(day);
});

meritProductSelect.addEventListener('change', drawMeritChart);
meritCountrySelect.addEventListener('change', drawMeritChart);

renderHeatmap(yearSelect.value, metricSelect.value);
const initialDays = DATA.days_by_year[yearSelect.value] || [];
if (initialDays.length) {
  dateInput.value = initialDays[0];
  renderDay(initialDays[0]);
}
</script>

</body>
</html>
"""


# ----------------------------------------------------------------------
# FLASK APP
# ----------------------------------------------------------------------
app = Flask(__name__)

_cached_html = None


def get_dashboard_html() -> str:
    """Build (once) and cache the dashboard HTML in memory."""
    global _cached_html
    if _cached_html is None:
        files = discover_files(DATA_DIR)
        df = load_all_data(files)
        payload = build_dashboard_payload(df)
        _cached_html = HTML_TEMPLATE.replace("__DATA_JSON__", json.dumps(payload))
    return _cached_html


@app.route("/")
def index():
    try:
        html = get_dashboard_html()
    except Exception as exc:
        return Response(f"Setup error: {exc}", status=500, mimetype="text/plain")
    return Response(html, mimetype="text/html")


@app.route("/merit-order")
def merit_order():
    delivery_date = request.args.get("date", "")
    if not re.match(r"^\d{4}-\d{2}-\d{2}$", delivery_date):
        return jsonify({"error": f"Invalid or missing 'date' parameter: '{delivery_date}'. Expected format YYYY-MM-DD."}), 400
    try:
        result = download_and_parse_merit_order(delivery_date)
        return jsonify(result), 200
    except Exception as exc:
        return jsonify({"error": str(exc)}), 502


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
