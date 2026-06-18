"""
aFRR Germany Merit Order Proxy
------------------------------------------------------------------------
A tiny public web service that downloads the official anonymous bid list
for a given delivery date from regelleistung.net, parses it, and returns
the result as JSON with CORS enabled so any browser-based dashboard can
call it directly (no local server needed on the reader's machine).

This replaces the "local helper server" approach: instead of running on
127.0.0.1 on each reader's computer (which corporate firewalls/security
software may block), this same logic runs once, hosted publicly (e.g. on
Render.com's free tier), and every reader's browser just calls its URL.

Endpoint:
    GET /merit-order?date=YYYY-MM-DD

Response (200):
    {
      "bids": [{"product": "POS_00_04", "country": "DE",
                 "price": 12.3, "volume": 5.0}, ...],
      "countries": ["AT", "CZ", "DE"]
    }

Response (4xx/5xx):
    {"error": "human readable message"}

Run locally for testing:
    pip install flask flask-cors pandas openpyxl requests
    python merit_order_proxy.py
    -> serves on http://127.0.0.1:5000

Deploy to Render.com (or similar):
    See DEPLOY.md in this same folder for step-by-step instructions.
"""

import io
import os
import re

import pandas as pd
import requests
from flask import Flask, request, jsonify
from flask_cors import CORS

app = Flask(__name__)
CORS(app)  # allow any origin to call this API (it's public, read-only data)

MERIT_ORDER_BASE_URL = (
    "https://www.regelleistung.net/apps/crds/api/v2/tenders/results/anonymous"
)


# ----------------------------------------------------------------------
# Column auto-detection (same logic validated against the real export)
# ----------------------------------------------------------------------
def find_column(columns, *keyword_groups):
    """
    Find the first column whose name contains ALL keywords from any one
    of the given keyword groups (case-insensitive). Returns the column
    name, or None if nothing matches.
    """
    upper_cols = {c: str(c).upper() for c in columns}
    for group in keyword_groups:
        for col, upper in upper_cols.items():
            if all(kw.upper() in upper for kw in group):
                return col
    return None


def download_and_parse_merit_order(delivery_date: str) -> dict:
    """
    Downloads the anonymous aFRR capacity-market bid list for the given
    date (YYYY-MM-DD) from regelleistung.net and returns:
      {
        "bids": [{"product": "POS_00_04", "country": "DE",
                   "price": 12.3, "volume": 5.0}, ...],
        "countries": ["DE", "AT", "CZ"],   # sorted, whichever appear
      }
    Raises RuntimeError with a human-readable message on any failure.
    """
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
# Routes
# ----------------------------------------------------------------------
@app.route("/")
def index():
    return jsonify({
        "service": "aFRR Germany Merit Order Proxy",
        "usage": "GET /merit-order?date=YYYY-MM-DD",
    })


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
