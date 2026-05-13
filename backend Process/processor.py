"""
processor.py — Lounge Report Financial Engine
==============================================
Processes Data.csv against 3 master Excel files and outputs final_data.json.

Rules from: Lounge_Report_Complete_Understanding.md
"""

import pandas as pd
import numpy as np
import json
import os
from datetime import datetime

# ─── CONFIG ──────────────────────────────────────────────────────────────────
BASE_DIR       = os.path.dirname(os.path.abspath(__file__))
DATA_CSV       = os.path.join(BASE_DIR, "Data.csv")
DOMESTIC_XLSX  = os.path.join(BASE_DIR, "Domestic Lounge.xlsx")
RAILWAY_XLSX   = os.path.join(BASE_DIR, "Railway Lounge.xlsx")
GLOBAL_XLSX    = os.path.join(BASE_DIR, "Global Lounge.xlsx")
CAFE_XLSX      = os.path.join(BASE_DIR, "Cafe at mall.xlsx")
PREBOOK_XLSX   = os.path.join(BASE_DIR, "Pre-Book.xlsx")
OUTPUT_JSON    = os.path.join(BASE_DIR, "final_data.json")

# ─── PARTNER → SHEET NAME MAPPINGS ───────────────────────────────────────────
# Domestic: sheet names are exact partner names (except 'vodafone' lowercase)
DOMESTIC_PARTNER_SHEET_MAP = {
    "SBI": "SBI", "HDFC": "HDFC", "ICICI": "ICICI", "IDFC Bank": "IDFC Bank",
    "RBL": "RBL", "Scapia": "Scapia", "Diners": "Diners",
    "Bank of Baroda": "Bank of Baroda", "Vodafone": "vodafone", "Yes Bank": "Yes Bank","GrayWall":"GrayWall","Dhanlaxmi Bank":"Dhanlaxmi Bank","Hettich":"Hettich"
}

# Railway: "IDFC Bank" in volume data maps to sheet named "IDFC"
RAILWAY_PARTNER_SHEET_MAP = {
    "IDFC Bank": "IDFC", "AU": "AU", "SBI": "SBI",
    "ICICI": "ICICI", "HDFC": "HDFC", "RBL": "RBL","DreamFolks OPS": "DreamFolks OPS"
}

# ─── DEBUG REASON CODES ───────────────────────────────────────────────────────
R_NO_RATE         = "No rate found for partner"
R_EXPIRED_RATE    = "Rate expired or not yet valid"
R_NO_COUNTRY_MAP  = "Country mapping missing"
R_ETT_OUTLET      = "ETT outlet not mapped"
R_NO_COST_OUTLET  = "Cost missing — outlet not in cost sheet"
R_NONLIVE_OUTLET  = "Outlet excluded — non-operational"
R_COST_EXPIRED    = "Cost contract expired"
R_INVALID_PRICE   = "Invalid price — needs update"
R_INVALID_COST    = "Invalid cost — needs update"
R_ETT_EXCLUSION   = "Intentional exclusion — ETT rule"
R_UNKNOWN_PARTNER = "Unknown partner — no master rate"
R_NO_RAILWAY_COST = "Cost missing — railway outlet"
R_LP_OUTLET       = "LoungePair outlet not mapped"
R_UNIMONI_OUTLET  = "Unimoni outlet not mapped"

# Partner name in Data.csv that triggers specific pricing rules
LOUNGE_PAIR_PARTNER = "LoungePair"
UNIMONI_PARTNER = "Unimoni"


def safe_float(val):
    """Convert to float; return NaN if invalid."""
    try:
        f = float(val)
        return f if not np.isnan(f) else np.nan
    except (TypeError, ValueError):
        return np.nan


def is_valid_price(val):
    """Returns True if val is a positive finite number."""
    v = safe_float(val)
    return (not np.isnan(v)) and v > 0


def to_date(val):
    """Convert various date formats to a date object (no time). Returns None on failure."""
    if pd.isna(val):
        return None
    if isinstance(val, (pd.Timestamp, datetime)):
        return val.date() if hasattr(val, 'date') else val
    try:
        return pd.to_datetime(val).date()
    except Exception:
        return None


def find_relevant_entry(entries, txn_date, date_key):
    """
    Given a list of entry dicts and a transaction date, find the most relevant entry.
    Returns (best_entry, is_expired_boolean).
    """
    txn_ts = pd.Timestamp(txn_date)
    valid_entries = []
    expired_entries = []
    
    for entry in entries:
        expiry = entry.get(date_key)
        if pd.isna(expiry):
            valid_entries.append(entry)
        elif txn_ts <= pd.Timestamp(expiry):
            valid_entries.append(entry)
        else:
            expired_entries.append(entry)
            
    if valid_entries:
        # Sort valid entries by expiry date to find the one that expires soonest after txn_date
        def sort_key(e):
            exp = e.get(date_key)
            return pd.Timestamp.max if pd.isna(exp) else pd.Timestamp(exp)
        return min(valid_entries, key=sort_key), False
    else:
        # All expired, pick the one that expired most recently
        def sort_key_exp(e):
            exp = e.get(date_key)
            return pd.Timestamp.min if pd.isna(exp) else pd.Timestamp(exp)
        return max(expired_entries, key=sort_key_exp), True


# ─── STEP 1: LOAD MASTER FILES INTO MEMORY ONCE ──────────────────────────────

def load_prebook_data():
    """Returns set of ignored outlet_ids and set of ignored (partner, service) pairs."""
    if not os.path.exists(PREBOOK_XLSX):
        print("  Pre-Book.xlsx not found. Skipping Pre-Book exclusions.")
        return set(), set()
        
    print("Loading Pre-Book exclusions...")
    xl = pd.ExcelFile(PREBOOK_XLSX)
    
    # 1. Outlets to ignore
    outlets_df = xl.parse("Outlets", dtype={"Outlet Id": str})
    ignored_outlets = set(outlets_df["Outlet Id"].astype(str).str.strip().unique())
    
    # 2. Partners to ignore for specific services
    partners_df = xl.parse("Partners")
    partners_df["partner_name"] = partners_df["partner_name"].astype(str).str.strip()
    partners_df["service_name"] = partners_df["service_name"].astype(str).str.strip()
    ignored_partner_services = set(zip(partners_df["partner_name"], partners_df["service_name"]))
    
    return ignored_outlets, ignored_partner_services


def load_domestic_master():
    """Returns dict of {partner_name: {outlet_id: {price, valid_till}}} and cost index."""
    print("Loading Domestic Lounge master...")
    xl = pd.ExcelFile(DOMESTIC_XLSX)
    partner_index = {}  # partner → {outlet_id → {price, valid_till}}
    
    for partner, sheet in DOMESTIC_PARTNER_SHEET_MAP.items():
        df = xl.parse(sheet, dtype={"Outlet Id": str})
        df["Outlet Id"] = df["Outlet Id"].astype(str).str.strip()
        df["Valid Till"] = pd.to_datetime(df["Valid Till"], errors="coerce")
        df["Prices"] = pd.to_numeric(df["Prices"], errors="coerce")
        idx = {}
        for _, row in df.iterrows():
            oid = row["Outlet Id"]
            if oid not in idx:
                idx[oid] = []
            idx[oid].append({"price": row["Prices"], "valid_till": row["Valid Till"]})
        partner_index[partner] = idx
    
    # CostSheet
    cost_df = xl.parse("CostSheet", dtype={"Outlet Id": str})
    cost_df["Outlet Id"] = cost_df["Outlet Id"].astype(str).str.strip()
    cost_df["Valid Till"] = pd.to_datetime(cost_df["Valid Till"], errors="coerce")
    cost_df["COST"] = pd.to_numeric(cost_df["COST"], errors="coerce")
    cost_index = {}
    for _, row in cost_df.iterrows():
        oid = row["Outlet Id"]
        if oid not in cost_index:
            cost_index[oid] = []
        cost_index[oid].append({"cost": row["COST"], "valid_till": row["Valid Till"]})
    
    return partner_index, cost_index


def load_railway_master():
    """Returns dict of {partner_name: {outlet_id: {price, valid_till}}} and cost index."""
    print("Loading Railway Lounge master...")
    xl = pd.ExcelFile(RAILWAY_XLSX)
    partner_index = {}
    
    for partner, sheet in RAILWAY_PARTNER_SHEET_MAP.items():
        df = xl.parse(sheet, dtype={"Outlet Id": str})
        df["Outlet Id"] = df["Outlet Id"].astype(str).str.strip()
        df["Valid Till"] = pd.to_datetime(df["Valid Till"], errors="coerce")
        df["Prices"] = pd.to_numeric(df["Prices"], errors="coerce")
        idx = {}
        for _, row in df.iterrows():
            oid = row["Outlet Id"]
            if oid not in idx:
                idx[oid] = []
            idx[oid].append({"price": row["Prices"], "valid_till": row["Valid Till"]})
        partner_index[partner] = idx
    
    # CostSheet — Railway uses "Cost" (capital C, lowercase ost)
    cost_df = xl.parse("CostSheet", dtype={"Outlet Id": str})
    cost_df["Outlet Id"] = cost_df["Outlet Id"].astype(str).str.strip()
    cost_df["Valid Till"] = pd.to_datetime(cost_df["Valid Till"], errors="coerce")
    # Handle column name "Cost" (Railway) vs "COST" (Domestic)
    cost_col = "Cost" if "Cost" in cost_df.columns else "COST"
    cost_df[cost_col] = pd.to_numeric(cost_df[cost_col], errors="coerce")
    cost_index = {}
    for _, row in cost_df.iterrows():
        oid = row["Outlet Id"]
        if oid not in cost_index:
            cost_index[oid] = []
        cost_index[oid].append({"cost": row[cost_col], "valid_till": row["Valid Till"]})
    
    return partner_index, cost_index


def load_global_master():
    """Returns region_map, sales_rates list, ett_index, cost_index."""
    print("Loading Global Lounge master...")
    xl = pd.ExcelFile(GLOBAL_XLSX)
    
    # ── Region Criteria: {country: {partner: region}} ──
    rc_df = xl.parse("Region Criteria")
    partner_cols = [c for c in rc_df.columns if c not in ("S.No.", "Country")]
    region_map = {}  # country → {partner → region}
    for _, row in rc_df.iterrows():
        country = str(row.get("Country", "")).strip()
        if not country or country.lower() == "nan":
            continue
        region_map[country] = {}
        for pc in partner_cols:
            val = str(row.get(pc, "")).strip()
            if val and val.lower() not in ("nan", ""):
                region_map[country][pc] = val
    
    # ── Sales Rate: list of {client, region, usd, start, end} ──
    sr_df = xl.parse("Sales Rate")
    sr_df["Start Date"] = pd.to_datetime(sr_df["Start Date"], errors="coerce")
    sr_df["End Date"]   = pd.to_datetime(sr_df["End Date"],   errors="coerce")
    sr_df["USD"] = pd.to_numeric(sr_df["USD"], errors="coerce")
    sr_df["Client"] = sr_df["Client"].astype(str).str.strip()
    sr_df["Region"] = sr_df["Region"].astype(str).str.strip()
    sales_rates = sr_df.to_dict("records")
    
    # ── ETT sheet: {df_outlet_id: {price, expiry}} ──
    ett_df = xl.parse("ETT", dtype={"DF Outlet ID": str})
    ett_df["DF Outlet ID"] = ett_df["DF Outlet ID"].astype(str).str.strip()
    ett_df["ETT Price"] = pd.to_numeric(ett_df["ETT Price"], errors="coerce")
    ett_df["Expiry Date"] = pd.to_datetime(ett_df["Expiry Date"], errors="coerce")
    ett_index = {}
    for _, row in ett_df.iterrows():
        oid = row["DF Outlet ID"]
        if oid not in ett_index:
            ett_index[oid] = []
        ett_index[oid].append({
            "price": row["ETT Price"],
            "expiry": row["Expiry Date"]
        })

    # ── LoungePair sheet: {df_outlet_id: {price, expiry}} ──
    lp_df = xl.parse("LoungePair", dtype={"DF Outlet ID": str})
    lp_df["DF Outlet ID"] = lp_df["DF Outlet ID"].astype(str).str.strip()
    lp_df["Price"] = pd.to_numeric(lp_df["Price"], errors="coerce")
    lp_df["Expiry Date"] = pd.to_datetime(lp_df["Expiry Date"], errors="coerce")
    lounge_pair_index = {}
    for _, row in lp_df.iterrows():
        oid = row["DF Outlet ID"]
        if oid not in lounge_pair_index:
            lounge_pair_index[oid] = []
        lounge_pair_index[oid].append({
            "price": row["Price"],
            "expiry": row["Expiry Date"]
        })
    print(f"  LoungePair entries loaded: {sum(len(v) for v in lounge_pair_index.values()):,}")
    
    # ── Unimoni sheet: {df_outlet_id: {price, expiry}} ──
    try:
        unimoni_df = xl.parse("Unimoni", dtype={"DF Outlet ID": str})
        unimoni_df["DF Outlet ID"] = unimoni_df["DF Outlet ID"].astype(str).str.strip()
        unimoni_df["Price"] = pd.to_numeric(unimoni_df["Price"], errors="coerce")
        unimoni_df["Expiry Date"] = pd.to_datetime(unimoni_df["Expiry Date"], errors="coerce")
        unimoni_index = {}
        for _, row in unimoni_df.iterrows():
            oid = row["DF Outlet ID"]
            if oid not in unimoni_index:
                unimoni_index[oid] = []
            unimoni_index[oid].append({
                "price": row["Price"],
                "expiry": row["Expiry Date"]
            })
        print(f"  Unimoni entries loaded: {sum(len(v) for v in unimoni_index.values()):,}")
    except ValueError:
        unimoni_index = {}
        print("  Unimoni sheet not found. Skipping Unimoni pricing.")
    
    # ── Cost Sheet: {outlet_id: {cost, group_category, expiry}} ──
    cs_df = xl.parse("Cost Sheet", dtype={"Outlet ID": str})
    cs_df["Outlet ID"] = cs_df["Outlet ID"].astype(str).str.strip()
    
    # Identify date columns starting from column "I" (index 8)
    date_cols_raw = cs_df.columns[8:]
    
    # Parse dates and ensure strictly descending chronological order
    # (Fixes Excel assigning future years to e.g. "16-Dec-")
    date_cols_parsed = []
    last_dt = None
    from datetime import datetime
    for c in date_cols_raw:
        try:
            parsed = pd.to_datetime(c)
            if last_dt is not None and parsed >= last_dt:
                while parsed >= last_dt:
                    parsed -= pd.DateOffset(years=1)
            date_cols_parsed.append((c, parsed))
            last_dt = parsed
        except Exception:
            pass
            
    cost_index = {}
    for _, row in cs_df.iterrows():
        oid = row["Outlet ID"]
        if not oid or str(oid).lower() == "nan":
            continue
            
        if oid not in cost_index:
            cost_index[oid] = []
            
        gc = ""
        if "Service Category" in row:
            gc = str(row["Service Category"]).strip()
        elif "Group Category" in row:
            gc = str(row["Group Category"]).strip()
            
        # Filter to only the dates that have a valid (non-blank, >0) cost
        valid_dates = []
        for d_col, parsed_dt in date_cols_parsed:
            cost_val = pd.to_numeric(row[d_col], errors="coerce")
            if pd.notna(cost_val) and cost_val > 0.0:
                valid_dates.append((d_col, parsed_dt, cost_val))
                
        # Calculate expiry based on the next valid date found
        for i, (d_col, parsed_dt, cost_val) in enumerate(valid_dates):
            if i == 0:
                expiry = pd.NaT
            else:
                expiry = valid_dates[i-1][1] - pd.Timedelta(days=1)
                
            cost_index[oid].append({
                "cost": cost_val,
                "group_category": gc,
                "expiry": expiry
            })
    
    return region_map, sales_rates, ett_index, lounge_pair_index, unimoni_index, cost_index


def load_cafe_master():
    """Returns partner_index, cost_index for Cafe at Mall"""
    if not os.path.exists(CAFE_XLSX):
        return {}, {}
    print("Loading Cafe at Mall master...")
    xl = pd.ExcelFile(CAFE_XLSX)
    
    def extract_product_rates(row, prefix):
        rates = {}
        for col in row.index:
            if isinstance(col, str) and "(" in col and ")" in col:
                pid = col.split("(")[-1].split(")")[0].strip()
                rates[f"{prefix}_{pid}"] = row[col]
        return rates

    # Cost Sheet
    cost_df = xl.parse("Cost", dtype={"Outlet Id": str})
    cost_df["Outlet Id"] = cost_df["Outlet Id"].astype(str).str.strip()
    cost_df["Rate end Date"] = pd.to_datetime(cost_df["Rate end Date"], errors="coerce")
    cost_index = {}
    for _, row in cost_df.iterrows():
        oid = row["Outlet Id"]
        if oid not in cost_index:
            cost_index[oid] = []
        cost_index[oid].append({
            "valid_till": row["Rate end Date"],
            **extract_product_rates(row, "cost")
        })
        
    # Yes Bank Sheet
    yb_df = xl.parse("Yes Bank", dtype={"Outlet Id": str})
    yb_df["Outlet Id"] = yb_df["Outlet Id"].astype(str).str.strip()
    yb_df["Rate end Date"] = pd.to_datetime(yb_df["Rate end Date"], errors="coerce")
    partner_index = {"Yes Bank": {}}
    for _, row in yb_df.iterrows():
        oid = row["Outlet Id"]
        if oid not in partner_index["Yes Bank"]:
            partner_index["Yes Bank"][oid] = []
        partner_index["Yes Bank"][oid].append({
            "valid_till": row["Rate end Date"],
            **extract_product_rates(row, "price")
        })
        
    return partner_index, cost_index

# ─── STEP 2: LOOKUP HELPERS ───────────────────────────────────────────────────

def lookup_dom_rail_revenue(outlet_id, partner, txn_date, partner_index):
    """Returns (price_float, debug_reason_or_None)."""
    if partner not in partner_index:
        return (np.nan, R_NO_RATE)
    outlet_map = partner_index[partner]
    if outlet_id not in outlet_map:
        return (np.nan, R_NO_RATE)
    entries = outlet_map[outlet_id]
    best_entry, is_expired = find_relevant_entry(entries, txn_date, date_key="valid_till")
    price = safe_float(best_entry["price"])
    if not is_valid_price(price):
        return (np.nan, R_INVALID_PRICE)
    if is_expired:
        return (price, R_EXPIRED_RATE)
    return (price, None)


def lookup_dom_rail_cost(outlet_id, txn_date, cost_index, missing_reason=R_NO_COST_OUTLET):
    """Returns (cost_float, debug_reason_or_None)."""
    if outlet_id not in cost_index:
        return (np.nan, missing_reason)
    entries = cost_index[outlet_id]
    best_entry, is_expired = find_relevant_entry(entries, txn_date, date_key="valid_till")
    cost = safe_float(best_entry["cost"])
    if not is_valid_price(cost):
        return (np.nan, R_INVALID_COST)
    if is_expired:
        return (np.nan, R_EXPIRED_RATE)
    return (cost, None)


def lookup_global_revenue(outlet_id, country, partner, txn_date,
                          region_map, sales_rates, ett_index, lounge_pair_index, unimoni_index):
    """Returns (usd_price, debug_reason_or_None)."""
    if partner == "ETT":
        if outlet_id in ett_index:
            entries = ett_index[outlet_id]
            best_entry, is_expired = find_relevant_entry(entries, txn_date, date_key="expiry")
            price = safe_float(best_entry["price"])
            if not is_valid_price(price):
                return (np.nan, R_INVALID_PRICE)
            if is_expired:
                return (price, R_EXPIRED_RATE)
            return (price, None)
        return (np.nan, R_ETT_OUTLET)

    if partner == LOUNGE_PAIR_PARTNER:
        if outlet_id in lounge_pair_index:
            entries = lounge_pair_index[outlet_id]
            best_entry, is_expired = find_relevant_entry(entries, txn_date, date_key="expiry")
            price = safe_float(best_entry["price"])
            if not is_valid_price(price):
                return (np.nan, R_INVALID_PRICE)
            if is_expired:
                return (price, R_EXPIRED_RATE)
            return (price, None)
        return (np.nan, R_LP_OUTLET)
    
    if partner == UNIMONI_PARTNER:
        if outlet_id in unimoni_index:
            entries = unimoni_index[outlet_id]
            best_entry, is_expired = find_relevant_entry(entries, txn_date, date_key="expiry")
            price = safe_float(best_entry["price"])
            if not is_valid_price(price):
                return (np.nan, R_INVALID_PRICE)
            if is_expired:
                return (price, R_EXPIRED_RATE)
            return (price, None)
        return (np.nan, R_UNIMONI_OUTLET)
    
    # Standard: country → region → sales rate
    country_regions = region_map.get(country, {})
    region = country_regions.get(partner)
    if region is None:
        return (np.nan, R_NO_COUNTRY_MAP)
    
    # Find matching rate by partner + region + date range
    txn_dt = pd.Timestamp(txn_date)
    
    # 1. Collect all matching partner/region entries
    matches = []
    for rate in sales_rates:
        if (str(rate.get("Client", "")).strip() == partner and
                str(rate.get("Region", "")).strip() == region):
            matches.append(rate)
            
    if not matches:
        return (np.nan, R_NO_RATE)
        
    # 2. Try to find an entry that is currently valid
    for rate in matches:
        start = rate.get("Start Date")
        end   = rate.get("End Date")
        usd   = safe_float(rate.get("USD"))
        
        start_ok = pd.isna(start) or txn_dt >= start
        end_ok   = pd.isna(end)   or txn_dt <= end
        
        if start_ok and end_ok:
            if is_valid_price(usd):
                return (usd, None)
            else:
                return (np.nan, R_INVALID_PRICE)
                
    # 3. Fallback: if none are valid, pick the first one and flag as expired
    # (In a better system, we might pick the one closest to the txn_date)
    rate = matches[0]
    usd = safe_float(rate.get("USD"))
    if not is_valid_price(usd):
        return (np.nan, R_INVALID_PRICE)
    return (usd, R_EXPIRED_RATE)


def lookup_global_cost(outlet_id, txn_date, cost_index):
    """Returns (usd_cost, debug_reason_or_None, service_category)."""
    if outlet_id not in cost_index:
        return (np.nan, R_NO_COST_OUTLET, "")
    entries = cost_index[outlet_id]
    best_entry, is_expired = find_relevant_entry(entries, txn_date, date_key="expiry")
    
    gc = best_entry.get("group_category", "")
    cost = safe_float(best_entry["cost"])
    if not is_valid_price(cost):
        return (np.nan, R_INVALID_COST, gc)
    if is_expired:
        return (np.nan, R_COST_EXPIRED, gc)
    return (cost, None, gc)


def lookup_cafe_revenue(outlet_id, partner, product_id, txn_date, partner_index):
    """Returns (price_float, debug_reason_or_None) for Cafe."""
    if partner not in partner_index:
        return (np.nan, R_UNKNOWN_PARTNER)
    outlet_map = partner_index[partner]
    if outlet_id not in outlet_map:
        return (np.nan, R_NO_RATE)
    entries = outlet_map[outlet_id]
    best_entry, is_expired = find_relevant_entry(entries, txn_date, date_key="valid_till")
    
    pid_str = str(int(float(product_id))) if pd.notna(product_id) else ""
    price_key = f"price_{pid_str}"
    price = safe_float(best_entry.get(price_key))
    
    if not is_valid_price(price):
        return (np.nan, R_INVALID_PRICE)
    if is_expired:
        return (price, R_EXPIRED_RATE)
    return (price, None)


def lookup_cafe_cost(outlet_id, product_id, txn_date, cost_index):
    """Returns (cost_float, debug_reason_or_None) for Cafe."""
    if outlet_id not in cost_index:
        return (np.nan, R_NO_COST_OUTLET)
    entries = cost_index[outlet_id]
    best_entry, is_expired = find_relevant_entry(entries, txn_date, date_key="valid_till")
    
    pid_str = str(int(float(product_id))) if pd.notna(product_id) else ""
    cost_key = f"cost_{pid_str}"
    cost = safe_float(best_entry.get(cost_key))
    
    if not is_valid_price(cost):
        return (np.nan, R_INVALID_COST)
    if is_expired:
        return (np.nan, R_EXPIRED_RATE)
    return (cost, None)


# ─── STEP 3: MAIN PROCESSING ──────────────────────────────────────────────────

def process():
    print("\n=== Lounge Report Financial Engine ===\n")
    
    # Load masters
    dom_partner_idx, dom_cost_idx = load_domestic_master()
    rail_partner_idx, rail_cost_idx = load_railway_master()
    g_region_map, g_sales_rates, g_ett_idx, g_lp_idx, g_unimoni_idx, g_cost_idx = load_global_master()
    cafe_partner_idx, cafe_cost_idx = load_cafe_master()
    prebook_outlets, prebook_partner_services = load_prebook_data()
    
    # ── Load volume data ──────────────────────────────────────────────────────
    print("Loading Data.csv...")
    df = pd.read_csv(DATA_CSV, dtype={"outlet_id_long": str})
    print(f"  Raw rows: {len(df):,}")
    
    # ── Normalize partner names ───────────────────────────────────────────────
    df["partner_name"] = df["partner_name"].astype(str).str.strip()
    df["partner_name"] = df["partner_name"].apply(
        lambda x: "Vodafone" if x.lower() == "vodafone" else x
    )
    
    # ── ETT Exclusion ─────────────────────────────────────────────────────────
    df["outlet_id_long"] = df["outlet_id_long"].astype(str).str.strip()
    ett_mask = (df["partner_name"] == "ETT") & (df["outlet_id_long"] == "890267672137")
    excluded_count = ett_mask.sum()
    print(f"  Excluding ETT+890267672137: {excluded_count:,} rows")
    df = df[~ett_mask].copy()

    # ── Pre-Book Exclusions ───────────────────────────────────────────────────
    # A. Outlet Exclusions
    prebook_oid_mask = df["outlet_id_long"].isin(prebook_outlets)
    oid_excluded_count = prebook_oid_mask.sum()
    if oid_excluded_count > 0:
        print(f"  Excluding {oid_excluded_count:,} rows based on Pre-Book Outlet IDs")
        df = df[~prebook_oid_mask].copy()

    # B. Partner-Service Exclusions
    if prebook_partner_services:
        def is_ignored_ps(row):
            return (row["partner_name"], row["service_name"]) in prebook_partner_services
        
        prebook_ps_mask = df.apply(is_ignored_ps, axis=1)
        ps_excluded_count = prebook_ps_mask.sum()
        if ps_excluded_count > 0:
            print(f"  Excluding {ps_excluded_count:,} rows based on Pre-Book Partner+Service pairs")
            df = df[~prebook_ps_mask].copy()
    
    # ── Date conversion ───────────────────────────────────────────────────────
    df["transaction_recorded_time_IST"] = pd.to_datetime(
        df["transaction_recorded_time_IST"], errors="coerce"
    )
    df = df[df["transaction_recorded_time_IST"].notna()].copy()
    df["txn_date"] = df["transaction_recorded_time_IST"].dt.date
    
    # ── PAX calculation ───────────────────────────────────────────────────────
    df["free_count"]  = pd.to_numeric(df["free_count"],  errors="coerce").fillna(0)
    df["Paid Visit"]  = pd.to_numeric(df["Paid Visit"],  errors="coerce").fillna(0)
    df["total_pax"] = np.where(
        df["partner_name"] == "IDFC Bank",
        df["free_count"],
        df["free_count"] + df["Paid Visit"]
    )
    
    # ── Initialize result columns ─────────────────────────────────────────────
    df["line_revenue"]   = np.nan
    df["line_cost"]      = np.nan
    df["line_margin"]    = np.nan
    df["currency"]       = ""
    df["unit_price"]     = np.nan
    df["unit_cost"]      = np.nan
    df["debug_reason"]   = ""
    df["has_debug"]      = False
    df["cost_sheet_service_category"] = ""
    
    # ── Set of known partners for each service ────────────────────────────────
    dom_partners  = set(DOMESTIC_PARTNER_SHEET_MAP.keys())
    rail_partners = set(RAILWAY_PARTNER_SHEET_MAP.keys())
    
    print(f"  Processing {len(df):,} rows...")
    
    def flag(idx, reasons):
        """Set debug flag with reason on a list of (idx, reason) tuples."""
        for i, reason in reasons:
            df.at[i, "debug_reason"] = reason
            df.at[i, "has_debug"] = True
    
    # ── DOMESTIC processing ──────────────────────────────────────────────────
    dom_df = df[df["service_name"] == "Domestic Lounge"]
    for idx, row in dom_df.iterrows():
        partner    = row["partner_name"]
        outlet_id  = row["outlet_id_long"]
        txn_date   = row["txn_date"]
        pax        = row["total_pax"]
        
        # Revenue
        if partner not in dom_partners:
            rev, rev_err = np.nan, R_UNKNOWN_PARTNER
        else:
            rev, rev_err = lookup_dom_rail_revenue(outlet_id, partner, txn_date, dom_partner_idx)
        
        # Cost
        cost, cost_err = lookup_dom_rail_cost(outlet_id, txn_date, dom_cost_idx)
        
        df.at[idx, "unit_price"] = rev
        df.at[idx, "unit_cost"]  = cost
        df.at[idx, "currency"]   = "INR"
        
        reasons = []
        if rev_err:   reasons.append(rev_err)
        if cost_err:  reasons.append(cost_err)
        if reasons:
            df.at[idx, "debug_reason"] = "; ".join(reasons)
            df.at[idx, "has_debug"] = True
        
        if is_valid_price(rev):
            df.at[idx, "line_revenue"] = rev * pax
        if is_valid_price(cost):
            df.at[idx, "line_cost"] = cost * pax
    
    # ── RAILWAY processing ────────────────────────────────────────────────────
    rail_df = df[df["service_name"] == "Railways Lounge"]
    for idx, row in rail_df.iterrows():
        partner   = row["partner_name"]
        outlet_id = row["outlet_id_long"]
        txn_date  = row["txn_date"]
        pax       = row["total_pax"]
        
        # Revenue
        if partner not in rail_partners:
            rev, rev_err = np.nan, R_UNKNOWN_PARTNER
        else:
            rev, rev_err = lookup_dom_rail_revenue(outlet_id, partner, txn_date, rail_partner_idx)
        
        # Cost
        cost, cost_err = lookup_dom_rail_cost(outlet_id, txn_date, rail_cost_idx,
                                               missing_reason=R_NO_RAILWAY_COST)
        
        df.at[idx, "unit_price"] = rev
        df.at[idx, "unit_cost"]  = cost
        df.at[idx, "currency"]   = "INR"
        
        reasons = []
        if rev_err:  reasons.append(rev_err)
        if cost_err: reasons.append(cost_err)
        if reasons:
            df.at[idx, "debug_reason"] = "; ".join(reasons)
            df.at[idx, "has_debug"] = True
        
        if is_valid_price(rev):
            df.at[idx, "line_revenue"] = rev * pax
        if is_valid_price(cost):
            df.at[idx, "line_cost"] = cost * pax
    
    # ── GLOBAL processing ─────────────────────────────────────────────────────
    global_df = df[df["service_name"] == "Global Lounge"]
    for idx, row in global_df.iterrows():
        partner   = row["partner_name"]
        outlet_id = row["outlet_id_long"]
        country   = str(row.get("outlet_country", "")).strip()
        txn_date  = row["txn_date"]
        pax       = row["total_pax"]
        
        # Revenue
        rev, rev_err = lookup_global_revenue(
            outlet_id, country, partner, txn_date,
            g_region_map, g_sales_rates, g_ett_idx, g_lp_idx, g_unimoni_idx
        )
        
        # Cost
        cost, cost_err, svc_cat = lookup_global_cost(outlet_id, txn_date, g_cost_idx)
        
        df.at[idx, "unit_price"] = rev
        df.at[idx, "unit_cost"]  = cost
        df.at[idx, "currency"]   = "USD"
        df.at[idx, "cost_sheet_service_category"] = svc_cat
        
        reasons = []
        if rev_err:  reasons.append(rev_err)
        if cost_err: reasons.append(cost_err)
        if reasons:
            df.at[idx, "debug_reason"] = "; ".join(reasons)
            df.at[idx, "has_debug"] = True
        
        if is_valid_price(rev):
            df.at[idx, "line_revenue"] = rev * pax
        if is_valid_price(cost):
            df.at[idx, "line_cost"] = cost * pax
    
    # ── CAFE AT MALL processing ───────────────────────────────────────────────
    # Fill NaN product_id with empty string temporarily to avoid issues, or keep as is.
    cafe_df = df[df["service_name"].astype(str).str.lower() == "cafe at mall"]
    for idx, row in cafe_df.iterrows():
        partner    = row["partner_name"]
        outlet_id  = row["outlet_id_long"]
        product_id = row.get("product_id")
        txn_date   = row["txn_date"]
        pax        = row["total_pax"]
        
        # Revenue
        rev, rev_err = lookup_cafe_revenue(outlet_id, partner, product_id, txn_date, cafe_partner_idx)
        
        # Cost
        cost, cost_err = lookup_cafe_cost(outlet_id, product_id, txn_date, cafe_cost_idx)
        
        df.at[idx, "unit_price"] = rev
        df.at[idx, "unit_cost"]  = cost
        df.at[idx, "currency"]   = "INR"
        
        reasons = []
        if rev_err:  reasons.append(rev_err)
        if cost_err: reasons.append(cost_err)
        if reasons:
            df.at[idx, "debug_reason"] = "; ".join(reasons)
            df.at[idx, "has_debug"] = True
        
        if is_valid_price(rev):
            df.at[idx, "line_revenue"] = rev * pax
        if is_valid_price(cost):
            df.at[idx, "line_cost"] = cost * pax

    # ── Compute margins ───────────────────────────────────────────────────────
    both_valid = df["line_revenue"].notna() & df["line_cost"].notna()
    df.loc[both_valid, "line_margin"] = (
        df.loc[both_valid, "line_revenue"] - df.loc[both_valid, "line_cost"]
    )
    
    print(f"\n  Rows with debug flags: {df['has_debug'].sum():,}")
    
    # ── SUMMARIZE for JSON output ─────────────────────────────────────────────
    # Format dates as strings
    df["txn_date_str"] = df["transaction_recorded_time_IST"].dt.strftime("%Y-%m-%d")
    df["month_label"]  = df["transaction_recorded_time_IST"].dt.strftime("%b %Y")
    df["month_sort"]   = df["transaction_recorded_time_IST"].dt.strftime("%Y-%m")
    
    # Replace NaN with None for JSON serialization
    df_out = df.replace({np.nan: None})
    df_out["outlet_id_long"] = df_out["outlet_id_long"].astype(str)
    
    # ── Main records (slim for performance) ──────────────────────────────────
    # Include terminal_type if it exists in the data
    base_cols = [
        "txn_date_str", "month_label", "month_sort",
        "outlet_id_long", "outlet_name", "outlet_city", "outlet_country",
        "service_name", "partner_name",
        "free_count", "Paid Visit", "total_pax",
        "unit_price", "unit_cost", "currency",
        "cost_sheet_service_category",
        "line_revenue", "line_cost", "line_margin",
        "has_debug", "debug_reason"
    ]
    if "terminal_type" in df_out.columns:
        base_cols.insert(base_cols.index("outlet_country") + 1, "terminal_type")
    
    if "Product_Name_1VK" in df_out.columns:
        # Rename to product_name for cleaner JSON
        df_out.rename(columns={"Product_Name_1VK": "product_name"}, inplace=True)
        base_cols.insert(base_cols.index("service_name") + 1, "product_name")

    records = df_out[base_cols].to_dict("records")
    
    # ── Debug records ─────────────────────────────────────────────────────────
    debug_records = df_out[df_out["has_debug"] == True][[
        "txn_date_str", "partner_name", "outlet_id_long", "outlet_name",
        "outlet_city", "outlet_country", "service_name",
        "total_pax", "debug_reason"
    ]].to_dict("records")
    
    # ── Monthly financial summary ─────────────────────────────────────────────
    monthly_inr = (df[df["currency"] == "INR"]
                   .groupby("month_label")
                   .agg(
                       month_sort=("month_sort", "first"),
                       pax=("total_pax", "sum"),
                       revenue=("line_revenue", "sum"),
                       cost=("line_cost", "sum")
                   ).reset_index()
                   .sort_values("month_sort"))
    monthly_inr["margin"] = monthly_inr["revenue"] - monthly_inr["cost"]
    monthly_inr = monthly_inr.replace({np.nan: 0})
    
    monthly_usd = (df[df["currency"] == "USD"]
                   .groupby("month_label")
                   .agg(
                       month_sort=("month_sort", "first"),
                       pax=("total_pax", "sum"),
                       revenue=("line_revenue", "sum"),
                       cost=("line_cost", "sum")
                   ).reset_index()
                   .sort_values("month_sort"))
    monthly_usd["margin"] = monthly_usd["revenue"] - monthly_usd["cost"]
    monthly_usd = monthly_usd.replace({np.nan: 0})
    
    # ── Top outlets by margin (INR) ────────────────────────────────────────────
    inr_df = df[df["currency"] == "INR"].copy()
    outlet_margin_inr = (inr_df.groupby(["outlet_id_long", "outlet_name", "outlet_city"])
                         .agg(revenue=("line_revenue", "sum"),
                              cost=("line_cost", "sum"),
                              pax=("total_pax", "sum"))
                         .reset_index())
    outlet_margin_inr["margin"] = outlet_margin_inr["revenue"] - outlet_margin_inr["cost"]
    outlet_margin_inr = outlet_margin_inr.replace({np.nan: 0}).sort_values("margin", ascending=False)
    top10_inr = outlet_margin_inr.head(10).to_dict("records")
    
    # ── Top outlets by margin (USD) ────────────────────────────────────────────
    usd_df = df[df["currency"] == "USD"].copy()
    outlet_margin_usd = (usd_df.groupby(["outlet_id_long", "outlet_name", "outlet_city"])
                         .agg(revenue=("line_revenue", "sum"),
                              cost=("line_cost", "sum"),
                              pax=("total_pax", "sum"))
                         .reset_index())
    outlet_margin_usd["margin"] = outlet_margin_usd["revenue"] - outlet_margin_usd["cost"]
    outlet_margin_usd = outlet_margin_usd.replace({np.nan: 0}).sort_values("margin", ascending=False)
    top10_usd = outlet_margin_usd.head(10).to_dict("records")
    
    # ── Global totals ──────────────────────────────────────────────────────────
    inr_rev   = float(df.loc[df["currency"]=="INR", "line_revenue"].sum(skipna=True) or 0)
    inr_cost  = float(df.loc[df["currency"]=="INR", "line_cost"].sum(skipna=True) or 0)
    inr_margin_pct = ((inr_rev - inr_cost) / inr_rev * 100) if inr_rev > 0 else 0
    usd_rev   = float(df.loc[df["currency"]=="USD", "line_revenue"].sum(skipna=True) or 0)
    usd_cost  = float(df.loc[df["currency"]=="USD", "line_cost"].sum(skipna=True) or 0)
    usd_margin_pct = ((usd_rev - usd_cost) / usd_rev * 100) if usd_rev > 0 else 0
    total_pax = int(df["total_pax"].sum())
    
    # ── Build output JSON ─────────────────────────────────────────────────────
    max_data_date = df["transaction_recorded_time_IST"].max()
    max_data_date_str = max_data_date.strftime("%d %b %Y") if pd.notna(max_data_date) else "N/A"

    output = {
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "max_data_date": max_data_date_str,
        "summary": {
            "total_records":     len(df),
            "total_pax":         total_pax,
            "debug_count":       int(df["has_debug"].sum()),
            "inr_revenue":       round(inr_rev, 2),
            "inr_cost":          round(inr_cost, 2),
            "inr_margin_pct":    round(inr_margin_pct, 2),
            "usd_revenue":       round(usd_rev, 2),
            "usd_cost":          round(usd_cost, 2),
            "usd_margin_pct":    round(usd_margin_pct, 2),
        },
        "monthly_inr":  monthly_inr.to_dict("records"),
        "monthly_usd":  monthly_usd.to_dict("records"),
        "top10_margin_inr": top10_inr,
        "top10_margin_usd": top10_usd,
        "records":      records,
        "debug_records": debug_records
    }
    
    # ── Write JSON ────────────────────────────────────────────────────────────
    print(f"\nWriting {OUTPUT_JSON}...")
    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, default=str)
    
    size_mb = os.path.getsize(OUTPUT_JSON) / 1024 / 1024
    print(f"\nDone! Output: {OUTPUT_JSON}")
    print(f"   File size:     {size_mb:.1f} MB")
    print(f"   Total records: {len(df):,}")
    print(f"   Debug flags:   {df['has_debug'].sum():,}")
    print(f"\n-- Financial Summary ---------------------")
    print(f"   INR Revenue:   INR {inr_rev:>15,.2f}")
    print(f"   INR Cost:      INR {inr_cost:>15,.2f}")
    print(f"   INR Margin:    {inr_margin_pct:.1f}%")
    print(f"   USD Revenue:   USD {usd_rev:>15,.2f}")
    print(f"   USD Cost:      USD {usd_cost:>15,.2f}")
    print(f"   USD Margin:    {usd_margin_pct:.1f}%")
    print(f"-----------------------------------------")


if __name__ == "__main__":
    process()
