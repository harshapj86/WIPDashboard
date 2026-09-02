#!/usr/bin/env python3
"""
build_data.py — converts the Aptronix master Excel workbook into data.json,
the file the dashboard (index.html) fetches at load time.

USAGE
    python scripts/build_data.py path/to/workbook.xlsx [--out data.json]

WHAT IT EXPECTS IN THE WORKBOOK
    - One or more "raw data" sheets: any sheet whose name starts with "Raw"
      (case-insensitive), e.g. "Raw Data 24-25", "Raw Data 2025-26",
      "Rawdata 26-27". New raw-data sheets (next fiscal year, etc.) are
      picked up automatically — no code change needed, just add the sheet
      and re-run.
    - "Location Master": ARM | State | Centre Name | Location Type
    - "CSAT": Ship-To | Metrics | <date columns...>   (values as fractions, e.g. 0.95)
    - "Revenue Target", "GP Target", "CSAT Target": Branch ID | <month columns...>
      (header row is row 4 — rows 1-3 are a title + instructions + a blank row)

WHAT IT PRODUCES (data.json)
    {
      months: [{MonthKey, MonthLabel, FY, Quarter, FYQ}, ...],
      gpCoveredMonths: [MonthKey, ...],
      centres: [{Centre, State, ARM, LocationType, City}, ...],
      fact_month:   [[MonthKey, Centre, Revenue, GP, Txns], ...],
      fact_bu:      [[MonthKey, Centre, BusinessUnit, Revenue, GP, Txns], ...],
      fact_cat:     [[MonthKey, Centre, Category, Revenue, GP, Txns], ...],
      fact_pf:      [[MonthKey, Centre, ProductFamily, Revenue, GP, Txns], ...],
      fact_item:    [[MonthKey, Centre, ItemName, Revenue, GP, Txns], ...],
      fact_txntype: [[MonthKey, Centre, TxnType, Revenue, GP, Txns], ...],
      fact_exec:    [[Centre, Executive, Revenue, GP, Txns], ...],
      fact_day:     [[DateStr(YYYY-MM-DD), Centre, Revenue, GP, Txns], ...],
      csat:            [[MonthKey, Centre, CSATPercent], ...],
      target_revenue:  [[MonthKey, Centre, TargetRevenue], ...],
      target_gp:       [[MonthKey, Centre, TargetGP], ...],
      target_csat:     [[MonthKey, Centre, TargetCSATPercent], ...],
    }

KNOWN DATA-QUALITY CAVEAT (see README)
    The raw data's "Product Family" column is inconsistently filled — some
    rows have real product families, some have an item code pasted in by
    mistake, many are blank. Blank/NA values are bucketed into "Other".
    This mirrors the original dashboard's behaviour but inherits the
    underlying sheet's messiness; clean up "Product Family" at the source
    if you want cleaner fact_pf breakdowns.
"""
import sys
import json
import argparse
from collections import defaultdict

import pandas as pd

MONTH_ABBR = {
    1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr", 5: "May", 6: "Jun",
    7: "Jul", 8: "Aug", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dec",
}
MONTH_NUM = {v.lower(): k for k, v in MONTH_ABBR.items()}
# also accept full month names, just in case a sheet uses them
FULL_MONTH_NUM = {
    "january": 1, "february": 2, "march": 3, "april": 4, "may": 5, "june": 6,
    "july": 7, "august": 8, "september": 9, "october": 10, "november": 11, "december": 12,
}


def month_key_from_any(value):
    """Parse a month cell into a YYYYMM int. Handles:
       - pandas/py datetime (Excel sometimes silently converts 'Jun 24' to a date)
       - strings like 'April 24', 'Apr 24', 'Apr-24', 'Jun 2024'
    Returns None if it can't be parsed (caller should skip such columns/rows).
    """
    if value is None:
        return None
    if hasattr(value, "year") and hasattr(value, "month"):
        return value.year * 100 + value.month
    s = str(value).strip()
    if not s:
        return None
    s = s.replace("-", " ").replace("_", " ")
    parts = s.split()
    if len(parts) != 2:
        return None
    mon_raw, yr_raw = parts[0].lower(), parts[1]
    mon = MONTH_NUM.get(mon_raw) or FULL_MONTH_NUM.get(mon_raw)
    if mon is None:
        return None
    try:
        yr = int(yr_raw)
    except ValueError:
        return None
    if yr < 100:
        yr += 2000
    return yr * 100 + mon


def month_label(month_key):
    y, m = divmod(month_key, 100)
    return f"{MONTH_ABBR[m]} {y % 100:02d}"


def fy_quarter(month_key):
    """Indian fiscal year: Apr-Mar. FY label uses the two years it spans."""
    y, m = divmod(month_key, 100)
    if m >= 4:
        fy_start = y
    else:
        fy_start = y - 1
    fy = f"FY{fy_start % 100:02d}-{(fy_start + 1) % 100:02d}"
    if m in (4, 5, 6):
        q = "Q1"
    elif m in (7, 8, 9):
        q = "Q2"
    elif m in (10, 11, 12):
        q = "Q3"
    else:
        q = "Q4"
    return fy, q


def load_raw_data(xls):
    """Concatenate every sheet whose name starts with 'raw' (case-insensitive)."""
    raw_sheet_names = [s for s in xls.sheet_names if s.strip().lower().startswith("raw")]
    if not raw_sheet_names:
        raise SystemExit("No sheet found whose name starts with 'Raw' — nothing to build from.")
    print(f"Raw-data sheets found: {raw_sheet_names}")
    # Column names have drifted slightly between years (e.g. "TXN Month" vs
    # "TXNMonth"). Normalise per-sheet BEFORE concatenating — normalising
    # after concat can produce duplicate-named columns instead of merging
    # them, since sheets with the old/new spelling didn't share a column.
    def normalise_columns(df):
        rename_map = {}
        for col in df.columns:
            key = str(col).strip().lower().replace(" ", "")
            if key == "txnmonth":
                rename_map[col] = "TXN Month"
            elif key == "branchid":
                rename_map[col] = "Branch ID"
            elif key == "branchcity":
                rename_map[col] = "Branch City"
            elif key == "transactiontype":
                rename_map[col] = "Transaction Type"
            elif key in ("item_name", "itemname"):
                rename_map[col] = "Item_Name"
            elif key == "businessunit":
                rename_map[col] = "Business Unit"
            elif key == "category":
                rename_map[col] = "Category"
            elif key == "productfamily":
                rename_map[col] = "Product Family"
            elif key == "executive":
                rename_map[col] = "Executive"
            elif key == "gpatamount":
                rename_map[col] = "GP AT Amount"
            elif key == "total":
                rename_map[col] = "Total"
            elif key == "txndate":
                rename_map[col] = "TXNDate"
        return df.rename(columns=rename_map)

    frames = []
    for name in raw_sheet_names:
        df = normalise_columns(xls.parse(name))
        df["__source_sheet"] = name
        frames.append(df)
    raw = pd.concat(frames, ignore_index=True, sort=False)

    required = ["Branch ID", "TXN Month", "TXNDate", "Transaction Type",
                "Item_Name", "Business Unit", "Category", "Executive",
                "GP AT Amount", "Total"]
    missing = [c for c in required if c not in raw.columns]
    if missing:
        raise SystemExit(f"Raw data is missing expected column(s): {missing}")

    raw["MonthKey"] = raw["TXN Month"].apply(month_key_from_any)
    bad = raw["MonthKey"].isna().sum()
    if bad:
        print(f"WARNING: {bad} row(s) had an unparseable TXN Month and were dropped.")
    raw = raw.dropna(subset=["MonthKey"])
    raw["MonthKey"] = raw["MonthKey"].astype(int)

    if "Product Family" not in raw.columns:
        raw["Product Family"] = None
    raw["Product Family"] = raw["Product Family"].fillna("Other")
    raw.loc[raw["Product Family"].astype(str).str.strip().isin(["", "NA", "None"]), "Product Family"] = "Other"

    # Every column we ever group by must never be NaN — a NaN grouping key
    # produces a literal `NaN` token in the JSON output, which is valid
    # Python but INVALID per the JSON spec, so browsers refuse to parse the
    # file entirely (this broke the dashboard until caught by testing).
    fallback_label = {
        "Business Unit": "Unspecified",
        "Category": "Unspecified",
        "Item_Name": "Unspecified item",
        "Executive": "Unassigned",
        "Transaction Type": "Unspecified",
    }
    for col, label in fallback_label.items():
        if col not in raw.columns:
            raw[col] = label
            continue
        raw[col] = raw[col].fillna(label)
        blank_mask = raw[col].astype(str).str.strip().isin(["", "NA", "None", "nan"])
        raw.loc[blank_mask, col] = label

    raw["Total"] = pd.to_numeric(raw["Total"], errors="coerce").fillna(0.0)
    raw["GP AT Amount"] = pd.to_numeric(raw["GP AT Amount"], errors="coerce").fillna(0.0)

    def to_date_str(v):
        if hasattr(v, "strftime"):
            return v.strftime("%Y-%m-%d")
        return None
    raw["DateStr"] = raw["TXNDate"].apply(to_date_str)

    return raw


def group_sum(df, group_cols):
    g = df.groupby(group_cols, dropna=False).agg(
        Revenue=("Total", "sum"),
        GP=("GP AT Amount", "sum"),
        Txns=("Total", "size"),
    ).reset_index()
    return g


def to_records(df, cols):
    """Build the [[...], [...]] row lists the dashboard expects, in `cols` order."""
    records = []
    for _, r in df.iterrows():
        row = []
        for c in cols:
            v = r[c]
            if c in ("Revenue", "GP"):
                v = round(float(v), 2)
            elif c in ("Txns", "MonthKey"):
                v = int(v)
            row.append(v)
        records.append(row)
    return records


def build_centres(xls, raw):
    centres_in_raw = sorted(raw["Branch ID"].dropna().unique().tolist())

    loc_master = {}
    if "Location Master" in xls.sheet_names:
        lm = xls.parse("Location Master")
        lm.columns = [str(c).strip() for c in lm.columns]
        for _, r in lm.iterrows():
            centre = str(r.get("Centre Name", "")).strip()
            if not centre:
                continue
            loc_master[centre] = {
                "ARM": str(r.get("ARM", "")).strip() or "Unmapped",
                "State": str(r.get("State", "")).strip() or "Unmapped",
                "LocationType": str(r.get("Location Type", "")).strip() or "Service Centre",
            }
    else:
        print("WARNING: no 'Location Master' sheet found — all centres will be Unmapped.")

    # Most common Branch City per centre, from the raw data itself.
    city_by_centre = {}
    if "Branch City" in raw.columns:
        city_counts = raw.groupby(["Branch ID", "Branch City"]).size().reset_index(name="n")
        for centre, sub in city_counts.groupby("Branch ID"):
            top = sub.sort_values("n", ascending=False).iloc[0]
            city_by_centre[centre] = str(top["Branch City"])

    centres = []
    for centre in centres_in_raw:
        meta = loc_master.get(centre, {"ARM": "Unmapped", "State": "Unmapped", "LocationType": "Repair Drop Off"})
        centres.append({
            "Centre": centre,
            "State": meta["State"],
            "ARM": meta["ARM"],
            "LocationType": meta["LocationType"],
            "City": city_by_centre.get(centre, ""),
        })
    return centres


def melt_month_sheet(xls, sheet_name, value_scale=1.0):
    """For Revenue Target / GP Target / CSAT Target sheets: header row is
    row 4 (index 3), first column is Branch ID, remaining columns are months.
    Returns [[MonthKey, Centre, Value], ...] skipping blank cells.
    """
    if sheet_name not in xls.sheet_names:
        print(f"WARNING: sheet '{sheet_name}' not found — skipping.")
        return []
    df = xls.parse(sheet_name, header=3)
    if df.empty or df.columns[0] is None:
        return []
    df = df.rename(columns={df.columns[0]: "Centre"})
    out = []
    for col in df.columns[1:]:
        mk = month_key_from_any(col)
        if mk is None:
            continue
        for _, r in df.iterrows():
            centre = r["Centre"]
            val = r[col]
            if pd.isna(centre) or pd.isna(val):
                continue
            try:
                val = round(float(val), 2)
            except (TypeError, ValueError):
                continue
            out.append([mk, str(centre).strip(), val])
    return out


def build_csat(xls):
    if "CSAT" not in xls.sheet_names:
        print("WARNING: no 'CSAT' sheet found.")
        return []
    df = xls.parse("CSAT")
    if df.empty:
        return []
    df = df.rename(columns={df.columns[0]: "Centre", df.columns[1]: "Metric"})
    out = []
    for col in df.columns[2:]:
        mk = month_key_from_any(col)
        if mk is None:
            continue
        for _, r in df.iterrows():
            centre = r["Centre"]
            val = r[col]
            if pd.isna(centre) or pd.isna(val):
                continue
            try:
                pct = round(float(val) * 100.0, 2)
            except (TypeError, ValueError):
                continue
            out.append([mk, str(centre).strip(), pct])
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("workbook", help="Path to the master .xlsx file")
    ap.add_argument("--out", default="data.json", help="Output path for data.json")
    args = ap.parse_args()

    print(f"Reading {args.workbook} ...")
    xls = pd.ExcelFile(args.workbook)

    raw = load_raw_data(xls)
    print(f"Loaded {len(raw):,} raw transaction line(s) across "
          f"{raw['MonthKey'].nunique()} month(s) and {raw['Branch ID'].nunique()} centre(s).")

    fact_month = group_sum(raw, ["MonthKey", "Branch ID"]).rename(columns={"Branch ID": "Centre"})
    fact_bu = group_sum(raw, ["MonthKey", "Branch ID", "Business Unit"]).rename(
        columns={"Branch ID": "Centre", "Business Unit": "BU"})
    fact_cat = group_sum(raw, ["MonthKey", "Branch ID", "Category"]).rename(
        columns={"Branch ID": "Centre"})
    fact_pf = group_sum(raw, ["MonthKey", "Branch ID", "Product Family"]).rename(
        columns={"Branch ID": "Centre", "Product Family": "PF"})
    fact_item = group_sum(raw, ["MonthKey", "Branch ID", "Item_Name"]).rename(
        columns={"Branch ID": "Centre", "Item_Name": "Item"})
    fact_txntype = group_sum(raw, ["MonthKey", "Branch ID", "Transaction Type"]).rename(
        columns={"Branch ID": "Centre", "Transaction Type": "TxnType"})
    fact_exec = group_sum(raw, ["Branch ID", "Executive"]).rename(
        columns={"Branch ID": "Centre"})
    fact_day = group_sum(raw, ["DateStr", "Branch ID"]).rename(columns={"Branch ID": "Centre"})
    fact_day = fact_day.dropna(subset=["DateStr"])

    month_keys = sorted(fact_month["MonthKey"].unique().tolist())
    months = []
    for mk in month_keys:
        fy, q = fy_quarter(mk)
        months.append({
            "MonthKey": mk,
            "MonthLabel": month_label(mk),
            "FY": fy,
            "Quarter": q,
            "FYQ": f"{fy} {q}",
        })

    data = {
        "months": months,
        "gpCoveredMonths": month_keys,  # all months currently treated as GP-covered
        "centres": build_centres(xls, raw),
        "fact_month": to_records(fact_month, ["MonthKey", "Centre", "Revenue", "GP", "Txns"]),
        "fact_bu": to_records(fact_bu, ["MonthKey", "Centre", "BU", "Revenue", "GP", "Txns"]),
        "fact_cat": to_records(fact_cat, ["MonthKey", "Centre", "Category", "Revenue", "GP", "Txns"]),
        "fact_pf": to_records(fact_pf, ["MonthKey", "Centre", "PF", "Revenue", "GP", "Txns"]),
        "fact_item": to_records(fact_item, ["MonthKey", "Centre", "Item", "Revenue", "GP", "Txns"]),
        "fact_txntype": to_records(fact_txntype, ["MonthKey", "Centre", "TxnType", "Revenue", "GP", "Txns"]),
        "fact_exec": to_records(fact_exec, ["Centre", "Executive", "Revenue", "GP", "Txns"]),
        "fact_day": to_records(fact_day, ["DateStr", "Centre", "Revenue", "GP", "Txns"]),
        "csat": build_csat(xls),
        "target_revenue": melt_month_sheet(xls, "Revenue Target"),
        "target_gp": melt_month_sheet(xls, "GP Target"),
        "target_csat": melt_month_sheet(xls, "CSAT Target"),
    }

    import os
    os.makedirs(os.path.dirname(args.out) or ".", exist_ok=True)
    with open(args.out, "w", encoding="utf-8") as f:
        # allow_nan=False makes json.dump raise loudly here (during the build,
        # where it's visible in the GitHub Action log) instead of silently
        # writing an invalid `NaN` token that would break JSON.parse in every
        # browser and leave the dashboard stuck on its loading screen.
        json.dump(data, f, ensure_ascii=False, separators=(",", ":"), allow_nan=False)

    print(f"Wrote {args.out} "
          f"({sum(len(v) if isinstance(v, list) else 0 for v in data.values()):,} total rows across all tables).")


if __name__ == "__main__":
    main()
