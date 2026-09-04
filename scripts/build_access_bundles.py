#!/usr/bin/env python3
"""
build_access_bundles.py — generates one dashboard bundle per access tier/
scope, each containing only the centres that role is allowed to see.

USAGE
    python scripts/build_access_bundles.py source/master.xlsx index.html \
        --out-dir dist --settings access/settings.csv --roles access/roles.csv

WHAT IT READS
    - access/settings.csv: Tier,ShowGP            (currently: all TRUE — every
      tier sees GP; this stays here so it's a one-line change later if that
      policy ever changes, not a rebuild)
    - access/roles.csv:    Email,Tier,Scope         (Scope = ARM name for
      "Area Manager", exact Centre name for "Centre Manager", blank for
      "Admin")

WHAT IT WRITES
    dist/
      _bundle-map.csv         — slug, tier, scope, path, allowed emails
                                 (reference for setting up Cloudflare Access
                                 policies — not served to the site)
      admin/index.html + data.json
      area-<slug>/index.html + data.json     (one per distinct ARM in roles.csv)
      centre-<slug>/index.html + data.json   (one per distinct centre in roles.csv)

Each bundle's data.json contains ONLY rows for centres that tier/scope is
allowed to see — this is real filtering at build time, not something the
browser could be tricked into revealing more of. GP is included or excluded
per access/settings.csv (currently included everywhere).
"""
import argparse
import csv
import json
import re
import shutil
from pathlib import Path

import pandas as pd

import build_data


def slugify(s):
    s = s.strip().lower()
    s = re.sub(r"[^a-z0-9]+", "-", s)
    return s.strip("-")


def read_settings(path):
    settings = {}
    with open(path, newline="", encoding="utf-8") as f:
        for row in csv.DictReader(f):
            tier = row["Tier"].strip()
            show_gp = row["ShowGP"].strip().upper() == "TRUE"
            settings[tier] = show_gp
    return settings


def read_roles(path):
    roles = []
    with open(path, newline="", encoding="utf-8") as f:
        for row in csv.DictReader(f):
            email = row["Email"].strip()
            tier = row["Tier"].strip()
            scope = row["Scope"].strip()
            if not email or not tier:
                continue
            if tier not in ("Admin", "Area Manager", "Centre Manager"):
                print(f"WARNING: skipping role row with unknown Tier '{tier}' for {email}")
                continue
            if tier != "Admin" and not scope:
                print(f"WARNING: skipping {tier} row for {email} — Scope is required "
                      f"(ARM name for Area Manager, Centre name for Centre Manager).")
                continue
            roles.append({"email": email, "tier": tier, "scope": scope})
    return roles


CENTRE_COL_INDEX = {
    "fact_month": 1, "fact_bu": 1, "fact_cat": 1, "fact_pf": 1,
    "fact_item": 1, "fact_txntype": 1, "fact_day": 1,
    "fact_exec": 0,
    "csat": 1, "target_revenue": 1, "target_gp": 1, "target_csat": 1,
}


def scope_data(master, allowed_centres):
    """Returns a copy of master data restricted to allowed_centres (a set).
    allowed_centres=None means no restriction (Admin — everything)."""
    if allowed_centres is None:
        return dict(master)
    out = dict(master)
    out["centres"] = [c for c in master["centres"] if c["Centre"] in allowed_centres]
    for key, idx in CENTRE_COL_INDEX.items():
        out[key] = [row for row in master[key] if row[idx] in allowed_centres]
    return out


def apply_gp_visibility(data, show_gp):
    data = dict(data)
    data["gpVisible"] = bool(show_gp)
    return data


def write_bundle(out_dir, slug, data, index_html_path):
    bundle_dir = out_dir / slug
    bundle_dir.mkdir(parents=True, exist_ok=True)
    with open(bundle_dir / "data.json", "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, separators=(",", ":"), allow_nan=False)
    shutil.copyfile(index_html_path, bundle_dir / "index.html")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("workbook", help="Path to the master .xlsx file")
    ap.add_argument("index_html", help="Path to the dashboard's index.html")
    ap.add_argument("--out-dir", default="dist", help="Output directory for all bundles")
    ap.add_argument("--settings", default="access/settings.csv")
    ap.add_argument("--roles", default="access/roles.csv")
    args = ap.parse_args()

    out_dir = Path(args.out_dir)
    if out_dir.exists():
        shutil.rmtree(out_dir)
    out_dir.mkdir(parents=True)

    print(f"Reading {args.workbook} ...")
    xls = pd.ExcelFile(args.workbook)
    master = build_data.build_master_data(xls)

    settings = read_settings(args.settings)
    roles = read_roles(args.roles)

    centre_arm = {c["Centre"]: c["ARM"] for c in master["centres"]}
    all_centre_names = set(centre_arm.keys())

    # Group roles into unique (tier, scope) bundles — several emails can
    # share one bundle (e.g. two people managing the same centre).
    bundles = {}  # (tier, scope) -> set of emails
    for r in roles:
        key = (r["tier"], r["scope"])
        bundles.setdefault(key, set()).add(r["email"])

    if not any(t == "Admin" for t, s in bundles):
        print("WARNING: no Admin rows found in access/roles.csv — add at least one "
              "so someone can actually reach the full dashboard.")

    bundle_map_rows = []
    for (tier, scope), emails in sorted(bundles.items()):
        if tier == "Admin":
            allowed = None
            slug = "admin"
        elif tier == "Area Manager":
            allowed = {c for c, arm in centre_arm.items() if arm == scope}
            if not allowed:
                print(f"WARNING: no centres found with ARM == '{scope}' "
                      f"(check spelling against Location Master) — skipping this bundle.")
                continue
            slug = "area-" + slugify(scope)
        else:  # Centre Manager
            if scope not in all_centre_names:
                print(f"WARNING: '{scope}' doesn't match any centre name exactly "
                      f"(check spelling against Location Master) — skipping this bundle.")
                continue
            allowed = {scope}
            slug = "centre-" + slugify(scope)

        show_gp = settings.get(tier, True)
        data = scope_data(master, allowed)
        data = apply_gp_visibility(data, show_gp)
        write_bundle(out_dir, slug, data, args.index_html)

        n_centres = len(allowed) if allowed is not None else len(all_centre_names)
        print(f"Built /{slug}/  tier={tier}  scope={scope or '(all)'}  "
              f"centres={n_centres}  gpVisible={show_gp}  emails={sorted(emails)}")
        bundle_map_rows.append({
            "slug": slug, "tier": tier, "scope": scope,
            "path": f"/{slug}/*", "emails": ";".join(sorted(emails)),
        })

    with open(out_dir / "_bundle-map.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=["slug", "tier", "scope", "path", "emails"])
        w.writeheader()
        w.writerows(bundle_map_rows)

    print(f"\nWrote {len(bundle_map_rows)} bundle(s) to {out_dir}/")
    print(f"See {out_dir}/_bundle-map.csv for the exact paths and emails to use "
          f"when setting up Cloudflare Access policies.")


if __name__ == "__main__":
    main()
