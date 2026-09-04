# Aptronix Service — Executive Dashboard

A self-updating, access-controlled executive dashboard for Aptronix service
centre performance: revenue, gross profit, CSAT, licenses, targets,
ARM/centre breakdowns, and AI-generated insights — all driven from one
master Excel workbook, with genuinely separate data per access tier.

**Hosting:** Cloudflare Pages + Cloudflare Access (both free at this scale).
Set up once per "Publish it" below — everything after that is automatic.

---

## How it works, end to end

```
source/master.xlsx  ──┐
access/settings.csv ──┼──►  scripts/build_access_bundles.py  ──►  dist/<tier>/  ──►  Cloudflare Pages
access/roles.csv    ──┘         (runs automatically in CI)         (one bundle          (one URL path
                                                                      per role)            per bundle,
                                                                                            gated by
                                                                                          Cloudflare Access)
```

1. You edit `source/master.xlsx` (transactions/targets/CSAT), `access/roles.csv`
   (who's in which tier), or `access/settings.csv` (GP visibility per tier) —
   and push.
2. A GitHub Action (`.github/workflows/deploy-cloudflare.yml`) rebuilds
   **every** access-tier bundle fresh from scratch and deploys the result to
   Cloudflare Pages. Nothing is cached between runs, so there's no way for a
   bundle to go stale.
3. Cloudflare Access sits in front of every bundle's URL path and only lets
   through the email addresses you've assigned to that tier/scope. Someone
   without permission never receives that bundle's `data.json` at all — this
   is real separation, not a hidden URL or a UI toggle.

## Access tiers

| Tier | Sees | Scope column in `access/roles.csv` |
|---|---|---|
| **Admin** | Every centre | (leave blank) |
| **Area Manager** | Only centres under their ARM | The ARM name, exactly as it appears in the workbook's Location Master sheet |
| **Centre Manager** | Only their one centre | The exact centre name, as in Location Master |

Each unique (tier, scope) combination becomes one URL path:
- Admin → `/admin/`
- Area Manager for "Naveen Sukka" → `/area-naveen-sukka/`
- Centre Manager for "Service Abids" → `/centre-service-abids/`

Multiple people can share one bundle (e.g. two people managing the same
centre) — just add both as separate rows in `access/roles.csv` with the same
Tier and Scope; they'll both be granted access to the same URL in Cloudflare
Access.

### `access/roles.csv`

```
Email,Tier,Scope
owner@example.com,Admin,
you@example.com,Admin,
naveen.sukka@example.com,Area Manager,Naveen Sukka
manager.abids@example.com,Centre Manager,Service Abids
```

Replace the example rows with your real people before going live — the
placeholder `@example.com` addresses won't match anyone.

### `access/settings.csv` — the GP toggle

```
Tier,ShowGP
Admin,TRUE
Area Manager,TRUE
Centre Manager,TRUE
```

Currently every tier sees GP. Flip any row to `FALSE` and push to hide GP
figures for that tier — same rebuild-and-deploy flow as any other change,
live in about a minute. (If you do this, the underlying number is genuinely
removed from that tier's `data.json`, not just hidden in the browser — real
security, not a UI toggle.)

---

## Publish it: one-time setup

### 1. Cloudflare Pages (hosting)

1. Sign up at dash.cloudflare.com (free).
2. **Workers & Pages → Create application → Pages → Get started → Upload assets**
   (**not** "Connect to Git" — this repo's GitHub Action does the building
   and pushes the finished result directly, so Cloudflare doesn't need to
   build anything itself).
3. Name the project `aptronix-dashboard` (or update `projectName` in
   `.github/workflows/deploy-cloudflare.yml` to match whatever you name it),
   then drag in any single placeholder file just to finish creating the
   project — the GitHub Action will overwrite it with the real bundles on
   its first run. Select **Deploy site**.
4. Get your credentials for the next step:
   - **Account ID**: shown on the right sidebar of any page in the Cloudflare
     dashboard.
   - **API Token**: **My Profile → API Tokens → Create Token → Custom Token
     → Account, Cloudflare Pages, Edit → Continue to summary → Create Token**.
5. In this GitHub repo: **Settings → Secrets and variables → Actions**, add:
   - `CLOUDFLARE_API_TOKEN`
   - `CLOUDFLARE_ACCOUNT_ID`
6. Push anything (or re-run the workflow manually from the **Actions** tab)
   — the Action builds every bundle and deploys to
   `https://aptronix-dashboard.pages.dev/<tier-slug>/`.

### 2. Cloudflare Access (who can see what)

This is a **separate Cloudflare product** (Zero Trust) — free for up to 50
users.

1. In the Cloudflare dashboard, open **Zero Trust** from the left nav. The
   first time, it asks you to pick a **team name** and confirm the **Free**
   plan.
2. **Zero Trust → Settings → Authentication → Login methods → Add new →
   One-time PIN.** This is no longer on by default for new accounts — skip
   it and nobody can ever log in, since no login email gets sent.
3. **Zero Trust → Access → Applications → Add an application → Self-hosted.**
4. After the first deploy, open `dist/_bundle-map.csv` (generated by the
   build, not committed to git — download it from the Action's run summary,
   or run the build script locally) for the exact path and email list to use
   for each application. For each row:
   - **Application domain**: `aptronix-dashboard.pages.dev`
   - **Path**: the `path` column (e.g. `/centre-service-abids/*`)
   - **Policy**: Allow → Include → **Emails** → paste the addresses from the
     `emails` column
5. Repeat for every row in `_bundle-map.csv`, including `/admin/*`.

Once this is set up, send each person the specific URL for their tier (e.g.
`https://aptronix-dashboard.pages.dev/centre-service-abids/`). Visiting it
prompts them to verify their email before they see anything.

---

## What's in `source/master.xlsx`

The build script reads these sheets:

| Sheet | Used for |
|---|---|
| Any sheet starting with **"Raw"** (`Raw Data 24-25`, `Raw Data 2025-26`, `Rawdata 26-27`, …) | All transaction-level data — revenue, GP, business unit, category, item, executive, day, txn type |
| **GP** | Overrides the headline GP figure with Apple's spare-part incentive included (see "How Gross Profit is calculated" below) |
| **Location Master** | Maps each centre to its ARM, State, and Location Type — also what `access/roles.csv` scopes are validated against |
| **CSAT** | Monthly CSAT % per centre |
| **Revenue Target**, **GP Target**, **CSAT Target** | Monthly targets per centre |

One sheet (`Revenue ` — note the trailing space in its name) holds legacy
FY22-23 reference data from before this dashboard existed and is **not**
read by the build script.

**This workbook is never opened and re-saved by any script here** — only
read. (Earlier in building this, adding sheets directly to it via a
load-then-save round-trip silently corrupted formula-derived date columns
in two of the raw-data sheets. Access configuration lives in the separate
`access/*.csv` files specifically to avoid ever touching this file's
contents again.)

### How Gross Profit is calculated

Two different figures feed into "GP" depending on which part of the
dashboard you're looking at:

- **Transaction GP** — summed straight from each raw sale's `GP AT Amount`
  column. This is what every breakdown (by business unit, category, item,
  executive, centre-explorer drill-down, day) is built from.
- **Incentive-inclusive GP** — the **GP** sheet additionally includes
  Apple's spare-part incentive, which isn't tied to any individual
  transaction (Apple pays it as a periodic lump sum per centre, not per
  repair). The **GP** sheet is a hand-maintained report: one block per
  fiscal year (new blocks get added below the previous ones as years
  pass), each headed by a row whose first cell is `Branch ID` or `Centre
  Name`, with 12 monthly columns per centre. `build_data.py` auto-detects
  however many blocks exist — a new fiscal year needs no code change, just
  a new block appended to the sheet, in the same header-then-centre-rows
  shape as the existing ones.

  Wherever that sheet has a figure for a given centre/month, it **replaces**
  the transaction-derived GP for:
  - `fact_month` (drives Executive Overview, the GP tab, and every trend
    chart)
  - `fact_day` (drives the "exact" KPI totals, which sum from daily data
    for date-range precision) — the month's incentive is added onto that
    centre's **1st of the month** entry, since there's no data saying
    which day it belongs to. Any FY, quarter, or full-month filter
    captures it correctly; a custom date range that excludes day 1 of a
    month won't.

  Every other breakdown (by business unit, category, product family, item,
  transaction type, or executive) is **not** adjusted — there's no basis
  for attributing a lump-sum incentive to a specific category or associate.
  So on months with a real incentive, you'll correctly see the headline GP
  run higher than what those breakdowns sum to on their own; that gap *is*
  the incentive. `gpCoveredMonths` in each bundle's `data.json` lists exactly
  which months the GP sheet had a figure for, which is what the dashboard's
  "GP coverage" indicator is built from.

### Adding a new fiscal year

Just add a new sheet named `Raw Data <whatever>` (as long as the name
starts with "Raw") with the same columns as the existing ones, and push.

---

## Known limitations / things worth knowing

- **Cloudflare Access gates the page, not the download.** Once someone is
  authenticated, their browser does receive their tier's full `data.json` —
  that's unavoidable for a client-rendered dashboard. What Access actually
  guarantees is that *only* their tier's data (already filtered at build
  time) ever reaches them; someone in Centre Manager tier can inspect their
  own centre's numbers in dev tools, but never another centre's, because
  those numbers were never sent to their browser in the first place.
- **`access/roles.csv` scope values must match Location Master exactly**
  (same spelling, same case) — the build script prints a warning and skips
  a role row if an Area Manager's ARM name or a Centre Manager's centre name
  doesn't match anything, rather than guessing.
- **Casing/whitespace normalization.** Text values in `Business Unit`
  and `Category` weren't spelled identically across fiscal years (e.g.
  "Incident Fees" vs "Incident fees"). Earlier this was treated as
  cosmetic, but it isn't — it silently splits one real category into two,
  which shows up as fake zero-gaps in trend charts wherever the other
  casing was used that month. `build_data.py` now collapses any column
  used for grouping (Business Unit, Category, Product Family, Item,
  Executive, Transaction Type) to its most-frequent spelling wherever case
  or leading/trailing whitespace is the only difference, and prints how
  many distinct values it merged so you can sanity-check after adding a
  new year's sheet.
- **"Product Family" data quality.** That column in the raw sheets is
  inconsistently filled (sometimes blank, sometimes has an item code
  pasted in by mistake). The build script buckets blanks into "Other" —
  functional, but the fact_pf breakdown will only get cleaner if the
  source column does.
- **Chart.js and Lucide icons load from a CDN** with automatic fallbacks
  (cdnjs → jsDelivr → unpkg). If a viewer's network blocks all three,
  charts fall back to a text notice and nav icons fall back to plain
  glyphs — the dashboard itself still works.
- **Dark mode preference** is stored per-browser via `localStorage`, so it
  persists on that device across visits but doesn't carry across devices.
- Fully responsive: sidebar becomes a slide-out drawer, tables scroll
  horizontally, and the filter bar reflows to one column below ~480px.

## Running the build yourself (optional)

You don't need to — the GitHub Action does this automatically on every
relevant push. To check locally before pushing:

```bash
pip install pandas openpyxl
python scripts/build_access_bundles.py source/master.xlsx index.html --out-dir dist
```

Prints a summary per bundle (tier, scope, centre count, GP visibility,
which emails) and writes `dist/_bundle-map.csv` for setting up Cloudflare
Access policies.
