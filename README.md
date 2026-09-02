# Aptronix Service — Executive Dashboard

A live, self-updating executive dashboard for Aptronix service centre
performance: revenue, gross profit, CSAT, licenses, targets, ARM/centre
breakdowns, and AI-generated insights — all driven from one master Excel
workbook.

**Live site:** `https://<your-username>.github.io/<repo-name>/` (set up in
"Publish it on GitHub Pages" below).

---

## How self-updating works

```
source/master.xlsx  →  scripts/build_data.py  →  data.json  →  index.html
   (you edit this)      (runs automatically)     (generated)   (reads this)
```

1. You update `source/master.xlsx` with new transactions, targets, or CSAT
   scores — the same workbook, same sheet names, no new format to learn —
   and push it to this repo (upload via GitHub's web UI, or `git push`).
2. That push automatically triggers a **GitHub Action**
   (`.github/workflows/update-dashboard.yml`), which runs
   `scripts/build_data.py` against the workbook and regenerates `data.json`.
3. The Action commits the new `data.json` back to the repo.
4. GitHub Pages redeploys automatically on every commit, so the live
   dashboard picks up the new data — typically within 1–2 minutes of your
   upload, no manual rebuild step required.

You can also trigger a rebuild manually any time from the repo's
**Actions** tab → *Rebuild dashboard data* → **Run workflow** — useful if
you've edited `source/master.xlsx` some other way and want to force a
refresh.

### Why it's split this way

The dashboard used to have every number typed directly into the HTML file
by hand. Now `index.html` is just the display logic — it fetches
`data.json` at load time instead. `data.json` is never hand-edited; it's
always a mechanical, reproducible output of `build_data.py` run against
`source/master.xlsx`. Updating the dashboard is now "replace the Excel
file," not "re-paste a decade of numbers into HTML."

---

## Repository structure

```
index.html                          the dashboard (fetches data.json at load)
data.json                           generated — do not hand-edit
source/master.xlsx                  the master workbook — edit this
scripts/build_data.py               the Excel → data.json converter
.github/workflows/update-dashboard.yml   auto-rebuild on every workbook push
```

## What's in `source/master.xlsx`

The build script reads these sheets:

| Sheet | Used for |
|---|---|
| Any sheet starting with **"Raw"** (`Raw Data 24-25`, `Raw Data 2025-26`, `Rawdata 26-27`, …) | All transaction-level data — revenue, GP, business unit, category, item, executive, day, txn type |
| **Location Master** | Maps each centre to its ARM, State, and Location Type |
| **CSAT** | Monthly CSAT % per centre |
| **Revenue Target**, **GP Target**, **CSAT Target** | Monthly targets per centre |

Two sheets (`Revenue `, `GP`) hold legacy FY22-23 reference data and are
**not** read by the build script — the dashboard's data starts from
whatever the "Raw…" sheets cover.

### Adding a new fiscal year

Just add a new sheet named `Raw Data <whatever>` (as long as the name
starts with "Raw") with the same columns as the existing ones, and push.
The build script picks up *every* sheet matching that pattern automatically
— you don't need to touch `build_data.py` or the workflow file.

### Adding a new centre

Add a row to **Location Master** with its ARM, State, and Location Type.
If a centre appears in the raw data but not in Location Master, the build
script still includes it — with State/ARM set to `Unmapped` and Location
Type defaulted to `Repair Drop Off` — so nothing breaks, but you'll want to
add the real mapping when you get a chance.

---

## Running the build yourself (optional)

You don't need to do this — the GitHub Action does it automatically on
every push to `source/master.xlsx`. But if you want to check the output
locally before pushing:

```bash
pip install pandas openpyxl
python scripts/build_data.py source/master.xlsx --out data.json
```

The script prints a summary (rows loaded, months found, centres found) and
any warnings (e.g. unparseable month columns) so you can sanity-check
before committing.

---

## Publish it on GitHub Pages

1. Create a new GitHub repository (public — GitHub Pages on a free
   personal account only serves public repos).
2. Push everything in this folder to that repository (root of the repo,
   not a subfolder) — either drag-and-drop upload via the GitHub web UI,
   or:
   ```bash
   git init
   git remote add origin https://github.com/<your-username>/<repo-name>.git
   git add .
   git commit -m "Initial dashboard"
   git branch -M main
   git push -u origin main
   ```
3. In the repo, go to **Settings → Pages**.
4. Under **Build and deployment → Source**, choose **Deploy from a branch**.
5. Under **Branch**, select `main` and folder `/ (root)`, then **Save**.
6. GitHub publishes at `https://<your-username>.github.io/<repo-name>/`
   within a minute or two of the first push.

From then on: **edit `source/master.xlsx`, push it, wait a minute or two —
the live dashboard updates itself.**

---

## Known limitations / things worth knowing

- **No access control.** A public GitHub Pages repo has no login — anyone
  with the link can view the dashboard and, via browser dev tools, the
  underlying `data.json`. Don't publish this way if the numbers need to
  stay private. (We discussed real access-control options earlier — pre-split
  builds behind Cloudflare Access, or a proper backend — if you want to
  revisit that.)
- **`data.json` loads via a synchronous request**, which needs a real
  HTTP(S) server — exactly what GitHub Pages is. It will *not* work if you
  double-click `index.html` and open it straight from your hard drive
  (`file://`); you'll see a "could not load data.json" message on the
  loading screen in that case. Always test through the published URL (or
  a local static server like `python -m http.server`), not by opening the
  file directly.
- **"Product Family" data quality.** That column in the raw sheets is
  inconsistently filled (sometimes blank, sometimes has an item code
  pasted in by mistake). The build script buckets blanks into "Other" —
  functional, but the fact_pf breakdown will only get cleaner if the
  source column does.
- **GP/transaction-count drift on historical months.** When this
  automation was first validated against the previous hand-built
  dashboard, revenue matched exactly for all 721 month×centre
  combinations checked; a small number of GP figures and transaction
  counts differed slightly, traced to the underlying workbook having been
  edited since that earlier snapshot was taken — not a bug in the
  conversion logic.
- **Chart.js and Lucide icons load from a CDN** with automatic fallbacks
  (cdnjs → jsDelivr → unpkg). If a viewer's network blocks all three,
  charts fall back to a text notice and nav icons fall back to plain
  glyphs — the dashboard itself still works.
- **Dark mode preference** is stored per-browser via `localStorage`, so it
  persists on that device across visits but doesn't carry across devices.
- Fully responsive: sidebar becomes a slide-out drawer, tables scroll
  horizontally, and the filter bar reflows to one column below ~480px.

## Updating the dashboard's *code* (not the data)

Data changes → just edit `source/master.xlsx` and push, as above.

Code/design changes (new features, styling, new views) → edit `index.html`
directly and push; GitHub Pages redeploys on every commit the same way.
