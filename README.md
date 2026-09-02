# AMP AIM Sample & Technology Progress Dashboard

A static, interactive dashboard summarizing sample-collection and technology
progress across the AMP AIM cohort — built to run for free, 24/7, on GitHub
Pages, and updated from your weekly Excel export.

## Why this is aggregate-only (read this before publishing)

The source spreadsheet is human-subjects research data: it carries per-subject
diagnosis, disease-team, and biopsy severity fields alongside the sample-tracking
columns. That data should not go on a public website. This dashboard is built
so the **raw spreadsheet never gets near the public site or the public git
repo** — `build_data.py` reads it locally and writes out *only counts*
(`data/dashboard.json`): totals per technology, per disease/tissue scope, per
status. No subject IDs, no diagnoses, no demographics, no free-text notes.

**Never commit the raw `.xlsx` export to the public dashboard repo** — not even
temporarily. Git remembers every commit, so deleting it later doesn't remove it
from history. The included `.gitignore` blocks `*.xlsx`/`*.xlsm`/`*.csv` as a
safety net, but the real rule is: keep the spreadsheet somewhere private (your
own machine, a private repo, wherever you already store it) and only ever
publish the generated `dashboard.json`. If your data-sharing rules ever change,
re-check with your PI/data manager before changing what this dashboard shows.

## What's in this folder

```
index.html                  the dashboard itself (one self-contained file)
data/dashboard.json         generated aggregate data — this is what's public
build_data.py                reads your weekly .xlsx, writes dashboard.json
requirements.txt            Python dependency (openpyxl) for build_data.py
automation/                 optional: fully automated weekly refresh (see below)
```

## What's on each tab

- **Recruitment** — how many subjects are enrolled, broken down by disease team
  and then by cohort within each team. A subject's cohort is read from their
  enrollment visit row (`Visit_Code` = `V01` / `VC1` / `VE1` / `VA1` / `VU1` —
  one per visit type), pulling the `Visit_Cohort` tag from that row.
- **Technologies** — completion by technology, a full status breakdown
  (completed / pending / QC fail / not applicable), and a heatmap you can group
  either by disease/tissue scope (`Data_Scope`) or by cohort (`Visit_Cohort`) —
  toggle at the top right of the heatmap card.

**About "disease team":** the source `Disease Team` column is broken (reads
`#REF!` for most rows — a lookup formula pointing at something that no longer
exists). Both tabs instead derive the disease team from the `Data_Scope`
prefix (`SLE-KDY` → Lupus (SLE), `RA-SYN` → Rheumatoid Arthritis (RA), etc.) —
see `DISEASE_LABELS` and `derive_disease()` near the top of `build_data.py`.
If you have the authoritative disease-team list, or the true mapping differs
from this proxy, edit that dictionary (or send it to me and I'll wire it in).

## Quick start — get it live on GitHub Pages

1. Create a new **public** GitHub repo (e.g. `amp-aim-dashboard`).
2. Push everything in this folder **except your raw `.xlsx` file** — the
   `.gitignore` here already excludes it, so a plain `git add .` is safe.
   ```
   git init
   git add .
   git commit -m "Initial dashboard"
   git branch -M main
   git remote add origin https://github.com/<you>/amp-aim-dashboard.git
   git push -u origin main
   ```
3. In the repo settings: **Settings → Pages → Build and deployment → Source:
   Deploy from a branch → Branch: `main` / `(root)`**. Save.
4. GitHub gives you a URL like `https://<you>.github.io/amp-aim-dashboard/` —
   that's your live, public, 24/7 dashboard. It typically goes live within a
   minute or two.

No server, no backend, no cost.

## Weekly update (manual — takes about 2 minutes)

1. Save your latest export as `data/AMP_AIM_Dataset.xlsx` somewhere on your
   machine. It's fine to keep it inside this project folder for convenience —
   the `.gitignore` here already excludes `*.xlsx` from what git will track —
   just don't `git add -f` it or move it outside the ignored `data/` path.
2. Rebuild the data file:
   ```
   pip install -r requirements.txt
   python3 build_data.py /path/to/AMP_AIM_Dataset.xlsx data/dashboard.json
   ```
3. Commit and push just the updated data file:
   ```
   git add data/dashboard.json
   git commit -m "Weekly data refresh"
   git push
   ```
   The live site updates automatically within a minute or two of the push.

## Fully automated weekly refresh (optional)

If you'd rather not run a command by hand each week, `automation/private-repo-update-and-deploy.yml`
sets up a GitHub Action that does it for you — but it needs to live in a
**separate private repo** that holds your raw spreadsheet (never the public
one). See the comments at the top of that file for the one-time setup. Once
configured, updating the spreadsheet in the private repo is the entire weekly
workflow; the public dashboard refreshes itself.

## Using Render instead of GitHub Pages

Nothing here requires Render — this is a static site, and GitHub Pages hosts
it for free with less setup. Render is worth it only if you outgrow this
(e.g. you want the dashboard to query a live database instead of a weekly
snapshot). If you do want it on Render: create a **Static Site** service
pointing at this same repo, with build command `(none)` and publish directory
`.` — Render will serve `index.html` directly, same as GitHub Pages.

## Customizing

- **Which technologies are tracked**: edit the `TECH_COLUMNS` list at the top
  of `build_data.py` — each entry is `(source column name, short key, display label)`.
- **Status classification**: the `classify()` function in `build_data.py` maps
  the many raw status strings (`[specimen available]`, `[not applicable]`,
  QC-fail suffixes, etc.) into 5 reportable buckets. If your export starts
  using a new status string, add a rule there.
- **Disease team labels**: edit `DISEASE_LABELS` in `build_data.py` (see "About
  disease team" above).
- **Which visit codes count as "enrollment"**: edit `ENROLLMENT_CODES_PRIORITY`
  in `build_data.py` if a cohort should instead be read from a different visit
  code, or if the priority order between them should change.
- **Colors / chart styling**: all CSS custom properties are declared at the
  top of `index.html` (`:root` for light mode, the media query below it for
  dark mode) — change values there rather than hunting through the chart code.

## Local preview / troubleshooting

Opening `index.html` directly by double-clicking it will show a "could not
load data" message — browsers block `fetch()` on `file://` URLs. Preview it
with a local server instead:
```
python3 -m http.server 8000
```
then open `http://localhost:8000/`.
