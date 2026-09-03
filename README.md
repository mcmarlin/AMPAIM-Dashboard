# AMP AIM Sample & Technology Progress Dashboard

A static, interactive dashboard summarizing sample-collection and technology
progress across the AMP AIM cohort — built to run for free, 24/7, on GitHub
Pages, and updated from your weekly Excel export.

**Live at:** https://mcmarlin.github.io/AMPAIM-Dashboard/
**Repo:** https://github.com/mcmarlin/AMPAIM-Dashboard

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
index.html                             the dashboard itself (one self-contained file)
data/dashboard.json                    generated aggregate data — this is what's public
data/Expected_Recruitment_Numbers.xlsx recruitment targets — edit this to update them
build_data.py                          reads your weekly .xlsx, writes dashboard.json
requirements.txt                       Python dependency (openpyxl) for build_data.py
automation/                            optional: fully automated weekly refresh (see below)
```

## What's on each tab

- **Recruitment** — how many subjects are enrolled, broken down by disease
  team/subgroup and then by cohort within each. A subject's cohort is read
  from their enrollment visit row (`Visit_Code` = `V01` / `VC1` / `VE1` /
  `VA1` / `VU1` — one per visit type), pulling the `Visit_Cohort` tag from
  that row. Bars show progress toward the network's Expected Recruitment
  targets; the top chart also splits Lupus (Kidney/Skin) and Psoriatic
  Disease (Arthritis/Psoriasis/Uveitis) into colored segments, and cohort
  bars are further split into Enrolled vs. Archival subjects.
- **Technologies** — completion by technology, a full status breakdown
  (completed / pending / QC fail / not applicable), and a small chart per
  disease team showing completed samples by technology, filterable by
  disease team, Pipeline, and dataset.

**About "disease team":** the source `Disease Team` column is broken (reads
`#REF!` for most rows — a lookup formula pointing at something that no longer
exists). Both tabs instead derive the disease team from the `Data_Scope`
value (`SLE-KDY` → Lupus Kidney, `RA-SYN` → Rheumatoid Arthritis (RA), etc.) —
see `DISEASE_LABELS`, `SPLIT_DISEASE_LABELS`, and `derive_diseases()` near the
top of `build_data.py`. If you have the authoritative disease-team list, or
the true mapping differs from this proxy, send it over and I'll wire it in.

## Quick start — get it live on GitHub Pages

This is already set up for the `mcmarlin/AMPAIM-Dashboard` repo, live at
https://mcmarlin.github.io/AMPAIM-Dashboard/. To (re)deploy it from scratch:

1. Create a new **public** GitHub repo named `AMPAIM-Dashboard` under the
   `mcmarlin` account.
2. Push everything in this folder **except your raw `.xlsx` file** — the
   `.gitignore` here already excludes it, so a plain `git add .` is safe.
   ```
   git init
   git add .
   git commit -m "Initial dashboard"
   git branch -M main
   git remote add origin https://github.com/mcmarlin/AMPAIM-Dashboard.git
   git push -u origin main
   ```
3. In the repo settings: **Settings → Pages → Build and deployment → Source:
   Deploy from a branch → Branch: `main` / `(root)`**. Save.
4. GitHub serves it at https://mcmarlin.github.io/AMPAIM-Dashboard/ — that's
   the live, public, 24/7 dashboard. It typically goes live within a minute
   or two of enabling Pages.

No server, no backend, no cost.

## Weekly update (manual — takes about 2 minutes)

There are two folders involved, and it's easy to mix them up:

- **Your project folder** — the local copy of this `AMPAIM-Dashboard` repo on
  your computer (the folder that has `index.html`, `README.md`, `build_data.py`,
  etc. in it — wherever you ran `git clone` or downloaded it to). Everything
  below happens from inside this folder.
- **Wherever you save the weekly export** — this can be anywhere on your
  computer. It does **not** need to be inside the project folder. The steps
  below use the project folder's own `data/` subfolder for convenience (and
  because it's already set up to be ignored by git), but any location works
  as long as you use the right path in step 2.

Steps:

1. **Save the file.** Save this week's export as
   `data/AMP AIM Dataset.xlsx` *inside your project folder* (i.e., save it
   into the `data` subfolder that's already there, replacing last week's
   copy) — note the file is always named **with spaces**, not underscores.
   This file will never be committed to git — the `.gitignore` in this
   folder already excludes `*.xlsx` — so it's safe to just keep overwriting it
   here each week.
2. **Open a terminal in your project folder and run the build script.** On
   Mac, open Terminal; on Windows, open Command Prompt or PowerShell. Then
   navigate into the project folder (for example
   `cd Documents/AMPAIM-Dashboard`, adjusted to wherever yours actually is),
   and run:
   ```
   python3 -m pip install -r requirements.txt
   python3 build_data.py "data/AMP AIM Dataset.xlsx" data/dashboard.json
   ```
   The quotes around the first path are required — the filename has spaces
   in it, and without quotes the terminal treats it as three separate
   arguments and the script won't find the file.
   (The `pip install` line only needs to be run once, the first time — you
   can skip it on later weeks. Using `python3 -m pip install...` rather than
   just `pip install...` matters especially on Windows, where a plain `pip`
   command can sometimes point at a different Python install than `python3`
   does — running pip "through" python3 like this guarantees the package
   lands where `build_data.py` will actually look for it. If you already ran
   `pip install -r requirements.txt` and still see
   `ModuleNotFoundError: No module named 'openpyxl'`, that mismatch is almost
   certainly why — rerun it as `python3 -m pip install -r requirements.txt`
   instead.) This reads the spreadsheet and
   overwrites `data/dashboard.json` with the new aggregate counts. If you
   saved the export somewhere other than `data/AMP AIM Dataset.xlsx`,
   replace that first (quoted) path with wherever you actually saved it —
   keep the quotes if that path has spaces too.
3. **Commit and push, from that same terminal, still inside the project
   folder:**
   ```
   git add data/dashboard.json
   git commit -m "Weekly data refresh"
   git push
   ```
   The live site updates automatically within a minute or two of the push.
   (Only `dashboard.json` gets pushed — the `.xlsx` file stays on your
   computer and is never uploaded anywhere.)

## Updating recruitment targets

The "of N" progress numbers on the Recruitment tab (e.g. "192 of 488") come
from `data/Expected_Recruitment_Numbers.xlsx`, not from the weekly sample
export — these targets change far less often, so they're kept in their own
small spreadsheet. To update one:

1. Open `data/Expected_Recruitment_Numbers.xlsx` (in your project folder,
   right next to `data/AMP AIM Dataset.xlsx`) in Excel.
2. Each row is one cohort. Edit the number in the **Expected** column, or
   type `Undefined` if a target isn't set yet. (The **Notes** column is just
   for your own reference — it isn't read by the dashboard.) Don't rename the
   **Disease Team** or **Cohort** values, or a target won't be found for that
   cohort.
3. Save the file, in place, still named `Expected_Recruitment_Numbers.xlsx`.
4. Run the normal weekly-update steps above (`build_data.py` then
   `git add` / `commit` / `push`) — the new targets are picked up
   automatically the next time the script runs, no separate step needed. If
   you're only updating targets and not the weekly sample export, you can
   still just rerun step 2 and step 3 of the weekly update as-is; the script
   reads both files every time.

If this file is ever missing or moved, `build_data.py` won't fail — it just
prints a warning and every recruitment bar shows "not yet set" until the
file is back in place.

## Fully automated weekly refresh (optional)

If you'd rather not run a command by hand each week, `automation/private-repo-update-and-deploy.yml`
sets up a GitHub Action that does it for you — but it needs to live in a
**separate private repo** that holds your raw spreadsheet (never the public
one). See the comments at the top of that file for the one-time setup. Once
configured, updating the spreadsheet in the private repo is the entire weekly
workflow; the public dashboard refreshes itself.

## Local preview / troubleshooting

Opening `index.html` directly by double-clicking it will show a "could not
load data" message — browsers block `fetch()` on `file://` URLs. Preview it
with a local server instead:
```
python3 -m http.server 8000
```
then open `http://localhost:8000/`.
