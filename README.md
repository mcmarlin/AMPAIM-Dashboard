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

**Never upload the raw `.xlsx` export to the public dashboard repo on
GitHub** — not even temporarily. GitHub remembers every version of every file
ever uploaded, so deleting it afterward doesn't remove it from the repo's
history (see "Removing a file you accidentally uploaded" below — a plain
delete isn't enough for anything sensitive). The real safeguard is keeping
both spreadsheets (`AMP-AIM_Dataset_Weekly_Update.xlsx` and `Target_Recruitment_Numbers.xlsx`)
in your `Data` folder — see "Where your files live" below — which sits
*outside* the project folder you upload from, so there's nothing to
accidentally drag in. Only ever upload the generated `dashboard.json`. If
your data-sharing rules ever change, re-check with your PI/data manager
before changing what this dashboard shows.

## What's in this folder

```
index.html            the dashboard itself (one self-contained file)
data/dashboard.json   generated aggregate data — this is what's public, and all that's public
build_data.py         reads your two spreadsheets, writes dashboard.json
requirements.txt      Python dependency (openpyxl) for build_data.py
automation/           optional: fully automated weekly refresh (see below)
```

## Where your files live

Your two source spreadsheets live in a **`Data` folder that sits next to this
project folder** (i.e., both are subfolders of the same parent —
`AMP-AIM_Progress_Dashboard\Data\` and `AMP-AIM_Progress_Dashboard\amp-dashboard_v4\`,
adjusted to whatever your project folder is actually named):

```
AMP-AIM_Progress_Dashboard\
├── Data\
│   ├── AMP-AIM_Dataset_Weekly_Update.xlsx   this week's export (subjects + visits + more)
│   └── Target_Recruitment_Numbers.xlsx      recruitment targets — edit to update them
└── amp-dashboard_v4\  (or whatever your project folder is named)
    ├── index.html
    ├── build_data.py
    ├── data\dashboard.json
    └── ...
```

Keeping `Data` outside the project folder is deliberate — it means your two
spreadsheets never sit anywhere near what you upload to GitHub, so there's no
risk of dragging one in by mistake, and you don't have to re-save them into a
new folder every time you get an updated project folder from Claude.

## What's on each tab

- **Recruitment** — how many subjects are enrolled, broken down by disease
  team/subgroup and then by cohort within each. This now reads straight off
  the `subjects` tab of the weekly export — one row per subject, no more
  hunting for a particular enrollment visit. A subject's status (Enrolled vs.
  Archival) comes from their `Subject Type`, and their cohort(s) come from
  their `Dashboard_Label` tag(s) — a simplified, harmonized version of the
  original cohort labels, still shown as "Cohort" in the dashboard. Bars show
  progress toward the network's recruitment targets; the top chart also
  splits Lupus (Kidney/Skin) and Psoriatic Disease (Arthritis/Psoriasis) into
  colored segments, and cohort bars are further split into Enrolled vs.
  Archival subjects.
- **Technologies** — completion by technology, a full status breakdown
  (completed / pending / QC fail / not applicable), and a small chart per
  disease team showing completed samples by technology, filterable by
  disease team, Pipeline, and dataset. This still reads the visit-level
  `visits` tab, since it's tracking samples/assays per visit rather than
  per subject. The dataset list includes **Post-Xenium H&E** and
  **Post Xenium IMC** (from the `Post-Xenium_HE` / `Post-Xenium_IMC`
  columns), shown right after **Xenium** everywhere technologies are
  listed.
- **Samples** — what specimen types were actually collected, per visit,
  broken down by disease team and cohort: FFPE Tissue Block, Frozen Tissue
  Block, Serum, PBMCs, Fresh Frozen Tissue - Cytodelics, PaxGene Tube, and
  Urine (from the `Tissue_Mold`, `Tissue_Cryo`, `Serum`, `PBMC`,
  `Cytodelics`, `Pax`, and `Urine` columns on the `visits` tab — each is
  just `Y`/blank per visit). A bar chart totals each sample type across the
  selected disease teams/cohorts, a mini-chart-per-disease-team view
  breaks that down further, and a matrix table shows raw counts by disease
  team and cohort — the most direct view of what's actually available for
  a given group. Disease team and cohort here come from a join back to the
  `subjects` tab (the same harmonized `Data_Scope`/`Dashboard_Label` used on
  the Recruitment tab), not from the `visits` tab's own, older cohort
  column, so cohort names stay consistent across all three tabs. A handful
  of visits that can't be matched to a subject are counted and called out
  in the tab's data-quality notes rather than silently dropped.

**About "disease team":** the source `Disease Team` column holds each row's
study/protocol name (STAMP, ELLIPSS, AIM for RA, LOCKIT, SSc Pilot), not a
diagnosis, so it isn't the grouping the dashboard needs. Both tabs instead
derive the disease team from the `Data_Scope` value (`SLE-KDY` → Lupus
Kidney, `RA-SYN` → Rheumatoid Arthritis (RA), etc.) — see `DISEASE_LABELS`,
`SPLIT_DISEASE_LABELS`, and `derive_diseases()` near the top of
`build_data.py`. If the true mapping ever needs to change, that's the place
to edit it.

## Quick start — get it live on GitHub Pages

This is already set up and done for the `mcmarlin/AMPAIM-Dashboard` repo,
live at https://mcmarlin.github.io/AMPAIM-Dashboard/ — you don't need to
repeat it. It's kept here only for reference, in case the site ever needs to
be rebuilt from scratch on a new repo. Your actual recurring workflow is
"Weekly update" below, which is drag-and-drop, not git.

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

This uses the layout from "Where your files live" above: a `Data` folder
with your two spreadsheets, sitting right next to your project folder.
Nothing here uses git or a git repo — you're not connected to GitHub through
git at all, so every step is either a local script or a drag-and-drop on
github.com.

Steps:

1. **Save this week's export.** Save/overwrite it as
   `AMP-AIM_Dataset_Weekly_Update.xlsx` in your `Data` folder (not inside the
   project folder). Nothing in this step touches GitHub; this file never
   leaves your computer.
2. **Open a terminal in your project folder and run the build script.** On
   Windows, open Command Prompt or PowerShell and navigate into your project
   folder (for example `cd Desktop\AMP-AIM_Progress_Dashboard\amp-dashboard_v4`,
   adjusted to your actual folder name), then run:
   ```
   python3 -m pip install -r requirements.txt
   python3 build_data.py "../Data/AMP-AIM_Dataset_Weekly_Update.xlsx" data/dashboard.json "../Data/Target_Recruitment_Numbers.xlsx"
   ```
   The `../Data/...` paths point one folder up and into `Data` — that's what
   makes this find your spreadsheets now that they live outside the project
   folder. (Quotes around each path aren't strictly required anymore since
   neither filename has spaces, but they don't hurt either.)
   (The `pip install` line only needs to be run once, ever — skip it on
   later weeks. Using `python3 -m pip install...` rather than just
   `pip install...` matters especially on Windows, where a plain `pip`
   command can sometimes point at a different Python install than `python3`
   does. If you see `ModuleNotFoundError: No module named 'openpyxl'` even
   after running it once, that mismatch is almost certainly why — rerun it
   as `python3 -m pip install -r requirements.txt`.)
   This reads both spreadsheets and overwrites `data/dashboard.json` (inside
   your project folder) with the new aggregate counts — check the terminal
   output for a `subjects=... visits=...` line confirming it worked, and for
   any `WARNING:` lines (they mean a file wasn't found at the path given).
3. **Upload just that one file to GitHub.** This is the step that's easy to
   get wrong, because GitHub uploads to *whatever folder you're currently
   browsing* — so you have to be inside the right folder before you drag
   anything in:
   1. Go to https://github.com/mcmarlin/AMPAIM-Dashboard in your browser.
   2. Click into the **`data`** folder (the address bar should end in
      `.../AMPAIM-Dashboard/tree/main/data`). Do this *before* uploading —
      uploading from the repo's front page puts the file at the repo root
      instead, where the live site will never see it.
   3. Click **Add file → Upload files**, then drag in `dashboard.json` from
      your project folder's `data` subfolder.
   4. GitHub will show it replacing the existing `data/dashboard.json` —
      that's your confirmation you're in the right place. Scroll down and
      click **Commit changes**.
   The live site updates within a minute or two. You never need to upload
   the `.xlsx` files, `build_data.py`, or anything else — `dashboard.json`
   is the only file that changes week to week.

## Updating recruitment targets

The "of N" progress numbers on the Recruitment tab (e.g. "192 of 488") come
from `Target_Recruitment_Numbers.xlsx` in your `Data` folder, not from the
weekly sample export — these targets change far less often, so they're kept
in their own small spreadsheet. To update one:

1. Open `Target_Recruitment_Numbers.xlsx` (in your `Data` folder, next to
   `AMP-AIM_Dataset_Weekly_Update.xlsx`) in Excel.
2. Each row is one cohort. Edit the number in the **Expected** column, or
   type `Undefined` if a target isn't set yet. (The **Notes** column is just
   for your own reference — it isn't read by the dashboard.) Don't rename the
   **Disease Team** or **Cohort** values, or a target won't be found for that
   cohort — the **Cohort** values need to match the new `Dashboard_Label`
   tags from the `subjects` tab (see "Generating a fresh Disease Team x
   Cohort reference" below if you need the current list).
3. Save the file, in place, still named `Target_Recruitment_Numbers.xlsx`,
   still in the `Data` folder.
4. Run steps 2 and 3 of the weekly update above (rebuild, then upload just
   `data/dashboard.json` into the repo's `data` folder on GitHub) — the new
   targets are picked up automatically, no separate step needed. If you're
   only updating targets and not this week's sample export, you can skip
   step 1 and go straight to rebuilding; the script reads both spreadsheets
   every time regardless of which one changed.

If this file is ever missing or moved, `build_data.py` won't fail — it just
prints a warning and every recruitment bar shows "not yet set" until the
file is back in place at `../Data/Target_Recruitment_Numbers.xlsx`.

### Generating a fresh Disease Team x Cohort reference

`automation/generate_cohort_reference.py` reads the current `subjects` tab
and writes out every Disease Team x Cohort combination that actually appears
in the data, in the new `Dashboard_Label` format, alongside each combo's
current subject count and its existing target (carried over automatically
when the name matches `Target_Recruitment_Numbers.xlsx` exactly). Useful
whenever cohort labels change and `Target_Recruitment_Numbers.xlsx` needs to
be brought up to date:
```
python3 automation/generate_cohort_reference.py "../Data/AMP-AIM_Dataset_Weekly_Update.xlsx" "../Data/Target_Recruitment_Numbers.xlsx" "../Data/Cohort_Reference_New_Format.xlsx"
```
This is a reference file only — it's never read by the dashboard itself.
Copy the Expected values you want to keep over into
`Target_Recruitment_Numbers.xlsx` by hand.

## Removing a file you accidentally uploaded

If you drag a file into the wrong folder on GitHub (like the stray root-level
`dashboard.json` some updates can create — see step 3 above), removing it is
another drag-and-drop-style flow, no command line needed:

1. Go to https://github.com/mcmarlin/AMPAIM-Dashboard and click through the
   folders to the file (for a file sitting at the repo root, it's right
   there on the front page).
2. Click the file's name to open it.
3. Click the trash-can icon in the top-right of the file view (or the **...**
   menu → **Delete file**).
4. Scroll down and click **Commit changes** (committing directly to `main`
   is fine for this).

That's the right move for an ordinary mistaken upload — an extra copy of
`dashboard.json`, an unwanted image, and so on. **It is not enough for
anything sensitive.** GitHub keeps every previous version of every file
forever, so a deleted file is still sitting in the repo's history and
recoverable by anyone — this matters most for `AMP-AIM_Dataset_Weekly_Update.xlsx`,
which must never be uploaded here at all (see "Why this is aggregate-only" above).
If that ever happens by mistake, deleting it isn't sufficient — stop and
either make the repo private immediately or contact GitHub support about
fully purging it from history, and check with your PI/data manager.

## Fully automated weekly refresh (optional)

This is a different, more advanced path than the drag-and-drop workflow
above — it needs git and a command line, so it's worth doing only if that's
comfortable for you. If you'd rather not run a command by hand each week,
`automation/private-repo-update-and-deploy.yml` sets up a GitHub Action that
does it for you — but it needs to live in a **separate private repo** that
holds your raw spreadsheet (never the public one). See the comments at the
top of that file for the one-time setup. Once configured, updating the
spreadsheet in the private repo is the entire weekly workflow; the public
dashboard refreshes itself.

## Local preview / troubleshooting

Opening `index.html` directly by double-clicking it will show a "could not
load data" message — browsers block `fetch()` on `file://` URLs. Preview it
with a local server instead:
```
python3 -m http.server 8000
```
then open `http://localhost:8000/`.
