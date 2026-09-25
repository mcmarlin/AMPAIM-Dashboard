#!/usr/bin/env python3
"""
build_data.py

Reads the weekly AMP AIM sample-tracking export (headers on row 3) and produces
a fully AGGREGATE, de-identified dashboard.json for the public progress dashboard.

Privacy design (do not change without re-checking with the data owner):
  - No Subject_ID, Subject-Visit_ID, or any other row-level identifier is ever
    written to the output.
  - No diagnosis (Enroll_Dx), demographic field (Sex/Race/Ethnicity/Age), or
    free-text Notes value is ever written to the output.
  - Only COUNTS are emitted, grouped by technology / disease-tissue scope /
    pipeline / cohort. Cells with fewer than MIN_CELL_COUNT contributing rows
    are still reported (counts, not identities) but the script makes it easy
    to raise that floor if small-cell suppression is ever required.

Usage:
    python3 build_data.py [input.xlsx] [output.json] [expected_recruitment.xlsx]

Defaults: data/AMP-AIM_Dataset_Weekly_Update.xlsx -> data/dashboard.json, reading
recruitment targets from data/Target_Recruitment_Numbers.xlsx.

The input workbook is a multi-tab weekly export. Two sheets matter here:
  - "subjects" - one row per subject. This is now the source of truth for the
    Recruitment tab: Data_Scope gives the disease team, Dashboard_Label gives
    the cohort(s, still shown as "Cohorts" in the UI), Subject Type gives
    enrolled/archival status, and Visits gives that subject's visit count.
  - "visits"   - one row per subject-visit, still used for the Technologies
    tab's per-assay/sample tracking.
Both sheets share the same 3-row header layout (see HEADER_ROW/DATA_START_ROW
below).
"""
import sys
import json
import datetime
from collections import defaultdict

import openpyxl

HEADER_ROW = 3
DATA_START_ROW = 4

# Technology (sample/assay) columns -> (key, display label)
# Trimmed to the 12 technologies actually tracked on the dashboard. (Plain
# Xenium, both "-NL" / Non-lesional columns, and the old OMRF Lab / Xenium
# (Slide) labels were dropped or renamed - see README history / requests.)
TECH_COLUMNS = [
    ("Xenium-Slide_Sample_ID", "xenium_slide", "Xenium"),
    ("Post-Xenium_HE", "post_xenium_he", "Post-Xenium H&E"),
    ("Post-Xenium_IMC", "post_xenium_imc", "Post Xenium IMC"),
    ("scFFPE_Sample_ID", "scffpe", "scFFPE"),
    ("scRNAseq_Sample_ID", "scrnaseq", "scRNA-seq"),
    ("Olink-Serum_Sample_ID", "olink_serum", "Olink (Serum)"),
    ("CyTOF-Blood_Sample_ID", "cytof_blood", "CyTOF (Blood)"),
    ("Genotyping_Sample_ID", "genotyping", "Genotyping"),
    ("OMRF-Lab_Sample_ID", "omrf_lab", "Autoantibody Testing"),
    ("bRNAseq_Sample_ID", "brnaseq", "Bulk RNA-Seq"),
    ("Olink-Urine_Sample_ID", "olink_urine", "Olink (Urine)"),
    ("mBioSeq-stool", "mbioseq_stool", "Microbiome (Stool)"),
    ("mBioSeq-skin", "mbioseq_skin", "Microbiome (Skin)"),
    ("anti-C1Q_Sample_ID", "anti_c1q", "Anti-C1Q"),
]

# Sample-type columns for the Samples tab -> (key, display label). These are
# simple Y/blank flags on the "visits" sheet (was a specimen of this type
# collected at this visit?), unlike TECH_COLUMNS above (which track a
# multi-status assay pipeline per sample). Order here is the order they're
# shown in on the dashboard (pills, tables, etc).
SAMPLE_TYPE_COLUMNS = [
    ("Tissue_Mold", "tissue_mold", "FFPE Tissue Block"),
    ("Tissue_Cryo", "tissue_cryo", "Fresh Frozen Tissue"),
    ("Serum", "serum", "Serum"),
    ("PBMC", "pbmc", "PBMCs"),
    ("Cytodelics", "cytodelics", "Frozen WB in Cytodelics (MDIPA)"),
    ("Pax", "pax", "PaxGene Tube"),
    ("Urine", "urine", "Urine"),
]


def is_collected(value):
    """Sample-type columns are just 'Y' (collected) or blank (not collected)."""
    if value is None:
        return False
    return str(value).strip().upper() == "Y"

STATUS_ORDER = ["completed", "pending", "not_applicable", "no_specimen", "qc_fail", "unknown"]
STATUS_LABELS = {
    "completed": "Completed",
    "pending": "Pending / Specimen Available",
    "not_applicable": "Not Applicable",
    "no_specimen": "No Specimen",
    "qc_fail": "QC Fail",
    "unknown": "Unknown / Not Yet Retrieved",
}

# "Disease Team" in the source holds each row's study/protocol name (e.g.
# STAMP, ELLIPSS, AIM for RA, LOCKIT, SSc Pilot) rather than a diagnosis, so
# it isn't the grouping the dashboard needs. Data_Scope ("<disease>-<tissue>",
# e.g. "SLE-KDY") is populated and reliable, so the disease/team grouping is
# derived from its prefix instead. EDIT THIS if your actual disease-team
# names differ, or add an explicit Subject_ID -> team override if Data_Scope
# isn't authoritative for some subjects.
DISEASE_LABELS = {
    "SLE": "Lupus (SLE)",
    "RA": "Rheumatoid Arthritis (RA)",
    "PsD": "Psoriatic Disease (PsD)",
    "SjD": "Sjögren's Disease (SjD)",
    "SSc": "Systemic Sclerosis (SSc)",
    "SLE/PsD": "SLE / PsD (combined)",
}

# SLE and PsD are further split by tissue (the part of Data_Scope after the
# hyphen) into the finer subgroups the dashboard now reports as their own
# "disease teams". (prefix, tissue) -> (key, display label). Any prefix/tissue
# combo NOT listed here just falls back to the whole-prefix grouping above
# (e.g. RA-SYN, SjD-SGL, SSc-SKN all stay as one bucket per prefix). PsD no
# longer has a separate Eye/Uveitis tissue on the "subjects" tab (Data_Scope
# there is just PsD-SYN / PsD-SKN); a handful of older PsD-EYE rows can still
# show up on the visit-level "visits" tab, and those simply roll up into the
# combined Psoriatic Disease (PsD) bucket via the fallback above.
SPLIT_DISEASE_LABELS = {
    ("SLE", "KDY"): ("sle_kdy", "Lupus Kidney"),
    ("SLE", "SKN"): ("sle_skn", "Lupus Skin"),
    ("PsD", "SYN"): ("psd_syn", "Psoriatic Arthritis"),
    ("PsD", "SKN"): ("psd_skn", "Psoriasis"),
}

# Label used for a blank/None Pipeline cell.
UNDEFINED_PIPELINE = "Undefined"

# A subject's "Subject Type" (on the "subjects" tab) tells us their
# recruitment status: actively-enrolled subjects are "enrolled", archival
# specimens are "archival", and Enabling subjects (a meaningfully different
# population from the other two) get their own "enabling" bucket.
# "Pre-Participation" (and anything else unrecognized) isn't mapped here at
# all - those subjects aren't placed in a cohort, and are counted in
# `unassigned_subjects` instead (see main()).
#
# The source tracker has used two different vocabularies for this column
# over time - both are mapped here so either a current or older export
# works:
#   - current (as of Sep 2026): Standard, Longitudinal -> enrolled;
#     Enabling -> enabling; Archival -> archival
#   - older: Enrolled -> enrolled; Enabling only -> enabling;
#     Archival only -> archival
# If the source tracker changes this vocabulary again, `unassigned_subjects`
# in the build_data.py console output will jump toward the total subject
# count (as it did when this comment was added) - that's the signal to
# check here first, and to confirm the new mapping with whoever maintains
# the tracker before guessing at it.
SUBJECT_STATUS_BY_TYPE = {
    # current vocabulary
    "Standard": "enrolled",
    "Longitudinal": "enrolled",
    "Enabling": "enabling",
    "Archival": "archival",
    # older vocabulary (kept for compatibility with older exports)
    "Enrolled": "enrolled",
    "Enabling only": "enabling",
    "Archival only": "archival",
}
RECRUIT_STATUS_ORDER = ["enrolled", "enabling", "archival"]
RECRUIT_STATUS_LABELS = {"enrolled": "Enrolled", "enabling": "Enabling", "archival": "Archival"}

# A subject's "Schedule_Status" (on the "subjects" tab, added ~Sep 2026) is a
# SEPARATE axis from Subject Type above - whether they've completed all
# their scheduled visits yet, not their enrollment/archival status. Only
# "Complete" is a real value seen so far; blank/anything else means their
# visit schedule isn't finished. This column may not exist on older exports
# (see SCHEDULE_STATUS_COL below) - when it's missing, every subject falls
# back to "not_complete", which just means this alternate view isn't useful
# until the export includes the column (a data-quality note on the
# dashboard calls this out).
SCHEDULE_STATUS_COL = "Schedule_Status"
SCHEDULE_STATUS_ORDER = ["complete", "not_complete"]
SCHEDULE_STATUS_LABELS = {"complete": "Complete", "not_complete": "Not Complete"}


def schedule_status_for(raw_value):
    return "complete" if clean_label(raw_value, default="") == "Complete" else "not_complete"


def empty_schedule_status():
    return {k: 0 for k in SCHEDULE_STATUS_ORDER}

# How many subjects are ultimately expected in each cohort, per the network's
# recruitment targets. Loaded at runtime (see load_expected_recruitment(),
# called from main()) from a spreadsheet - data/Target_Recruitment_Numbers.xlsx
# by default - with columns "Disease Team", "Cohort", "Expected" (an optional
# "Notes" column is ignored). Edit that spreadsheet directly to update targets;
# no code change needed. A blank or "Undefined" Expected cell means "not yet
# defined". Keyed by the ORIGINAL (pre-split) disease-team label, since that's
# the level the targets were given at; a cohort's new, finer disease-subgroup
# membership (see SPLIT_DISEASE_LABELS) is looked up from the real data, and
# this table is then consulted by that cohort's old/whole-disease label.
EXPECTED_RECRUITMENT_PATH_DEFAULT = "data/Target_Recruitment_Numbers.xlsx"
EXPECTED_RECRUITMENT = {}  # populated by load_expected_recruitment() in main()


def load_expected_recruitment(path):
    """
    Read the recruitment-target spreadsheet and return
    {disease_team_label: {cohort_name_lowercased: int_or_None}}.
    Returns {} (every target "not yet set") if the file is missing or doesn't
    have the expected headers - the dashboard still builds fine either way,
    it just won't show "of N" progress numbers until the file is in place.
    """
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
    except FileNotFoundError:
        print(f"WARNING: recruitment-target file not found at {path!r} - "
              f"recruitment bars will show as \"not yet set\" until it's added.",
              file=sys.stderr)
        return {}
    ws = wb[wb.sheetnames[0]]
    headers = [c.value.strip() if isinstance(c.value, str) else c.value for c in ws[1]]
    col_idx = {h: i for i, h in enumerate(headers) if h}
    required = ["Disease Team", "Cohort", "Expected"]
    missing = [c for c in required if c not in col_idx]
    if missing:
        print(f"WARNING: {path!r} is missing column(s) {missing} (expected "
              f"'Disease Team', 'Cohort', 'Expected' headers on row 1) - "
              f"recruitment bars will show as \"not yet set\".", file=sys.stderr)
        return {}

    out = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        team_cell = row[col_idx["Disease Team"]]
        cohort_cell = row[col_idx["Cohort"]]
        if team_cell is None or cohort_cell is None:
            continue
        team = str(team_cell).strip()
        cohort = str(cohort_cell).strip().lower()
        raw = row[col_idx["Expected"]]
        value = int(raw) if isinstance(raw, (int, float)) else None
        out.setdefault(team, {})[cohort] = value
    return out


def expected_for(old_label, cohort_name):
    """Look up a cohort's recruitment target. None if not defined/found."""
    return EXPECTED_RECRUITMENT.get(old_label, {}).get(cohort_name.strip().lower())


def empty_recruit_status():
    return {k: 0 for k in RECRUIT_STATUS_ORDER}


def derive_diseases(scope_label):
    """
    Return the list of (key, label, old_label, old_group_key) disease-team
    entries a given Data_Scope value belongs to. Usually a single entry (e.g.
    'RA-SYN' -> [('ra', 'Rheumatoid Arthritis (RA)', 'Rheumatoid Arthritis
    (RA)', 'ra')]), but Data_Scope can encode more than one disease/tissue in
    a single cell with a '/' - either in the prefix ('SLE/PsD-SKN', an
    overlap-syndrome subject) or in the tissue ('PsD-SKN/SYN', a subject with
    data under both tissues). In either case the subject is counted in EACH
    matching disease team, but only once in the site-wide subject total (that
    total is a distinct count of Subject_ID and never goes through this
    function).
    `old_label`/`old_group_key` are the pre-split whole-disease label/key
    (e.g. still "Lupus (SLE)" / "sle" for both Lupus subgroups) - used to
    roll the new finer subgroups back up into one combined bar with one
    combined Expected Recruitment target, since that's the level those
    targets were given at.
    """
    if not scope_label or scope_label == "Unknown":
        return [("unknown", "Unknown", "Unknown", "unknown")]
    if "-" in scope_label:
        prefix_part, tissue_part = scope_label.split("-", 1)
    else:
        prefix_part, tissue_part = scope_label, ""
    prefixes = [p.strip() for p in prefix_part.split("/") if p.strip()] or [prefix_part.strip()]
    tissues = [t.strip() for t in tissue_part.split("/") if t.strip()] or [""]

    seen = {}
    for prefix in prefixes:
        old_label = DISEASE_LABELS.get(prefix, prefix)
        old_group_key = prefix.lower().replace(" ", "_")
        for tissue in tissues:
            split = SPLIT_DISEASE_LABELS.get((prefix, tissue))
            if split:
                key, label = split
            else:
                key = old_group_key
                label = old_label
            if key not in seen:
                seen[key] = (key, label, old_label, old_group_key)
    return list(seen.values())


# For "Cohorts within each disease team" specifically, Psoriasis (PsD-SKN)
# and Psoriatic Arthritis (PsD-SYN) are combined into one "Psoriatic Disease
# (PsD)" card/table rather than shown separately - unlike Lupus Kidney/Skin,
# which stay split there. Everything else (including Lupus) keeps its normal
# fine-grained key/label in this view. This affects ONLY the merged
# recruitment.cohort_view output built in main() - the fine-grained
# by_disease/by_cohort output used elsewhere (disease filter pills, KPI
# totals, the top "Subjects enrolled by disease team" bar, Samples tab,
# etc.) is untouched.
COHORT_VIEW_MERGE = {
    "psd_syn": "psd",
    "psd_skn": "psd",
}
COHORT_VIEW_MERGE_LABELS = {
    "psd": "Psoriatic Disease (PsD)",
}

# Individual cohort-level recruitment targets are normally looked up by a
# disease team's pre-split "old_label" (e.g. both Lupus subgroups share
# "Lupus (SLE)"). Lupus Kidney and Lupus Skin need their OWN separate
# targets even for a same-named cohort (e.g. both have an "Enabling Case"
# cohort), so their cohort-level targets are looked up by their own fine
# label instead - meaning Target_Recruitment_Numbers.xlsx needs separate
# "Lupus Kidney" / "Lupus Skin" rows now rather than shared "Lupus (SLE)"
# rows for cohort-level targets. Psoriatic Disease doesn't need this
# override: its subgroups are combined into one row in the cohort view
# above, so there's only ever one row per cohort name there again.
TARGET_LOOKUP_OVERRIDE = {
    "sle_kdy": "Lupus Kidney",
    "sle_skn": "Lupus Skin",
}


def target_label_for(disease_key, old_label):
    """Which label to use as the "Disease Team" key when looking up a
    cohort's recruitment target - see TARGET_LOOKUP_OVERRIDE above."""
    return TARGET_LOOKUP_OVERRIDE.get(disease_key, old_label)


def cohort_view_group(disease_key, disease_label, old_label):
    """
    (key, label, target_label) for the merged "Cohorts within each disease
    team" view - see COHORT_VIEW_MERGE above. Defaults to this disease
    team's own fine-grained key/label (with Lupus's target-label override
    applied) when it isn't one of the merged groups.
    """
    merged_key = COHORT_VIEW_MERGE.get(disease_key)
    if merged_key:
        merged_label = COHORT_VIEW_MERGE_LABELS[merged_key]
        return merged_key, merged_label, target_label_for(merged_key, merged_label)
    return disease_key, disease_label, target_label_for(disease_key, old_label)


def classify(value):
    """Collapse the many raw status strings into 6 reportable buckets."""
    if value is None:
        return "unknown"
    s = str(value).strip()
    if s == "" or s in ("None", "#REF!", "[not yet retrieved]"):
        return "unknown"
    low = s.lower()
    if "qc fail" in low or "qc-fail" in low or "failed qc" in low:
        return "qc_fail"
    # "No Specimen" (and any bracketed reason alongside it, e.g.
    # "[No Specimen][Block Exhausted]") means there was never a specimen to
    # run this technology on - kept as its own bucket, distinct from the
    # rarer "Not Applicable"/"Not Available" cases below, because the
    # Technologies tab's "Status breakdown by technology" chart excludes it
    # entirely (not shown as a segment, not counted in that chart's total).
    if "no specimen" in low:
        return "no_specimen"
    if "not applicable" in low or "not available" in low:
        return "not_applicable"
    if "not ordered" in low:
        return "not_applicable"
    if "specimen available" in low or "pending" in low:
        return "pending"
    if s.startswith("[") and s.endswith("]"):
        # any other bracketed status we haven't explicitly mapped
        return "unknown"
    # anything else is an actual sample/barcode ID (or "Yes") -> completed
    return "completed"


def clean_label(value, default="Unknown"):
    """Clean a single-value field (Data_Scope, Pipeline, ...) into one label."""
    if value is None:
        return default
    s = str(value).strip()
    if s == "" or s in ("None", "#REF!") or s.lower() == "n/a":
        return default
    if s.startswith("[") and s.endswith("]"):
        inner = s[1:-1]
        parts = [p.strip() for p in inner.split("][") if p.strip()]
        s = ", ".join(parts) if parts else default
    if s.lower() == "not yet retrieved":
        s = "Not yet retrieved"
    return s


def split_cohort_tags(value):
    """
    Dashboard_Label can hold MULTIPLE tags back-to-back in one cell, e.g.
    "[PsD Eye][PsD axSpA]" for a subject who belongs to both. Returns a LIST
    of individual cohort tags rather than one combined string, so a subject
    like that is counted once in EACH cohort - never as a separate
    "PsD Eye, PsD axSpA" bucket.
    """
    if value is None:
        return ["Unknown"]
    s = str(value).strip()
    if s == "" or s in ("None", "#REF!"):
        return ["Unknown"]
    if s.startswith("[") and s.endswith("]"):
        inner = s[1:-1]
        parts = [p.strip() for p in inner.split("][") if p.strip()]
        if not parts:
            return ["Unknown"]
        return ["Not yet retrieved" if p.lower() == "not yet retrieved" else p for p in parts]
    if s.lower() == "not yet retrieved":
        return ["Not yet retrieved"]
    return [s]


def empty_status_counts():
    return {k: 0 for k in STATUS_ORDER}


def add_counts(dst, src):
    for k in STATUS_ORDER:
        dst[k] = dst.get(k, 0) + (src.get(k, 0) if src else 0)
    return dst


def main():
    in_path = sys.argv[1] if len(sys.argv) > 1 else "data/AMP-AIM_Dataset_Weekly_Update.xlsx"
    out_path = sys.argv[2] if len(sys.argv) > 2 else "data/dashboard.json"
    expected_path = sys.argv[3] if len(sys.argv) > 3 else EXPECTED_RECRUITMENT_PATH_DEFAULT

    global EXPECTED_RECRUITMENT
    EXPECTED_RECRUITMENT = load_expected_recruitment(expected_path)

    wb = openpyxl.load_workbook(in_path, data_only=True)

    # The workbook is a multi-tab weekly export now. Visit-level sample/assay
    # tracking (Technologies tab) lives on the "visits" sheet. Fall back to
    # the first sheet (the old single-sheet layout) with a warning if
    # "visits" isn't found, rather than silently mis-parsing whatever sheet
    # happens to be first.
    if "visits" in wb.sheetnames:
        ws = wb["visits"]
    else:
        print(f"WARNING: no 'visits' sheet found in {in_path!r} - falling back to "
              f"the first sheet ({wb.sheetnames[0]!r}) for Technologies-tab data. "
              f"Counts may be wrong if that isn't actually the visit-level data.",
              file=sys.stderr)
        ws = wb[wb.sheetnames[0]]

    headers = [ws.cell(row=HEADER_ROW, column=c).value for c in range(1, ws.max_column + 1)]
    col_idx = {h: i + 1 for i, h in enumerate(headers) if h}

    missing = [col for col, _, _ in TECH_COLUMNS if col not in col_idx]
    if missing:
        print(f"WARNING: expected columns not found and will be skipped: {missing}", file=sys.stderr)

    subject_idx = col_idx.get("Subject_ID")
    data_scope_idx = col_idx.get("Data_Scope")
    pipeline_idx = col_idx.get("Pipeline")

    subjects_seen = set()
    n_visits = 0

    # tech key -> overall status counts
    tech_totals = {key: empty_status_counts() for _, key, _ in TECH_COLUMNS}
    # tech key -> {scope_label: status_counts}  (all pipelines combined)
    tech_by_scope = {key: defaultdict(empty_status_counts) for _, key, _ in TECH_COLUMNS}
    # tech key -> {pipeline_label: status_counts}  (all scopes combined)
    tech_by_pipeline = {key: defaultdict(empty_status_counts) for _, key, _ in TECH_COLUMNS}
    # tech key -> {scope_label: {pipeline_label: status_counts}}  (for disease-filtered pipeline views)
    tech_by_scope_pipeline = {key: defaultdict(lambda: defaultdict(empty_status_counts)) for _, key, _ in TECH_COLUMNS}

    scopes_seen = set()
    pipelines_seen = set()

    for r in range(DATA_START_ROW, ws.max_row + 1):
        subj = ws.cell(row=r, column=subject_idx).value if subject_idx else None
        subj_clean = clean_label(subj, default=None)
        # skip fully blank rows
        row_has_any = any(
            ws.cell(row=r, column=col_idx[col]).value is not None
            for col, _, _ in TECH_COLUMNS
            if col in col_idx
        )
        if subj_clean is None and not row_has_any:
            continue

        if subj_clean:
            subjects_seen.add(subj_clean)
        n_visits += 1

        scope_label = clean_label(ws.cell(row=r, column=data_scope_idx).value if data_scope_idx else None)
        pipeline_label = clean_label(
            ws.cell(row=r, column=pipeline_idx).value if pipeline_idx else None,
            default=UNDEFINED_PIPELINE,
        )
        scopes_seen.add(scope_label)
        pipelines_seen.add(pipeline_label)

        for col, key, _ in TECH_COLUMNS:
            if col not in col_idx:
                continue
            val = ws.cell(row=r, column=col_idx[col]).value
            status = classify(val)
            tech_totals[key][status] += 1
            tech_by_scope[key][scope_label][status] += 1
            tech_by_pipeline[key][pipeline_label][status] += 1
            tech_by_scope_pipeline[key][scope_label][pipeline_label][status] += 1

    def totals_block(counts):
        total = sum(counts.values())
        completed = counts["completed"]
        pct = round(100 * completed / total, 1) if total else 0.0
        eligible = counts["completed"] + counts["pending"] + counts["qc_fail"]
        pct_elig = round(100 * completed / eligible, 1) if eligible else None
        return {
            "counts": counts,
            "total": total,
            "pct_complete": pct,
            "eligible": eligible,
            "pct_eligible": pct_elig,
        }

    # Kept in TECH_COLUMNS order (not re-sorted by count) so the "Datasets
    # shown" pills - and anything else that lists technologies without doing
    # its own sort - show a stable order (e.g. the two Post-Xenium steps
    # right after Xenium) rather than one that reshuffles week to week as
    # completed counts change. The main "completed samples" bar chart and
    # the per-disease-team panels both sort by count themselves in the UI,
    # so they're unaffected by this order.
    technologies = []
    for _, key, label in TECH_COLUMNS:
        block = totals_block(tech_totals[key])
        block["key"] = key
        block["label"] = label
        technologies.append(block)

    scopes_sorted = sorted(s for s in scopes_seen if s != "Unknown") + (
        ["Unknown"] if "Unknown" in scopes_seen else []
    )
    pipelines_sorted = sorted(pipelines_seen)  # "EDP1","EDP2","NRP","Undefined" sorts naturally

    scope_disease = {}
    for scope in scopes_sorted:
        scope_disease[scope] = [{"key": k, "label": l} for k, l, _old, _grp in derive_diseases(scope)]

    # ---------- Recruitment: subjects & visits per disease team / cohort ----------
    # One row per subject on the "subjects" tab - Data_Scope, Dashboard_Label,
    # Subject Type and Visits are all read straight off that subject's own
    # row, no more scanning multiple visit rows to find an "enrollment" one.
    if "subjects" in wb.sheetnames:
        subj_ws = wb["subjects"]
    else:
        subj_ws = None
        print(f"WARNING: no 'subjects' sheet found in {in_path!r} - the "
              f"Recruitment tab will show zero subjects.", file=sys.stderr)

    disease_totals = defaultdict(lambda: {
        "subjects": 0, "visits": 0, "label": "", "old_label": "", "old_group_key": "",
        "status": empty_recruit_status(),
    })
    cohort_totals = defaultdict(lambda: {
        "subjects": 0, "visits": 0, "disease_key": "", "disease_label": "", "old_label": "",
        "status": empty_recruit_status(),
    })
    # Rolled back up to the ORIGINAL, pre-split disease team (e.g. "sle"
    # combines sle_kdy + sle_skn) - this is where a single, shared Expected
    # Recruitment target still makes sense (see the "combined bar" note
    # above derive_diseases). Subject membership here is a real set, so a
    # subject touching both Lupus subgroups still counts once.
    combined_totals = defaultdict(lambda: {"subject_set": set(), "visits": 0, "label": "", "status": empty_recruit_status()})
    # The merged view used ONLY by "Cohorts within each disease team" - see
    # cohort_view_group()/COHORT_VIEW_MERGE above. Same shape as
    # disease_totals/cohort_totals, plus a "target_label" (which label to
    # use when looking up that row's recruitment target - see
    # TARGET_LOOKUP_OVERRIDE) and a "schedule_status" breakdown alongside
    # the usual enrolled/enabling/archival "status".
    cohort_view_disease_totals = defaultdict(lambda: {
        "subjects": 0, "visits": 0, "label": "", "target_label": "",
        "status": empty_recruit_status(), "schedule_status": empty_schedule_status(),
    })
    cohort_view_cohort_totals = defaultdict(lambda: {
        "subjects": 0, "visits": 0, "disease_key": "", "disease_label": "", "target_label": "",
        "status": empty_recruit_status(), "schedule_status": empty_schedule_status(),
    })
    unassigned_subjects = 0
    # Subject_ID -> {"scope":, "tags":} for EVERY subject on the "subjects"
    # sheet (including Pre-Participation ones, unlike the recruitment totals
    # below) - used by the Samples tab further down to attribute each visit
    # row on the "visits" sheet to a disease team/cohort via a join, instead
    # of trusting that sheet's own (older, less consistent) per-row values.
    subject_join = {}

    if subj_ws is not None:
        subj_headers = [subj_ws.cell(row=HEADER_ROW, column=c).value for c in range(1, subj_ws.max_column + 1)]
        subj_col_idx = {h: i + 1 for i, h in enumerate(subj_headers) if h}
        required_subj_cols = ["Subject_ID", "Data_Scope", "Dashboard_Label", "Subject Type", "Visits"]
        missing_subj_cols = [c for c in required_subj_cols if c not in subj_col_idx]
        if missing_subj_cols:
            print(f"WARNING: 'subjects' sheet is missing column(s) {missing_subj_cols} - "
                  f"Recruitment-tab counts will be incomplete.", file=sys.stderr)

        sid_idx = subj_col_idx.get("Subject_ID")
        sscope_idx = subj_col_idx.get("Data_Scope")
        slabel_idx = subj_col_idx.get("Dashboard_Label")
        stype_idx = subj_col_idx.get("Subject Type")
        svisits_idx = subj_col_idx.get("Visits")
        sched_idx = subj_col_idx.get(SCHEDULE_STATUS_COL)
        if sched_idx is None:
            print(f"WARNING: 'subjects' sheet has no {SCHEDULE_STATUS_COL!r} column - the "
                  f"Recruitment tab's Complete/Not Complete view will show every subject as "
                  f"\"Not Complete\" until it's present.", file=sys.stderr)

        for r in range(DATA_START_ROW, subj_ws.max_row + 1):
            sid_raw = subj_ws.cell(row=r, column=sid_idx).value if sid_idx else None
            sid = clean_label(sid_raw, default=None)
            if sid is None:
                continue  # blank row

            scope = clean_label(subj_ws.cell(row=r, column=sscope_idx).value if sscope_idx else None)
            tags = split_cohort_tags(subj_ws.cell(row=r, column=slabel_idx).value if slabel_idx else None)
            subject_join[sid] = {"scope": scope, "tags": tags}

            type_raw = subj_ws.cell(row=r, column=stype_idx).value if stype_idx else None
            recruit_status = SUBJECT_STATUS_BY_TYPE.get(clean_label(type_raw, default=""))
            if recruit_status is None:
                # "Pre-Participation" (or any other/unrecognized Subject
                # Type) isn't placed in a cohort here, but is still counted
                # in overall totals elsewhere on the site.
                unassigned_subjects += 1
                continue

            visits_raw = subj_ws.cell(row=r, column=svisits_idx).value if svisits_idx else None
            n_subject_visits = int(visits_raw) if isinstance(visits_raw, (int, float)) else 0
            sched_status = schedule_status_for(subj_ws.cell(row=r, column=sched_idx).value if sched_idx else None)

            touched_groups = {}
            touched_cv = {}  # cv_key -> (cv_label, cv_target_label) - see cohort_view_group()
            # A subject can land in more than one disease-team bucket - see
            # derive_diseases() for the combo Data_Scope cases ("PsD-SKN/SYN",
            # "SLE/PsD-SKN", etc). They're counted once in EACH matching team,
            # but the site-wide subject total is a separate, plain distinct-
            # Subject_ID count (below) so it's never inflated by this.
            for disease_key, disease_label, old_label, old_group_key in derive_diseases(scope):
                disease_totals[disease_key]["subjects"] += 1
                disease_totals[disease_key]["visits"] += n_subject_visits
                disease_totals[disease_key]["label"] = disease_label
                disease_totals[disease_key]["old_label"] = old_label
                disease_totals[disease_key]["old_group_key"] = old_group_key
                disease_totals[disease_key]["status"][recruit_status] += 1
                touched_groups[old_group_key] = old_label
                cv_key, cv_label, cv_target_label = cohort_view_group(disease_key, disease_label, old_label)
                touched_cv[cv_key] = (cv_label, cv_target_label)
                # ...and within that team, counts in EVERY cohort tag they
                # carry (a subject in both "[PsD Eye]" and "[PsD axSpA]" adds
                # 1 to each).
                for tag in sorted(set(tags)):
                    ck = (disease_key, tag)
                    cohort_totals[ck]["subjects"] += 1
                    cohort_totals[ck]["visits"] += n_subject_visits
                    cohort_totals[ck]["disease_key"] = disease_key
                    cohort_totals[ck]["disease_label"] = disease_label
                    cohort_totals[ck]["old_label"] = old_label
                    cohort_totals[ck]["status"][recruit_status] += 1
            # Roll this subject up into each ORIGINAL (pre-split) team they
            # touched - once each, even if they hit more than one new
            # subgroup within that same original team.
            for group_key, old_label in touched_groups.items():
                g = combined_totals[group_key]
                g["subject_set"].add(sid)
                g["visits"] += n_subject_visits
                g["label"] = old_label
                g["status"][recruit_status] += 1
            # Same dedup idea, for the merged "Cohorts within each disease
            # team" view: touched_cv only has ONE entry per merged group
            # even if this subject hit it via more than one underlying
            # subgroup (e.g. both PsD tissues), so this subject - and each
            # of its cohort tags - is only counted once per merged group.
            for cv_key, (cv_label, cv_target_label) in touched_cv.items():
                cvd = cohort_view_disease_totals[cv_key]
                cvd["subjects"] += 1
                cvd["visits"] += n_subject_visits
                cvd["label"] = cv_label
                cvd["target_label"] = cv_target_label
                cvd["status"][recruit_status] += 1
                cvd["schedule_status"][sched_status] += 1
                for tag in sorted(set(tags)):
                    cvck = (cv_key, tag)
                    c = cohort_view_cohort_totals[cvck]
                    c["subjects"] += 1
                    c["visits"] += n_subject_visits
                    c["disease_key"] = cv_key
                    c["disease_label"] = cv_label
                    c["target_label"] = cv_target_label
                    c["status"][recruit_status] += 1
                    c["schedule_status"][sched_status] += 1

    def disease_expected(member_keys, cohort_totals_by_key):
        """
        Sum this disease team's cohort-level recruitment targets - one
        lookup per distinct (TARGET LABEL, cohort) pair, not per (disease
        subgroup, cohort) pair. Those aren't the same thing: Lupus Kidney and
        Lupus Skin have DIFFERENT target labels (via TARGET_LOOKUP_OVERRIDE),
        so their same-named "Enabling Case" cohort correctly gets summed
        once per subgroup (two different targets). PsD's subgroups share the
        SAME target label (no override) - and a cohort tag like "PsD Skin"
        or "PsD Synovium" isn't tissue-locked, so the same cohort can show
        up under BOTH the psd_syn and psd_skn subgroups (whichever tissue
        each contributing subject's Data_Scope happens to touch). Without
        deduping by the actual lookup key, that cohort's single target would
        get added in twice - once per subgroup - inflating the combined PsD
        target well past what Target_Recruitment_Numbers.xlsx actually says.
        Returns (total_or_None, partial): partial=True means at least one
        contributing cohort has no target defined yet, so the sum is a floor
        (actual target is >= this), not the full picture.
        """
        total, any_known, any_unknown = 0, False, False
        seen = set()
        for ck, cv in cohort_totals_by_key.items():
            if cv["disease_key"] not in member_keys:
                continue
            label = target_label_for(cv["disease_key"], cv["old_label"])
            target_key = (label, ck[1].strip().lower())
            if target_key in seen:
                continue  # same (target label, cohort) already counted via another subgroup
            seen.add(target_key)
            v = expected_for(label, ck[1])
            if v is None:
                any_unknown = True
            else:
                total += v
                any_known = True
        return (total if any_known else None), any_unknown

    # Fine-grained (Lupus Kidney / Lupus Skin / ... ) - used for the disease
    # filter pills, KPI totals, and the per-subgroup cohort cards. No target
    # numbers at this granularity (see the module docstring / derive_diseases
    # note) - those live one level up, on by_disease_group below.
    by_disease = sorted(
        [{"key": k, "label": v["label"], "subjects": v["subjects"], "visits": v["visits"], "status": v["status"]}
         for k, v in disease_totals.items()],
        key=lambda d: d["subjects"], reverse=True,
    )

    # Rolled back up to the original disease team - one shared Expected
    # Recruitment target per team. "segments" (fine subgroups, e.g. Kidney/
    # Skin) is kept for internal bookkeeping (disease-filter-pill math), but
    # the dashboard's "Subjects enrolled by disease team" bar itself now
    # stacks by individual COHORT instead - see "cohort_segments" below.
    subgroup_to_group = {k: v["old_group_key"] for k, v in disease_totals.items()}
    by_disease_group = []
    for gk, g in combined_totals.items():
        member_keys = [k for k, og in subgroup_to_group.items() if og == gk]
        segments = sorted(
            [{"key": k, "label": disease_totals[k]["label"], "subjects": disease_totals[k]["subjects"]}
             for k in member_keys],
            key=lambda s: s["subjects"], reverse=True,
        )
        expected, partial = disease_expected(member_keys, cohort_totals)
        # Per-cohort breakdown for the stacked bar - literally the sum of
        # the "Cohorts within each disease team" card(s) for this team's
        # member subgroup(s), rolled up under whichever cohort_view group(s)
        # they map to (merged into one for PsD, one-per-subgroup for
        # everyone else - see cohort_view_group()). cohort_view_cohort_totals
        # is already deduplicated per (cv_key, cohort tag), so no subject is
        # double-counted within a single segment; a subject carrying more
        # than one cohort tag still shows up in each of those segments,
        # exactly like the cards below - the bar's own displayed total stays
        # the true, deduplicated `subjects` count above regardless.
        cv_keys_for_group = sorted(set(
            cohort_view_group(k, disease_totals[k]["label"], disease_totals[k]["old_label"])[0]
            for k in member_keys
        ))
        cohort_segments = sorted(
            [{"disease_key": v2["disease_key"], "disease_label": v2["disease_label"], "cohort": ck[1],
              "subjects": v2["subjects"], "visits": v2["visits"], "status": v2["status"],
              "expected": expected_for(v2["target_label"], ck[1])}
             for ck, v2 in cohort_view_cohort_totals.items() if v2["disease_key"] in cv_keys_for_group],
            key=lambda d: (d["disease_label"], d["cohort"].lower()),
        )
        by_disease_group.append({
            "key": gk, "label": g["label"], "subjects": len(g["subject_set"]), "visits": g["visits"],
            "status": g["status"], "expected": expected, "expected_partial": partial,
            "segments": segments, "cohort_segments": cohort_segments,
        })
    by_disease_group.sort(key=lambda d: d["subjects"], reverse=True)

    by_cohort_detail = sorted(
        [{"disease_key": v["disease_key"], "disease_label": v["disease_label"], "cohort": ck[1],
          "subjects": v["subjects"], "visits": v["visits"], "status": v["status"],
          "expected": expected_for(target_label_for(v["disease_key"], v["old_label"]), ck[1])}
         for ck, v in cohort_totals.items()],
        # Alphabetical by cohort name within each disease team.
        key=lambda d: (d["disease_label"], d["cohort"].lower()),
    )

    # The merged "Cohorts within each disease team" view - same shape as
    # by_disease/by_cohort above, built from cohort_view_disease_totals/
    # cohort_view_cohort_totals instead (Psoriasis + Psoriatic Arthritis
    # combined into one "Psoriatic Disease (PsD)" entry; everything else,
    # Lupus included, unchanged from the fine-grained view). This is what
    # the "Cohorts within each disease team" card/table renders from; the
    # by_disease/by_cohort above stay fine-grained for everything else
    # (filter pills, KPI totals, the top disease-team bar, Samples tab).
    cohort_view_by_disease = sorted(
        [{"key": k, "label": v["label"], "subjects": v["subjects"], "visits": v["visits"],
          "status": v["status"], "schedule_status": v["schedule_status"]}
         for k, v in cohort_view_disease_totals.items()],
        key=lambda d: d["subjects"], reverse=True,
    )
    cohort_view_by_cohort = sorted(
        [{"disease_key": v["disease_key"], "disease_label": v["disease_label"], "cohort": ck[1],
          "subjects": v["subjects"], "visits": v["visits"], "status": v["status"],
          "schedule_status": v["schedule_status"],
          "expected": expected_for(v["target_label"], ck[1])}
         for ck, v in cohort_view_cohort_totals.items()],
        key=lambda d: (d["disease_label"], d["cohort"].lower()),
    )

    recruitment = {
        "by_disease": by_disease,
        "by_disease_group": by_disease_group,
        "by_cohort": by_cohort_detail,
        "cohort_view": {"by_disease": cohort_view_by_disease, "by_cohort": cohort_view_by_cohort},
        "unassigned_subjects": unassigned_subjects,
        "status_order": RECRUIT_STATUS_ORDER,
        "status_labels": RECRUIT_STATUS_LABELS,
        "schedule_status_order": SCHEDULE_STATUS_ORDER,
        "schedule_status_labels": SCHEDULE_STATUS_LABELS,
        "schedule_status_available": sched_idx is not None if subj_ws is not None else False,
    }

    # ---------- Samples: what specimen types were collected, per visit ----------
    # Walks the "visits" sheet again (one row per visit, same as the
    # Technologies-tab loop above) but joins each row to its subject on the
    # "subjects" sheet (via subject_join, built above) to get the disease
    # team and cohort(s) - using the same harmonized values as the
    # Recruitment tab, rather than that row's own (older, less consistent)
    # Data_Scope/Visit_Cohort. A visit whose subject isn't found on the
    # "subjects" sheet still counts in the overall per-type totals, just not
    # in the by-disease-team/by-cohort breakdowns (see unassigned_visits).
    missing_sample_cols = [col for col, _, _ in SAMPLE_TYPE_COLUMNS if col not in col_idx]
    if missing_sample_cols:
        print(f"WARNING: expected sample-type columns not found and will be skipped: "
              f"{missing_sample_cols}", file=sys.stderr)

    sample_type_totals = {key: 0 for _, key, _ in SAMPLE_TYPE_COLUMNS}
    sample_by_disease = defaultdict(lambda: {
        "visits": 0, "label": "", "counts": {key: 0 for _, key, _ in SAMPLE_TYPE_COLUMNS},
    })
    sample_by_cohort = defaultdict(lambda: {
        "visits": 0, "disease_key": "", "disease_label": "",
        "counts": {key: 0 for _, key, _ in SAMPLE_TYPE_COLUMNS},
    })
    sample_unassigned_visits = 0

    for r in range(DATA_START_ROW, ws.max_row + 1):
        subj = ws.cell(row=r, column=subject_idx).value if subject_idx else None
        subj_clean = clean_label(subj, default=None)
        row_has_any = any(
            ws.cell(row=r, column=col_idx[col]).value is not None
            for col, _, _ in TECH_COLUMNS
            if col in col_idx
        )
        if subj_clean is None and not row_has_any:
            continue  # same "is this a real visit row" check as the tech loop above

        collected = {}
        for col, key, _ in SAMPLE_TYPE_COLUMNS:
            got = col in col_idx and is_collected(ws.cell(row=r, column=col_idx[col]).value)
            collected[key] = got
            if got:
                sample_type_totals[key] += 1

        join = subject_join.get(subj_clean) if subj_clean else None
        if join is None:
            sample_unassigned_visits += 1
            continue

        for disease_key, disease_label, _old_label, _old_group_key in derive_diseases(join["scope"]):
            d = sample_by_disease[disease_key]
            d["visits"] += 1
            d["label"] = disease_label
            for key, got in collected.items():
                if got:
                    d["counts"][key] += 1
            for tag in sorted(set(join["tags"])):
                ck = (disease_key, tag)
                c = sample_by_cohort[ck]
                c["visits"] += 1
                c["disease_key"] = disease_key
                c["disease_label"] = disease_label
                for key, got in collected.items():
                    if got:
                        c["counts"][key] += 1

    samples = {
        "sample_types": [{"key": key, "label": label} for _, key, label in SAMPLE_TYPE_COLUMNS],
        "by_type": [
            {"key": key, "label": label, "collected": sample_type_totals[key], "total_visits": n_visits}
            for _, key, label in SAMPLE_TYPE_COLUMNS
        ],
        "by_disease": sorted(
            [{"key": k, "label": v["label"], "visits": v["visits"], "counts": v["counts"]}
             for k, v in sample_by_disease.items()],
            key=lambda d: d["label"],
        ),
        "by_cohort": sorted(
            [{"disease_key": v["disease_key"], "disease_label": v["disease_label"], "cohort": ck[1],
              "visits": v["visits"], "counts": v["counts"]}
             for ck, v in sample_by_cohort.items()],
            key=lambda d: (d["disease_label"], d["cohort"].lower()),
        ),
        "unassigned_visits": sample_unassigned_visits,
        "total_visits": n_visits,
    }

    by_scope_matrix = {
        key: {scope: tech_by_scope[key].get(scope, empty_status_counts()) for scope in scopes_sorted}
        for _, key, _ in TECH_COLUMNS
    }
    by_pipeline_matrix = {
        key: {p: tech_by_pipeline[key].get(p, empty_status_counts()) for p in pipelines_sorted}
        for _, key, _ in TECH_COLUMNS
    }
    by_scope_pipeline_matrix = {
        key: {
            scope: {p: tech_by_scope_pipeline[key][scope].get(p, empty_status_counts()) for p in pipelines_sorted}
            for scope in scopes_sorted
        }
        for _, key, _ in TECH_COLUMNS
    }

    output = {
        "generated_at": datetime.datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ"),
        "source_rows": n_visits,
        "totals": {
            "subjects": len(subjects_seen),
            "visits": n_visits,
        },
        "recruitment": recruitment,
        "status_order": STATUS_ORDER,
        "status_labels": STATUS_LABELS,
        "technologies": technologies,
        "scopes": scopes_sorted,
        "pipelines": pipelines_sorted,
        "scope_disease": scope_disease,
        "by_scope": by_scope_matrix,
        "by_pipeline": by_pipeline_matrix,
        "by_scope_pipeline": by_scope_pipeline_matrix,
        "samples": samples,
    }

    with open(out_path, "w") as f:
        json.dump(output, f, indent=2)

    print(f"Wrote {out_path}")
    print(f"  subjects={len(subjects_seen)} visits={n_visits} technologies={len(technologies)}")
    print(f"  disease teams={len(by_disease)} cohorts={len(by_cohort_detail)} unassigned_subjects={unassigned_subjects}")
    print(f"  pipelines={pipelines_sorted}")
    print(f"  sample types={len(samples['sample_types'])} sample-cohorts={len(samples['by_cohort'])} "
          f"unassigned_visits={sample_unassigned_visits}")


if __name__ == "__main__":
    main()
