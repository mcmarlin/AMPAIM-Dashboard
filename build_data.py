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
    visit cohort. Cells with fewer than MIN_CELL_COUNT contributing rows are
    still reported (counts, not identities) but the script makes it easy to
    raise that floor if small-cell suppression is ever required.

Usage:
    python3 build_data.py [input.xlsx] [output.json]

Defaults: data/AMP_AIM_Dataset.xlsx -> data/dashboard.json
"""
import sys
import json
import datetime
from collections import defaultdict

import openpyxl

HEADER_ROW = 3
DATA_START_ROW = 4

# Technology (sample/assay) columns -> (key, display label)
TECH_COLUMNS = [
    ("Xenium-Slide_Sample_ID", "xenium_slide", "Xenium (Slide)"),
    ("Xenium_Sample_ID", "xenium", "Xenium"),
    ("scFFPE_Sample_ID", "scffpe", "scFFPE"),
    ("scRNAseq_Sample_ID", "scrnaseq", "scRNA-seq"),
    ("Olink-Serum_Sample_ID", "olink_serum", "Olink (Serum)"),
    ("CyTOF-Blood_Sample_ID", "cytof_blood", "CyTOF (Blood)"),
    ("Genotyping_Sample_ID", "genotyping", "Genotyping"),
    ("OMRF-Lab_Sample_ID", "omrf_lab", "OMRF Lab"),
    ("bRNAseq_Sample_ID", "brnaseq", "Bulk RNA-seq"),
    ("Xenium-NL_Sample_ID", "xenium_nl", "Xenium (Non-lesional)"),
    ("scFFPE-NL_Sample_ID", "scffpe_nl", "scFFPE (Non-lesional)"),
    ("scRNAseq-NL_Sample_ID", "scrnaseq_nl", "scRNA-seq (Non-lesional)"),
    ("Olink-Urine_Sample_ID", "olink_urine", "Olink (Urine)"),
    ("mBioSeq-stool", "mbioseq_stool", "Microbiome (Stool)"),
    ("mBioSeq-skin", "mbioseq_skin", "Microbiome (Skin)"),
    ("anti-C1Q_Sample_ID", "anti_c1q", "Anti-C1Q"),
]

STATUS_ORDER = ["completed", "pending", "not_applicable", "qc_fail", "unknown"]
STATUS_LABELS = {
    "completed": "Completed",
    "pending": "Pending / Specimen Available",
    "not_applicable": "Not Applicable / No Specimen",
    "qc_fail": "QC Fail",
    "unknown": "Unknown / Not Yet Retrieved",
}

# Visit codes that mark a subject's ENROLLMENT into a cohort (one per Visit_Type:
# Scheduled->V01, Control->VC1, Enabling->VE1, Archival->VA1, Unscheduled->VU1).
# The Visit_Cohort tag on this row is the subject's cohort assignment. Checked
# in this priority order when a subject has more than one such row.
ENROLLMENT_CODES_PRIORITY = ["V01", "VC1", "VE1", "VA1", "VU1"]

# "Disease Team" in the source is unreliable (mostly #REF! - a broken lookup
# formula upstream). Data_Scope ("<disease>-<tissue>", e.g. "SLE-KDY") is
# populated and reliable, so the disease/team grouping is derived from its
# prefix instead. EDIT THIS if your actual disease-team names differ, or add
# an explicit Subject_ID -> team override if Data_Scope isn't authoritative
# for some subjects.
DISEASE_LABELS = {
    "SLE": "Lupus (SLE)",
    "RA": "Rheumatoid Arthritis (RA)",
    "PsD": "Psoriatic Disease (PsD)",
    "SjD": "Sjögren's Disease (SjD)",
    "SSc": "Systemic Sclerosis (SSc)",
    "SLE/PsD": "SLE / PsD (combined)",
}


def derive_disease(scope_label):
    """('SLE-KDY' -> ('sle', 'Lupus (SLE)')); unknown/blank scope -> ('unknown', 'Unknown')."""
    if not scope_label or scope_label == "Unknown":
        return "unknown", "Unknown"
    prefix = scope_label.split("-")[0].strip()
    label = DISEASE_LABELS.get(prefix, prefix)
    key = prefix.lower().replace("/", "_").replace(" ", "_")
    return key, label


def classify(value):
    """Collapse the many raw status strings into 5 reportable buckets."""
    if value is None:
        return "unknown"
    s = str(value).strip()
    if s == "" or s in ("None", "#REF!", "[not yet retrieved]"):
        return "unknown"
    low = s.lower()
    if "qc fail" in low or "qc-fail" in low or "failed qc" in low:
        return "qc_fail"
    if "not applicable" in low or "no specimen" in low or "not available" in low:
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
    if value is None:
        return default
    s = str(value).strip()
    if s == "" or s in ("None", "#REF!"):
        return default
    if s.startswith("[") and s.endswith("]"):
        inner = s[1:-1]
        # some cells hold multiple bracketed tags back-to-back, e.g.
        # "[Cohort 1a: PsO][Cohort 3: PsO to PsA Risk]" -> split on the
        # internal "][" boundary instead of leaving it in the label.
        parts = [p.strip() for p in inner.split("][") if p.strip()]
        s = ", ".join(parts) if parts else default
    if s.lower() == "not yet retrieved":
        s = "Not yet retrieved"
    return s


def empty_status_counts():
    return {k: 0 for k in STATUS_ORDER}


def main():
    in_path = sys.argv[1] if len(sys.argv) > 1 else "data/AMP_AIM_Dataset.xlsx"
    out_path = sys.argv[2] if len(sys.argv) > 2 else "data/dashboard.json"

    wb = openpyxl.load_workbook(in_path, data_only=True)
    ws = wb[wb.sheetnames[0]]

    headers = [ws.cell(row=HEADER_ROW, column=c).value for c in range(1, ws.max_column + 1)]
    col_idx = {h: i + 1 for i, h in enumerate(headers) if h}

    missing = [col for col, _, _ in TECH_COLUMNS if col not in col_idx]
    if missing:
        print(f"WARNING: expected columns not found and will be skipped: {missing}", file=sys.stderr)

    subject_idx = col_idx.get("Subject_ID")
    data_scope_idx = col_idx.get("Data_Scope")
    visit_cohort_idx = col_idx.get("Visit_Cohort")
    visit_code_idx = col_idx.get("Visit_Code")

    subjects_seen = set()
    n_visits = 0
    # subject -> list of (visit_code_raw, cohort_label, scope_label) for every row
    subj_rows = defaultdict(list)

    # tech key -> overall status counts
    tech_totals = {key: empty_status_counts() for _, key, _ in TECH_COLUMNS}
    # tech key -> {scope_label: status_counts}
    tech_by_scope = {key: defaultdict(empty_status_counts) for _, key, _ in TECH_COLUMNS}
    # tech key -> {cohort_label: status_counts}
    tech_by_cohort = {key: defaultdict(empty_status_counts) for _, key, _ in TECH_COLUMNS}

    scopes_seen = set()
    cohorts_seen = set()

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
        cohort_label = clean_label(ws.cell(row=r, column=visit_cohort_idx).value if visit_cohort_idx else None)
        scopes_seen.add(scope_label)
        cohorts_seen.add(cohort_label)

        if subj_clean:
            code_raw = ws.cell(row=r, column=visit_code_idx).value if visit_code_idx else None
            code_raw = str(code_raw).strip() if code_raw is not None else ""
            subj_rows[subj_clean].append((code_raw, cohort_label, scope_label))

        for col, key, _ in TECH_COLUMNS:
            if col not in col_idx:
                continue
            val = ws.cell(row=r, column=col_idx[col]).value
            status = classify(val)
            tech_totals[key][status] += 1
            tech_by_scope[key][scope_label][status] += 1
            tech_by_cohort[key][cohort_label][status] += 1

    def totals_block(counts):
        total = sum(counts.values())
        completed = counts["completed"]
        # pct_complete: completed / all visits (conservative, denominator = every visit row)
        pct = round(100 * completed / total, 1) if total else 0.0
        # pct_eligible: completed / (completed + pending + qc_fail), i.e. excludes
        # cells that were never applicable/ordered or never populated for this
        # subject-visit. This reads as "of the samples actually in the pipeline
        # for this technology, how many are done" and is usually the more honest
        # progress number when a technology only applies to a subset of scopes.
        eligible = counts["completed"] + counts["pending"] + counts["qc_fail"]
        pct_elig = round(100 * completed / eligible, 1) if eligible else None
        return {
            "counts": counts,
            "total": total,
            "pct_complete": pct,
            "eligible": eligible,
            "pct_eligible": pct_elig,
        }

    technologies = []
    for _, key, label in TECH_COLUMNS:
        block = totals_block(tech_totals[key])
        block["key"] = key
        block["label"] = label
        technologies.append(block)
    technologies.sort(key=lambda t: t["pct_complete"], reverse=True)

    scopes_sorted = sorted(s for s in scopes_seen if s != "Unknown") + (
        ["Unknown"] if "Unknown" in scopes_seen else []
    )
    cohorts_sorted = sorted(c for c in cohorts_seen if c != "Unknown") + (
        ["Unknown"] if "Unknown" in cohorts_seen else []
    )

    # ---------- Recruitment: subjects & visits per disease team / cohort ----------
    def pick_enrollment(rows):
        """Pick the (cohort, scope) for a subject from their enrollment-code row(s)."""
        by_code = {}
        for code, cohort, scope in rows:
            if code not in ENROLLMENT_CODES_PRIORITY:
                continue
            if code not in by_code or (by_code[code][0] == "Unknown" and cohort != "Unknown"):
                by_code[code] = (cohort, scope)
        for code in ENROLLMENT_CODES_PRIORITY:
            if code in by_code and by_code[code][0] != "Unknown":
                return by_code[code][0], by_code[code][1]
        for code in ENROLLMENT_CODES_PRIORITY:
            if code in by_code:
                return by_code[code][0], by_code[code][1]
        return None

    disease_totals = defaultdict(lambda: {"subjects": 0, "visits": 0, "label": ""})
    cohort_totals = defaultdict(lambda: {"subjects": 0, "visits": 0, "disease_key": "", "disease_label": ""})
    unassigned_subjects = 0

    for subj, rows in subj_rows.items():
        picked = pick_enrollment(rows)
        n_subject_visits = len(rows)
        if picked is None:
            unassigned_subjects += 1
            continue
        cohort, scope = picked
        disease_key, disease_label = derive_disease(scope)
        disease_totals[disease_key]["subjects"] += 1
        disease_totals[disease_key]["visits"] += n_subject_visits
        disease_totals[disease_key]["label"] = disease_label
        ck = (disease_key, cohort)
        cohort_totals[ck]["subjects"] += 1
        cohort_totals[ck]["visits"] += n_subject_visits
        cohort_totals[ck]["disease_key"] = disease_key
        cohort_totals[ck]["disease_label"] = disease_label

    by_disease = sorted(
        [{"key": k, "label": v["label"], "subjects": v["subjects"], "visits": v["visits"]}
         for k, v in disease_totals.items()],
        key=lambda d: d["subjects"], reverse=True,
    )
    by_cohort_detail = sorted(
        [{"disease_key": v["disease_key"], "disease_label": v["disease_label"], "cohort": ck[1],
          "subjects": v["subjects"], "visits": v["visits"]}
         for ck, v in cohort_totals.items()],
        key=lambda d: (d["disease_label"], -d["subjects"]),
    )
    recruitment = {
        "by_disease": by_disease,
        "by_cohort": by_cohort_detail,
        "unassigned_subjects": unassigned_subjects,
    }

    by_scope_matrix = {
        key: {scope: tech_by_scope[key].get(scope, empty_status_counts()) for scope in scopes_sorted}
        for _, key, _ in TECH_COLUMNS
    }
    by_cohort_matrix = {
        key: {cohort: tech_by_cohort[key].get(cohort, empty_status_counts()) for cohort in cohorts_sorted}
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
        "cohorts": cohorts_sorted,
        "by_scope": by_scope_matrix,
        "by_cohort": by_cohort_matrix,
    }

    with open(out_path, "w") as f:
        json.dump(output, f, indent=2)

    print(f"Wrote {out_path}")
    print(f"  subjects={len(subjects_seen)} visits={n_visits} technologies={len(technologies)}")
    print(f"  disease teams={len(by_disease)} cohorts={len(by_cohort_detail)} unassigned_subjects={unassigned_subjects}")


if __name__ == "__main__":
    main()
