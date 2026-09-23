#!/usr/bin/env python3
"""
generate_cohort_reference.py

One-off helper (not part of the regular dashboard build) that lists every
Disease Team x Cohort combination found in the new "subjects" tab, in the
harmonized Dashboard_Label format, so the recruitment-target spreadsheet
(data/Target_Recruitment_Numbers.xlsx) can be updated to match.

"Disease Team" here uses whatever label build_data.py's target_label_for()
actually looks cohort-level targets up by - which for most disease teams is
still the ROLLED-UP team name (e.g. "Rheumatoid Arthritis (RA)"), but for
Lupus is now the fine, per-tissue label ("Lupus Kidney" / "Lupus Skin"
separately, not one shared "Lupus (SLE)" row), so their same-named cohorts
(e.g. both have an "Enabling Case" cohort) can get different targets. See
TARGET_LOOKUP_OVERRIDE in build_data.py to add another disease team to this
list, or COHORT_VIEW_MERGE if a team should instead be COMBINED on the
"Cohorts within each disease team" dashboard view (as Psoriasis/Psoriatic
Arthritis now are - that combine doesn't change how their targets are
looked up, only how they're displayed, so it needs no entry here).
"Cohort" is each individual Dashboard_Label tag. Where a (Disease Team,
Cohort) pair matches an existing row in Target_Recruitment_Numbers.xlsx
EXACTLY (case-insensitive), its Expected value is carried over automatically
and noted; every other row is left blank for manual entry. "Current
Subjects" is a reference-only count (enrolled/enabling + archival) from the
latest data pull - not written back into the dashboard itself.

Usage:
    python3 automation/generate_cohort_reference.py \
        [dataset.xlsx] [target_recruitment.xlsx] [output.xlsx]

Defaults match build_data.py's own defaults.
"""
import os
import sys
from collections import defaultdict

import openpyxl
from openpyxl.styles import Font

# build_data.py lives in the project root, one directory up from this
# script (automation/). Python only puts THIS script's own directory on
# sys.path, not its parent - so "import build_data" fails with
# "ModuleNotFoundError: No module named 'build_data'" unless the project
# root is added explicitly here, regardless of what directory you run
# this command from.
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import build_data as bd


def build_rows(dataset_path, target_path):
    bd.EXPECTED_RECRUITMENT = bd.load_expected_recruitment(target_path)

    wb = openpyxl.load_workbook(dataset_path, data_only=True)
    subj_ws = wb["subjects"]

    subj_headers = [subj_ws.cell(row=bd.HEADER_ROW, column=c).value for c in range(1, subj_ws.max_column + 1)]
    subj_col_idx = {h: i + 1 for i, h in enumerate(subj_headers) if h}
    sid_idx = subj_col_idx["Subject_ID"]
    sscope_idx = subj_col_idx["Data_Scope"]
    slabel_idx = subj_col_idx["Dashboard_Label"]
    stype_idx = subj_col_idx["Subject Type"]

    combo_subjects = defaultdict(set)  # (old_label, cohort_tag) -> set of subject ids

    for r in range(bd.DATA_START_ROW, subj_ws.max_row + 1):
        sid = bd.clean_label(subj_ws.cell(row=r, column=sid_idx).value, default=None)
        if sid is None:
            continue
        type_raw = subj_ws.cell(row=r, column=stype_idx).value
        status = bd.SUBJECT_STATUS_BY_TYPE.get(bd.clean_label(type_raw, default=""))
        if status is None:
            continue  # Pre-Participation / unrecognized - not placed in a cohort

        scope = bd.clean_label(subj_ws.cell(row=r, column=sscope_idx).value)
        tags = bd.split_cohort_tags(subj_ws.cell(row=r, column=slabel_idx).value)

        target_labels_seen = set()
        for disease_key, _disease_label, old_label, _old_group_key in bd.derive_diseases(scope):
            target_label = bd.target_label_for(disease_key, old_label)
            if target_label in target_labels_seen:
                continue
            target_labels_seen.add(target_label)
            for tag in sorted(set(tags)):
                combo_subjects[(target_label, tag)].add(sid)

    rows = []
    for (target_label, cohort), sids in combo_subjects.items():
        expected = bd.expected_for(target_label, cohort)
        notes = "Carried over from old sheet (exact name match)" if expected is not None else ""
        rows.append({
            "Disease Team": target_label,
            "Cohort": cohort,
            "Expected": expected if expected is not None else None,
            "Notes": notes,
            "Current Subjects": len(sids),
        })

    rows.sort(key=lambda r: (r["Disease Team"], r["Cohort"].lower()))
    return rows


def write_xlsx(rows, out_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    headers = ["Disease Team", "Cohort", "Expected", "Notes", "Current Subjects"]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for r in rows:
        ws.append([r[h] for h in headers])
    widths = [26, 24, 10, 42, 16]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = w
    wb.save(out_path)


def main():
    dataset_path = sys.argv[1] if len(sys.argv) > 1 else "data/AMP-AIM_Dataset_Weekly_Update.xlsx"
    target_path = sys.argv[2] if len(sys.argv) > 2 else "data/Target_Recruitment_Numbers.xlsx"
    out_path = sys.argv[3] if len(sys.argv) > 3 else "data/Cohort_Reference_New_Format.xlsx"

    rows = build_rows(dataset_path, target_path)
    write_xlsx(rows, out_path)
    print(f"Wrote {out_path} ({len(rows)} Disease Team x Cohort combinations)")


if __name__ == "__main__":
    main()
