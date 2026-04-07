#!/usr/bin/env python3
"""
USRA Stage 1 - Build Eligibility Pairs (updated to include SupervisorDepartment)

Key update:
  - Department-based conflict checks use the union of:
      StudentDepartment(s) + SupervisorDepartment(s)

Rules enforced:
  1) Own department: reviewer cannot review if ReviewerDept overlaps (StudentDept ∪ SupervisorDept)
  2) Supervisor NSID: reviewer cannot review if listed as supervisor/co-supervisor (NSID match)
  3) Reviewer has student in competition: cannot review if StudentDepartmentList overlaps (StudentDept ∪ SupervisorDept)
  4) Other disclosed conflicts: cannot review if ConflictDepartmentList overlaps (StudentDept ∪ SupervisorDept)

Outputs:
  - USRA_eligibility_pairs.xlsx (EligibilityPairs + diagnostics)
"""

import re
import sys
import pandas as pd

# ----------------------------
# Configuration (edit as needed)
# ----------------------------
APPLICANTS_FILE = "2026 Applicants.xlsx"
REVIEWERS_FILE = "2026 USRA Reviewers.xlsx"
REVIEWERS_SHEET = "2026 Committee Membership"

OUTPUT_FILE = "USRA_eligibility_pairs.xlsx"

# Applicants columns
COL_APP_ID = "AppID"
COL_STUDENT_DEPT = "StudentDepartment"
COL_SUPERVISOR_DEPT = "SupervisorDepartment"
COL_SUPERVISOR_NSID = "SupervisorNSID"

# Reviewers columns
COL_REV_NSID = "ReviewerNSID"
COL_REV_NAME = "ReviewerName"
COL_REV_DEPT = "ReviewerDept"
COL_HAS_STUDENT = "HasStudentInCompetition"
COL_STUDENT_DEPT_LIST = "StudentDepartmentList"
COL_CONFLICT_DEPT_LIST = "ConflictDepartmentList"

# ----------------------------
# Helpers
# ----------------------------
def split_semicolon_lower(value):
    """Split semicolon-separated text into normalized lowercase tokens."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return []
    s = str(value).strip()
    if s == "" or s.lower() == "nan":
        return []
    parts = [p.strip() for p in s.split(";")]
    out = []
    for p in parts:
        p = re.sub(r"\s+", " ", p).strip().lower()
        if p:
            out.append(p)
    return out

def norm_lower(value):
    """Lowercase + trim for NSIDs and yes/no fields."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return str(value).strip().lower()

def eligible(app_row, rev_row):
    """
    Returns True if reviewer is eligible for the application, else False.
    Department-based rules use AllDepartment_arr = StudentDepartment_arr ∪ SupervisorDepartment_arr.
    """
    app_depts = set(app_row["AllDepartment_arr"])
    rev_depts = set(rev_row["ReviewerDept_arr"])

    # Rule 1: own department overlap (student OR supervisor department triggers)
    if app_depts & rev_depts:
        return False

    # Rule 2: supervisor/co-supervisor NSID conflict
    if rev_row["ReviewerNSID_norm"] in set(app_row["SupervisorNSID_arr"]):
        return False

    # Rule 3: reviewer has a student in the competition -> exclude student department(s)
    if rev_row["HasStudent_norm"] == "yes":
        if app_depts & set(rev_row["StudentDepartmentList_arr"]):
            return False

    # Rule 4: other disclosed departmental conflicts
    if app_depts & set(rev_row["ConflictDepartmentList_arr"]):
        return False

    return True

# ----------------------------
# Main
# ----------------------------
def main():
    # Read inputs
    try:
        apps = pd.read_excel(APPLICANTS_FILE, engine="openpyxl")
    except Exception as e:
        print(f"ERROR: Could not read applicants file: {APPLICANTS_FILE}\n{e}")
        sys.exit(1)

    try:
        revs = pd.read_excel(REVIEWERS_FILE, sheet_name=REVIEWERS_SHEET, engine="openpyxl")
    except Exception as e:
        print(f"ERROR: Could not read reviewers file: {REVIEWERS_FILE} (sheet: {REVIEWERS_SHEET})\n{e}")
        sys.exit(1)

    # Required columns
    required_app_cols = [COL_APP_ID, COL_STUDENT_DEPT, COL_SUPERVISOR_DEPT, COL_SUPERVISOR_NSID]
    required_rev_cols = [COL_REV_NSID, COL_REV_NAME, COL_REV_DEPT, COL_HAS_STUDENT,
                         COL_STUDENT_DEPT_LIST, COL_CONFLICT_DEPT_LIST]

    missing_app = [c for c in required_app_cols if c not in apps.columns]
    missing_rev = [c for c in required_rev_cols if c not in revs.columns]

    if missing_app:
        print(f"ERROR: Missing applicants columns: {missing_app}")
        sys.exit(1)
    if missing_rev:
        print(f"ERROR: Missing reviewers columns: {missing_rev}")
        sys.exit(1)

    # Normalize applicant fields
    apps["StudentDepartment_arr"] = apps[COL_STUDENT_DEPT].apply(split_semicolon_lower)
    apps["SupervisorDepartment_arr"] = apps[COL_SUPERVISOR_DEPT].apply(split_semicolon_lower)
    apps["SupervisorNSID_arr"] = apps[COL_SUPERVISOR_NSID].apply(split_semicolon_lower)
    apps["AppID_int"] = apps[COL_APP_ID].astype(int)

    # Union of student + supervisor departments
    def union_depts(row):
        return list(dict.fromkeys(row["StudentDepartment_arr"] + row["SupervisorDepartment_arr"]))
    apps["AllDepartment_arr"] = apps.apply(union_depts, axis=1)

    # Normalize reviewer fields
    revs["ReviewerNSID_norm"] = revs[COL_REV_NSID].apply(norm_lower)
    revs["ReviewerDept_arr"] = revs[COL_REV_DEPT].apply(split_semicolon_lower)
    revs["HasStudent_norm"] = revs[COL_HAS_STUDENT].apply(norm_lower)
    revs["StudentDepartmentList_arr"] = revs[COL_STUDENT_DEPT_LIST].apply(split_semicolon_lower)
    revs["ConflictDepartmentList_arr"] = revs[COL_CONFLICT_DEPT_LIST].apply(split_semicolon_lower)

    # Build eligibility pairs
    rows = []
    for _, a in apps.iterrows():
        for _, r in revs.iterrows():
            if eligible(a, r):
                rows.append({
                    "AppID": int(a["AppID_int"]),
                    "ReviewerNSID": r["ReviewerNSID_norm"],
                    "Eligible": "Yes"
                })

    elig_df = pd.DataFrame(rows)

    # Diagnostics
    app_counts = elig_df.groupby("AppID").size().rename("EligibleReviewerCount").reset_index()
    app_diag = apps[[COL_APP_ID, COL_STUDENT_DEPT, COL_SUPERVISOR_DEPT, COL_SUPERVISOR_NSID]].copy()
    app_diag[COL_APP_ID] = app_diag[COL_APP_ID].astype(int)
    app_diag = app_diag.merge(app_counts, on="AppID", how="left").fillna({"EligibleReviewerCount": 0})
    apps_lt2 = app_diag[app_diag["EligibleReviewerCount"] < 2].sort_values("EligibleReviewerCount")

    rev_counts = elig_df.groupby("ReviewerNSID").size().rename("EligibleAppCount").reset_index()
    rev_diag = revs[["ReviewerNSID_norm", COL_REV_NAME, COL_REV_DEPT, COL_HAS_STUDENT,
                     COL_STUDENT_DEPT_LIST, COL_CONFLICT_DEPT_LIST]].copy()
    rev_diag = rev_diag.rename(columns={"ReviewerNSID_norm": "ReviewerNSID"})
    rev_diag = rev_diag.merge(rev_counts, on="ReviewerNSID", how="left").fillna({"EligibleAppCount": 0})

    # Write output workbook
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        elig_df.sort_values(["AppID", "ReviewerNSID"]).to_excel(writer, index=False, sheet_name="EligibilityPairs")
        app_diag.sort_values("AppID").to_excel(writer, index=False, sheet_name="AppDiagnostics")
        rev_diag.sort_values("ReviewerNSID").to_excel(writer, index=False, sheet_name="ReviewerDiagnostics")
        apps_lt2.to_excel(writer, index=False, sheet_name="Apps_LT2_Eligible")

    # Console summary
    print("Stage 1 complete (updated with SupervisorDepartment).")
    print(f"Applicants: {len(apps)}")
    print(f"Reviewers: {len(revs)}")
    print(f"Eligible pairs written: {len(elig_df)}")
    print(f"Apps with <2 eligible reviewers: {len(apps_lt2)}")
    print(f"Output file: {OUTPUT_FILE}")

if __name__ == "__main__":
    main()