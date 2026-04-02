import os
import csv
from datetime import datetime
import pandas as pd

# ============================================================
# CONFIGURATION
# ============================================================

ORGANIZED_ROOT = r"C:\Users\ljh440\USRA 2026 Applications\USRA 2026_Organized"
ROSTER_XLSX    = r"C:\Users\ljh440\USRA 2026 Applications\Inputs\2026 Tri-Agency USRA Student Application List.xlsx"

REPORT_PREFIX = "USRA_2026_folder_and_file_audit"

FALLBACK_FOLDER_NAME = "_NO_STUDENT_FOLDER_FOUND"

# ---- Excel column names (case-insensitive) ----
FIRST_NAME_COLUMNS = ["Student First Name", "First Name", "Given Name"]
LAST_NAME_COLUMNS  = ["Student Last Name", "Last Name", "Surname"]

# ---- Exact filename prefixes (case-insensitive) ----
REQUIRED_FILES = {
    "FacultyApplication": "2026 usra faculty application form_",
    "StudentApplication": "2026 usra student application form_",
}

TRANSCRIPTS = {
    "USask": "2026 usra usask transcript_",
    "NonUSask": "2026 usra non-usask transcript_",
}

# ============================================================
# HELPERS
# ============================================================

def normalize(s):
    return s.lower().strip()


def sanitize_name(s):
    return " ".join(s.split()).strip()


def list_files(folder):
    return [
        f for f in os.listdir(folder)
        if os.path.isfile(os.path.join(folder, f))
    ]


def has_prefix(files, prefix):
    prefix = normalize(prefix)
    for f in files:
        if normalize(f).startswith(prefix):
            return True
    return False


def pick_column(df, candidates):
    for c in candidates:
        c_norm = c.lower()
        if c_norm in df.columns:
            return c_norm
    raise ValueError(f"Missing required column: one of {candidates}")

# ============================================================
# MAIN LOGIC
# ============================================================

def main():
    if not os.path.isdir(ORGANIZED_ROOT):
        raise SystemExit(f"Organized root not found: {ORGANIZED_ROOT}")
    if not os.path.isfile(ROSTER_XLSX):
        raise SystemExit(f"Roster Excel not found: {ROSTER_XLSX}")

    df = pd.read_excel(ROSTER_XLSX)
    df.columns = [c.strip().lower() for c in df.columns]

    col_fn = pick_column(df, FIRST_NAME_COLUMNS)
    col_ln = pick_column(df, LAST_NAME_COLUMNS)

    roster_names = []
    for _, r in df.iterrows():
        fn = sanitize_name(str(r[col_fn]))
        ln = sanitize_name(str(r[col_ln]))
        if fn and ln:
            roster_names.append(f"{fn} {ln}")

    roster_set = set(roster_names)

    folder_names = {
        d for d in os.listdir(ORGANIZED_ROOT)
        if os.path.isdir(os.path.join(ORGANIZED_ROOT, d))
    }

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_csv = os.path.join(
        ORGANIZED_ROOT,
        f"{REPORT_PREFIX}_{timestamp}.csv"
    )

    rows = []

    # ---- Excel → Folder → Files ----
    for name in sorted(roster_names):
        folder = os.path.join(ORGANIZED_ROOT, name)
        exists = os.path.isdir(folder)

        if not exists:
            rows.append([
                name, "NO", "", "", "", "", "MissingFolder"
            ])
            continue

        files = list_files(folder)

        has_faculty = has_prefix(files, REQUIRED_FILES["FacultyApplication"])
        has_student = has_prefix(files, REQUIRED_FILES["StudentApplication"])
        has_usask   = has_prefix(files, TRANSCRIPTS["USask"])
        has_non     = has_prefix(files, TRANSCRIPTS["NonUSask"])

        issues = []
        if not has_faculty:
            issues.append("MissingFacultyApplication")
        if not has_student:
            issues.append("MissingStudentApplication")
        if not (has_usask or has_non):
            issues.append("MissingTranscript")

        status = "Complete" if not issues else ";".join(issues)

        rows.append([
            name,
            "YES",
            "YES" if has_faculty else "NO",
            "YES" if has_student else "NO",
            "YES" if has_usask else "NO",
            "YES" if has_non else "NO",
            status
        ])

    # ---- Extra folders not in roster ----
    for folder in sorted(folder_names - roster_set):
        if folder == FALLBACK_FOLDER_NAME:
            rows.append([folder, "N/A", "N/A", "N/A", "N/A", "N/A", "FallbackFolder"])
        else:
            rows.append([folder, "NO", "", "", "", "", "ExtraFolder_NotInRoster"])

    # ---- Write CSV ----
    with open(out_csv, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow([
            "NameOrFolder",
            "FolderExists",
            "FacultyApplicationForm",
            "StudentApplicationForm",
            "USaskTranscript",
            "NonUSaskTranscript",
            "Status"
        ])
        w.writerows(rows)

    print(f"\nAudit report written to:\n{out_csv}")

if __name__ == "__main__":
    main()
