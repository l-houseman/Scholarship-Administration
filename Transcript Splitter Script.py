#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
USRA 2026 — Split & name consolidated USask transcripts

Features:
  - Pre-flight report-only: checks expected folders under ORGANIZED_ROOT, writes CSVs
  - Diagnostics: lists PDF↔Excel set differences
  - Header cleanup: removes trailing fields (e.g., Student Number / Gender / Degree) from extracted name line
  - Canonical casing: normalizes only when name is all-caps or all-lower
  - Split routing: when a student's folder is missing, routes output to a fallback tree and logs it

Conventions:
  - Detect transcript starts via header: "Name: Last, First" + "Student Number: ####"
  - Append legend (page 1) to end of each student's output
  - Use Excel roster (Student Number -> First/Last) for canonical "Firstname Lastname"
  - Normal destination: ORGANIZED_ROOT/<Firstname Lastname>/
  - Filenames: "{YEAR} USRA USask Transcript_<Student>.pdf"
"""

import os
import re
import csv
from datetime import datetime

import pypdf                  # pip install pypdf
import pdfplumber             # pip install pdfplumber
import pandas as pd           # pip install pandas openpyxl
from rapidfuzz import process, fuzz  # pip install rapidfuzz


# ============================== CONFIG ==============================
# --- Inputs/Outputs ---
CONSOLIDATED_PDF = r"\USRA 2026 Applications\Inputs\2026 Tri-Agency USRA USask Transcripts.pdf"
ORGANIZED_ROOT   = r"\USRA 2026 Applications\USRA 2026_Organized"
ROSTER_XLSX      = r"\USRA 2026 Applications\Inputs\2026 Tri-Agency USRA USask Student Transcript Info.xlsx"

YEAR_LABEL       = "2026"
OUTPUT_TEMPLATE  = "{year} USRA USask Transcript_{student}.pdf"

# --- Execution switches ---
DRY_RUN        = False   # False → PDFs written
PREFLIGHT_ONLY = False   # True → run preflight/diagnostics then exit
DIAGNOSTICS    = True

# Preflight behavior: report-only mode (folders remain unchanged)
CREATE_FOLDERS_IN_PREFLIGHT = False

# Split behavior:
# - If canonical folder exists → write there
# - If missing → route to fallback tree (and log it)
CREATE_FOLDERS_IN_SPLIT     = False
ROUTE_MISSING_TO_FALLBACK   = True
FALLBACK_FOLDER_NAME        = "_NO_STUDENT_FOLDER_FOUND"

# If fallback routing is disabled and folder missing, skipping controlled here
REQUIRE_EXISTING_FOLDERS    = True

# --- Excel column name mapping (case-insensitive) ---
MAPPING = {
    "student_number": ["Student Number", "StudentNumber", "ID", "Student ID"],
    "first_name":     ["Student First Name", "First Name", "FirstName", "Given Name"],
    "last_name":      ["Student Last Name", "Last Name", "LastName", "Surname", "Family Name"],
}

# --- Header regex observed in your PDF ---
RE_STUDENT_NUM = re.compile(r"(?i)\bStudent\s*Number\s*:\s*([0-9]{5,})")
RE_NAME        = re.compile(r"(?i)\bName\s*:\s*([^\n\r,]+)\s*,\s*([^\n\r]+)")
# ====================================================================


def die(msg):
    raise SystemExit(f"[FATAL] {msg}")


def ensure_dir(p):
    if not os.path.isdir(p):
        os.makedirs(p, exist_ok=True)


def sanitize_person_name(s: str) -> str:
    if not s:
        return ""
    t = " ".join(s.strip().split())
    for ch in '<>:"/\\|?*':
        t = t.replace(ch, "-")
    return t.strip(" .")


def unique_path(dst: str) -> str:
    if not os.path.exists(dst):
        return dst
    root, ext = os.path.splitext(dst)
    i = 1
    while True:
        cand = f"{root} ({i}){ext}"
        if not os.path.exists(cand):
            return cand
        i += 1


def clean_extracted_field(s: str) -> str:
    """
    Removes common PDF text extraction artifacts like **, stray asterisks, and extra whitespace.
    """
    if not s:
        return ""
    t = re.sub(r"[*_]+", "", s)
    t = " ".join(t.split())
    return t.strip()


def clean_detected_first(first: str) -> str:
    """
    Extracted header lines can contain trailing fields on the same line, e.g.
    'Zachary James Student Number: ... Gender: ... Degree: ...'
    Keep only the given-name portion.
    """
    if not first:
        return ""
    parts = re.split(r"(?i)\b(Student\s*Number|Gender|Degree)\b\s*:\s*", first, maxsplit=1)
    return parts[0].strip()


def normalize_name_case(name: str) -> str:
    """
    Normalize casing only when clearly unintentional (all caps or all lower).
    Keeps hyphens/apostrophes and handles common particles.
    """
    if not name:
        return ""
    s = " ".join(name.split()).strip()

    if not any(ch.isalpha() for ch in s):
        return s
    if not (s.isupper() or s.islower()):
        return s

    particles = {"de", "del", "della", "der", "den", "van", "von", "da", "di", "la", "le", "du", "st", "st."}

    def cap_token(tok: str) -> str:
        if not tok:
            return tok
        low = tok.lower()

        if low in particles:
            return low

        if low.startswith("o'") and len(low) > 2:
            return "O'" + low[2:].capitalize()

        if low.startswith("mc") and len(low) > 2:
            return "Mc" + low[2:].capitalize()

        return low.capitalize()

    out = []
    for i, token in enumerate(re.split(r"\s+", s)):
        parts = token.split("-")
        parts = [cap_token(p) for p in parts]
        rebuilt = "-".join(parts)

        if i == 0 and rebuilt in particles:
            rebuilt = rebuilt.capitalize()

        out.append(rebuilt)

    return " ".join(out)


def read_roster(xlsx_path: str):
    if not os.path.isfile(xlsx_path):
        die(f"Roster Excel not found: {xlsx_path}")

    df = pd.read_excel(xlsx_path, engine="openpyxl")
    df.columns = [c.strip().lower() for c in df.columns]

    def pick(colset):
        for want in colset:
            c = want.strip().lower()
            if c in df.columns:
                return c
        return None

    col_num = pick(MAPPING["student_number"])
    col_fn  = pick(MAPPING["first_name"])
    col_ln  = pick(MAPPING["last_name"])
    if not (col_num and col_fn and col_ln):
        die(f"Roster missing required columns. Found: {list(df.columns)} ; need: {MAPPING}")

    by_num = {}     # '11362014' -> 'Firstname Lastname'
    canon = set()
    for _, row in df.iterrows():
        num = str(row[col_num]).strip()
        fn  = str(row[col_fn]).strip()
        ln  = str(row[col_ln]).strip()
        if not (num and fn and ln):
            continue
        full = sanitize_person_name(normalize_name_case(f"{fn} {ln}"))
        by_num[num] = full
        canon.add(full)

    canon_list = sorted(canon)
    return by_num, canon_list


def detect_starts(pl_pdf):
    """
    Return list of start entries:
      [{'page': idx, 'student_number': num, 'last': ln, 'first': fn}]
    """
    starts = []
    for i in range(len(pl_pdf.pages)):
        txt = pl_pdf.pages[i].extract_text() or ""
        mnum  = RE_STUDENT_NUM.search(txt)
        mname = RE_NAME.search(txt)
        if mnum and mname:
            snum = mnum.group(1).strip()

            last_raw  = clean_extracted_field(mname.group(1))
            first_raw = clean_extracted_field(mname.group(2))

            last  = last_raw.strip()
            first = clean_detected_first(first_raw)

            starts.append({"page": i, "student_number": snum, "last": last, "first": first})

    # de-dup safeguard
    starts = sorted({(s["page"], s["student_number"]): s for s in starts}.values(), key=lambda x: x["page"])
    if not starts:
        die("No transcript start pages detected. Check header regex.")
    if starts[0]["page"] == 0:
        die("First detected start is page 1 (legend). Check header regex.")
    return starts


def canonical_name(detected_first, detected_last, student_number, roster_by_num, canon_names):
    """
    Prefer roster mapping by student number; otherwise fall back to detected,
    then optionally fuzzy-match to canonical names for stabilization.
    """
    if student_number and student_number in roster_by_num:
        return normalize_name_case(roster_by_num[student_number])

    raw = sanitize_person_name(normalize_name_case(f"{detected_first} {detected_last}"))
    if not canon_names:
        return raw
    best = process.extractOne(raw, canon_names, scorer=fuzz.WRatio, score_cutoff=92)
    return best[0] if best else raw


def preflight_and_diagnostics(organized_root, roster_by_num, starts, canon_names, ts, verbose=True):
    """
    - Build sets: pdf_nums, roster_nums
    - For each PDF student, compute canonical folder path, check existence
    - Write preflight report + diagnostics CSVs
    """
    pdf_nums    = set()
    roster_nums = set(roster_by_num.keys())

    pre_rows = []
    created_count = 0
    missing_count = 0

    for s in starts:
        num = s["student_number"]
        pdf_nums.add(num)

        detected_name_lf = f"{s['last']}, {s['first']}".strip().strip(",")

        canonical = canonical_name(s["first"], s["last"], num, roster_by_num, canon_names)
        canonical = sanitize_person_name(canonical)

        name_empty = (canonical.strip() == "")
        folder_component = canonical if canonical else "_UNNAMED"
        folder = os.path.join(organized_root, folder_component)
        exists = os.path.isdir(folder)
        in_roster = (num in roster_by_num)

        status = []
        status.append("InRoster" if in_roster else "NotInRoster")
        if name_empty:
            status.append("NameEmpty")

        if exists:
            status.append("FolderOK")
        else:
            status.append("FolderMissing")
            missing_count += 1
            if CREATE_FOLDERS_IN_PREFLIGHT:
                ensure_dir(folder)
                created_count += 1
                status.append("FolderAutoCreated")
                exists = os.path.isdir(folder)

        pre_rows.append([
            num, detected_name_lf, canonical, in_roster, folder, exists, ";".join(status)
        ])

    # diagnostics sets
    pdf_not_in_excel = sorted(pdf_nums - roster_nums)
    excel_not_in_pdf = sorted(roster_nums - pdf_nums)

    # Write CSVs (utf-8-sig for Excel)
    pre_path = os.path.join(organized_root, f"preflight_report_{ts}.csv")
    d1_path  = os.path.join(organized_root, f"diagnostics_pdf_not_in_excel_{ts}.csv")
    d2_path  = os.path.join(organized_root, f"diagnostics_excel_not_in_pdf_{ts}.csv")

    with open(pre_path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["StudentNumber","DetectedName_LastFirst","CanonicalName","InRoster","FolderPath","FolderExists","Status"])
        w.writerows(pre_rows)

    with open(d1_path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["StudentNumber_in_PDF_but_NOT_in_Excel"])
        for n in pdf_not_in_excel:
            w.writerow([n])

    with open(d2_path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["StudentNumber_in_Excel_but_NOT_in_PDF"])
        for n in excel_not_in_pdf:
            w.writerow([n])

    if verbose and DIAGNOSTICS:
        print("\n=== Diagnostics ===")
        print(f"PDF → unique students: {len(pdf_nums)}")
        print(f"Excel → unique students: {len(roster_nums)}")
        print(f"In PDF but not in Excel: {len(pdf_not_in_excel)}  (see {d1_path})")
        print(f"In Excel but not in PDF: {len(excel_not_in_pdf)}  (see {d2_path})")
        print("\n=== Pre-flight folders ===")
        print(f"Missing folders detected: {missing_count}")
        if CREATE_FOLDERS_IN_PREFLIGHT:
            print(f"Folders auto-created:    {created_count}")
        print(f"Pre-flight report:        {pre_path}")

    return {
        "pdf_nums": pdf_nums,
        "roster_nums": roster_nums,
        "pdf_not_in_excel": pdf_not_in_excel,
        "excel_not_in_pdf": excel_not_in_pdf,
        "preflight_report_path": pre_path,
        "auto_created": created_count,
        "missing": missing_count
    }


def main():
    # sanity checks
    if not os.path.isfile(CONSOLIDATED_PDF):
        die(f"Consolidated PDF not found: {CONSOLIDATED_PDF}")
    if not os.path.isdir(ORGANIZED_ROOT):
        die(f"Organized root not found: {ORGANIZED_ROOT}")

    # roster and canonical names
    roster_by_num, canon_names = read_roster(ROSTER_XLSX)

    # detect starts and build segments
    with pdfplumber.open(CONSOLIDATED_PDF) as pl:
        starts = detect_starts(pl)

        segments = []
        for idx, s in enumerate(starts):
            start_p = s["page"]
            end_p   = (starts[idx + 1]["page"] - 1) if (idx + 1 < len(starts)) else (len(pl.pages) - 1)
            segments.append({
                "start": start_p,
                "end": end_p,
                "student_number": s["student_number"],
                "detected_first": s["first"],
                "detected_last": s["last"],
            })

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")

    # ---------- Pre-flight + diagnostics ----------
    _ = preflight_and_diagnostics(
        organized_root=ORGANIZED_ROOT,
        roster_by_num=roster_by_num,
        starts=starts,
        canon_names=canon_names,
        ts=ts,
        verbose=True
    )

    if PREFLIGHT_ONLY:
        print("\nPREFLIGHT_ONLY=True → Skipping split/write; exiting now.")
        return

    # ---------- Split/write ----------
    reader = pypdf.PdfReader(CONSOLIDATED_PDF)
    legend_page = reader.pages[0]  # page 1 legend

    fallback_root = os.path.join(ORGANIZED_ROOT, FALLBACK_FOLDER_NAME)
    if ROUTE_MISSING_TO_FALLBACK:
        ensure_dir(fallback_root)

    log_path = os.path.join(ORGANIZED_ROOT, f"USRA_2026_split_log_{ts}.csv")
    with open(log_path, "w", newline="", encoding="utf-8-sig") as lf:
        lw = csv.writer(lf)
        lw.writerow(["StartPage","EndPage","StudentNumber","DetectedName","FinalStudentName","DestFolder","OutputFile","Action","Notes"])

        print("\n=== Planned outputs ===")
        for seg in segments:
            s_first = seg["detected_first"]
            s_last  = seg["detected_last"]
            s_num   = seg["student_number"]

            final_name = canonical_name(s_first, s_last, s_num, roster_by_num, canon_names)
            final_name = sanitize_person_name(final_name)

            folder_name = sanitize_person_name(final_name)
            dest_folder = os.path.join(ORGANIZED_ROOT, folder_name) if folder_name else ""
            folder_exists = os.path.isdir(dest_folder) if dest_folder else False

            notes = ""
            used_fallback = False

            # Folder decision
            if folder_name and folder_exists:
                pass
            else:
                if folder_name and CREATE_FOLDERS_IN_SPLIT:
                    ensure_dir(dest_folder)
                    folder_exists = os.path.isdir(dest_folder)
                    notes = "FolderAutoCreated"
                elif ROUTE_MISSING_TO_FALLBACK:
                    used_fallback = True
                    sub = sanitize_person_name(f"{s_num}_{final_name}") if final_name else sanitize_person_name(f"{s_num}_UNKNOWN")
                    dest_folder = os.path.join(fallback_root, sub)
                    ensure_dir(dest_folder)
                    notes = "RoutedToFallback_NoStudentFolder" if folder_name else "RoutedToFallback_EmptyName"
                elif REQUIRE_EXISTING_FOLDERS:
                    out_path_tmp = os.path.join(dest_folder if dest_folder else ORGANIZED_ROOT, "N/A")
                    lw.writerow([seg["start"]+1, seg["end"]+1, s_num, f"{s_last}, {s_first}", final_name,
                                 dest_folder if dest_folder else "", out_path_tmp, "Skip_NoStudentFolder", ""])
                    continue
                else:
                    if dest_folder:
                        ensure_dir(dest_folder)
                        notes = "FolderCreated_Default"
                    else:
                        used_fallback = True
                        sub = sanitize_person_name(f"{s_num}_UNKNOWN")
                        dest_folder = os.path.join(fallback_root, sub)
                        ensure_dir(dest_folder)
                        notes = "RoutedToFallback_EmptyName"

            # Output name/path
            student_token = f"{final_name}_{s_num}" if used_fallback else final_name
            student_token = sanitize_person_name(student_token) if student_token else sanitize_person_name(f"{s_num}_UNKNOWN")
            out_name = OUTPUT_TEMPLATE.format(year=YEAR_LABEL, student=student_token)
            out_path = unique_path(os.path.join(dest_folder, out_name))

            action_tag = "Fallback" if used_fallback else "Normal"
            print(f"- pages {seg['start']+1}–{seg['end']+1}  ->  {final_name} (#{s_num})  ->  {out_name}  [{action_tag}]")

            if DRY_RUN:
                action = "Plan_Fallback" if used_fallback else "Plan"
                lw.writerow([seg["start"]+1, seg["end"]+1, s_num, f"{s_last}, {s_first}", final_name,
                             dest_folder, out_path, action, notes or "Legend will be appended"])
                continue

            try:
                writer = pypdf.PdfWriter()
                for p in range(seg["start"], seg["end"] + 1):
                    if p == 0:
                        continue  # legend excluded from student pages
                    writer.add_page(reader.pages[p])
                writer.add_page(legend_page)

                with open(out_path, "wb") as out_f:
                    writer.write(out_f)

                action = "Written_Fallback" if used_fallback else "Written"
                lw.writerow([seg["start"]+1, seg["end"]+1, s_num, f"{s_last}, {s_first}", final_name,
                             dest_folder, out_path, action, (notes + ";" if notes else "") + "Legend appended"])
            except Exception as ex:
                lw.writerow([seg["start"]+1, seg["end"]+1, s_num, f"{s_last}, {s_first}", final_name,
                             dest_folder, out_path, "Error", str(ex)])

    print(f"\nSplit log written: {log_path}")
    if DRY_RUN:
        print("DRY_RUN=True — no PDFs were created. Set DRY_RUN=False for a real run.")


if __name__ == "__main__":
    main()