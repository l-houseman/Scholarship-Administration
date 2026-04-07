#!/usr/bin/env python3
"""
USRA Stage 2 - Assign reviewers (2/app), enforce 10–12 load, prefer CIHR/SSHRC stream eligibility,
and export names (StudentName, Reviewer1Name, Reviewer2Name) automatically.

Inputs (same folder recommended):
  - USRA_eligibility_pairs.xlsx  (sheet: 'EligibilityPairs')
  - 2026 Applicants_assignments protocol.xlsx  (Applicants; sheet name default)
  - 2026 USRA Reviewers_conflicts.xlsx (sheet: '2026 Committee Membership')

Outputs:
  - USRA_assignments.xlsx
    * Assignments_ByApp (AppID, StudentName, AwardStream, Reviewer1NSID, Reviewer1Name, Reviewer2NSID, Reviewer2Name)
    * Assignments_ByReviewer (ReviewerNSID, ReviewerName, AssignedAppCount, AssignedAppIDs)
    * LoadSummary (ReviewerNSID, ReviewerName, AssignedAppCount, Within10to12)
    * Exceptions (fallbacks, infeasibilities, out-of-range loads)

Notes:
  - CIHR/SSHRC apps: try to assign BOTH reviewers with that stream; if not enough, fill remaining slot(s) from any COI-eligible reviewer and flag in Exceptions.
  - NSERC apps: any COI-eligible reviewer is fine.
"""

import re
import sys
import random
from collections import defaultdict
import pandas as pd

# ----------------------------
# Configuration
# ----------------------------
ELIG_FILE = "USRA_eligibility_pairs.xlsx"
ELIG_SHEET = "EligibilityPairs"

APPS_FILE = "2026 Applicants.xlsx"    # has 'AppID', 'StudentName', 'Which award are you applying for?'
REVS_FILE = "2026 USRA Reviewers.xlsx"          # has 'ReviewerNSID', 'ReviewerName', 'StreamReview Eligibility'
REVS_SHEET = "2026 Committee Membership"

OUT_FILE = "USRA_assignments.xlsx"

COL_APP_ID = "AppID"
COL_STUDENT_NAME = "StudentName"
COL_AWARD = "Which award are you applying for?"

COL_REV_NSID = "ReviewerNSID"
COL_REV_NAME = "ReviewerName"
COL_STREAMS = "StreamReview Eligibility"

REVIEWS_PER_APP = 2
MIN_LOAD = 10
MAX_LOAD = 12

RANDOM_SEED = 20260129
random.seed(RANDOM_SEED)

# ----------------------------
# Helpers
# ----------------------------
def norm_lower(x):
    if x is None or (isinstance(x, float) and pd.isna(x)): return ""
    return str(x).strip().lower()

def split_semicolon_upper_tokens(x):
    if x is None or (isinstance(x, float) and pd.isna(x)): return set()
    return set([p.strip().upper() for p in str(x).split(";") if p.strip()])

def parse_award_stream(award_text):
    s = norm_lower(award_text)
    if "cihr" in s: return "CIHR"
    if "sshrc" in s: return "SSHRC"
    if "nserc" in s: return "NSERC"
    return "UNKNOWN"

def pick_two_reviewers(preferred_pool, fallback_pool, loads, max_load):
    selected, used_fallback = [], False
    def candidates(pool): return [r for r in pool if loads.get(r,0) < max_load and r not in selected]
    # first pick
    p1 = candidates(preferred_pool)
    if p1:
        p1.sort(key=lambda r: (loads.get(r,0), random.random()))
        selected.append(p1[0])
    else:
        f1 = candidates(fallback_pool)
        if not f1: return [], True
        f1.sort(key=lambda r: (loads.get(r,0), random.random()))
        selected.append(f1[0]); used_fallback = True
    # second pick
    p2 = candidates(preferred_pool)
    if p2:
        p2.sort(key=lambda r: (loads.get(r,0), random.random()))
        selected.append(p2[0])
    else:
        f2 = candidates(fallback_pool)
        if not f2: return [], True
        f2.sort(key=lambda r: (loads.get(r,0), random.random()))
        selected.append(f2[0]); used_fallback = True
    return selected, used_fallback

def repair_min_load(assign_by_app, eligible_by_app, loads, min_load, max_load, stream_req_by_app, stream_ok):
    apps_by_rev = defaultdict(set)
    for app_id, revs in assign_by_app.items():
        for r in revs: apps_by_rev[r].add(app_id)
    changed = True; iters = 0
    while changed and iters < 2000:
        iters += 1; changed = False
        under = [r for r,c in loads.items() if c < min_load]
        over  = [r for r,c in loads.items() if c > min_load]
        if not under or not over: break
        u = min(under, key=lambda r: loads[r])
        for d in sorted(over, key=lambda r: loads[r], reverse=True):
            for app_id in list(apps_by_rev[d]):
                if u in assign_by_app[app_id]: continue
                if u not in eligible_by_app[app_id]: continue
                req = stream_req_by_app.get(app_id, "UNKNOWN")
                if req in ("CIHR","SSHRC"):
                    cur_ok = [stream_ok[req].get(rid,False) for rid in assign_by_app[app_id]]
                    u_ok = stream_ok[req].get(u,False)
                    if all(cur_ok) and not u_ok:  # avoid breaking a perfect stream match
                        continue
                # swap
                assign_by_app[app_id].remove(d); assign_by_app[app_id].append(u)
                apps_by_rev[d].remove(app_id); apps_by_rev[u].add(app_id)
                loads[d] -= 1; loads[u] += 1
                changed = True; break
            if changed: break
    return assign_by_app, loads

# ----------------------------
# Main
# ----------------------------
def main():
    # Read Stage 1 eligibility
    try:
        elig = pd.read_excel(ELIG_FILE, sheet_name=ELIG_SHEET, engine="openpyxl")
    except Exception as e:
        print(f"ERROR reading {ELIG_FILE}: {e}"); sys.exit(1)
    for col in (COL_APP_ID, COL_REV_NSID):
        if col not in elig.columns:
            print(f"ERROR: {ELIG_SHEET} missing column {col}"); sys.exit(1)
    elig[COL_APP_ID] = elig[COL_APP_ID].astype(int)
    elig[COL_REV_NSID] = elig[COL_REV_NSID].apply(norm_lower)

    # Read Applicants (has AppID, StudentName, Which award...)
    try:
        apps = pd.read_excel(APPS_FILE, engine="openpyxl")
    except Exception as e:
        print(f"ERROR reading {APPS_FILE}: {e}"); sys.exit(1)
    for col in (COL_APP_ID, COL_STUDENT_NAME, COL_AWARD):
        if col not in apps.columns:
            print(f"ERROR: Applicants missing column {col}"); sys.exit(1)
    apps[COL_APP_ID] = apps[COL_APP_ID].astype(int)
    apps["AwardStream"] = apps[COL_AWARD].apply(parse_award_stream)

    # Read Reviewers (has ReviewerNSID, ReviewerName, StreamReview Eligibility)
    try:
        revs = pd.read_excel(REVS_FILE, sheet_name=REVS_SHEET, engine="openpyxl")
    except Exception as e:
        print(f"ERROR reading {REVS_FILE} ({REVS_SHEET}): {e}"); sys.exit(1)
    for col in (COL_REV_NSID, COL_REV_NAME, COL_STREAMS):
        if col not in revs.columns:
            print(f"ERROR: Reviewers missing column {col}"); sys.exit(1)
    revs[COL_REV_NSID] = revs[COL_REV_NSID].apply(norm_lower)
    revs["StreamSet"] = revs[COL_STREAMS].apply(split_semicolon_upper_tokens)

    # Lookups for names
    app_name = dict(zip(apps[COL_APP_ID], apps[COL_STUDENT_NAME]))       # AppID -> StudentName
    rev_name = dict(zip(revs[COL_REV_NSID], revs[COL_REV_NAME]))         # ReviewerNSID -> ReviewerName

    # Stream OK matrices
    stream_ok = {
        "CIHR":  {rid: ("CIHR"  in s) for rid, s in zip(revs[COL_REV_NSID], revs["StreamSet"])},
        "SSHRC": {rid: ("SSHRC" in s) for rid, s in zip(revs[COL_REV_NSID], revs["StreamSet"])},
        "NSERC": {rid: ("NSERC" in s) for rid, s in zip(revs[COL_REV_NSID], revs["StreamSet"])},
    }

    # Eligible reviewers per app (from Stage 1)
    eligible_by_app = defaultdict(list)
    for app_id, rid in zip(elig[COL_APP_ID], elig[COL_REV_NSID]):
        eligible_by_app[int(app_id)].append(rid)

    # Stream requirement per app
    stream_req_by_app = dict(zip(apps[COL_APP_ID], apps["AwardStream"]))

    # Pre-check CIHR/SSHRC pools
    exceptions = []
    for app_id, req in stream_req_by_app.items():
        if req in ("CIHR","SSHRC"):
            pool = eligible_by_app.get(app_id, [])
            preferred = [r for r in pool if stream_ok[req].get(r, False)]
            if len(preferred) < 2:
                exceptions.append({
                    "AppID": app_id,
                    "Issue": "Insufficient stream-eligible reviewers (will require fallback)",
                    "AwardStream": req,
                    "PreferredEligibleCount": len(preferred),
                    "TotalEligibleCount": len(pool)
                })

    # Initialize loads
    reviewers = list(revs[COL_REV_NSID].unique())
    loads = {rid: 0 for rid in reviewers}

    # Sort apps (hardest first): CIHR/SSHRC by preferred-count then total-count; NSERC later
    app_ids = list(apps[COL_APP_ID].unique())
    def app_key(app_id):
        pool = eligible_by_app.get(app_id, [])
        req = stream_req_by_app.get(app_id, "UNKNOWN")
        if req in ("CIHR","SSHRC"):
            preferred = [r for r in pool if stream_ok[req].get(r, False)]
            return (len(preferred), len(pool))
        return (9999, len(pool))
    app_ids.sort(key=app_key)

    # Assign
    assignments_by_app = {}
    for app_id in app_ids:
        pool = list(dict.fromkeys(eligible_by_app.get(app_id, [])))
        if len(pool) < REVIEWS_PER_APP:
            exceptions.append({"AppID": app_id, "Issue": "Fewer than 2 COI-eligible reviewers", "AwardStream": stream_req_by_app.get(app_id,"UNKNOWN")})
            assignments_by_app[app_id] = []
            continue

        req = stream_req_by_app.get(app_id, "UNKNOWN")
        if req in ("CIHR","SSHRC"):
            preferred_pool = [r for r in pool if stream_ok[req].get(r, False)]
            chosen, used_fallback = pick_two_reviewers(preferred_pool, pool, loads, MAX_LOAD)
            if not chosen or len(chosen) < 2:
                exceptions.append({"AppID": app_id, "Issue": "Could not assign 2 under max load", "AwardStream": req})
                assignments_by_app[app_id] = chosen or []
            else:
                assignments_by_app[app_id] = chosen
                if used_fallback:
                    exceptions.append({"AppID": app_id, "Issue": "Fallback used for CIHR/SSHRC", "AwardStream": req})
        else:
            chosen, _ = pick_two_reviewers(pool, pool, loads, MAX_LOAD)
            if not chosen or len(chosen) < 2:
                exceptions.append({"AppID": app_id, "Issue": "Could not assign 2 under max load", "AwardStream": req})
                assignments_by_app[app_id] = chosen or []
            else:
                assignments_by_app[app_id] = chosen

        for rid in assignments_by_app[app_id]:
            loads[rid] = loads.get(rid,0) + 1

    # Repair to reach MIN_LOAD where possible without breaking stream matches
    assignments_by_app, loads = repair_min_load(assignments_by_app, eligible_by_app, loads, MIN_LOAD, MAX_LOAD, stream_req_by_app, stream_ok)

    # Build outputs with names
    rows_app = []
    for app_id in sorted(assignments_by_app.keys()):
        assigned = assignments_by_app[app_id]
        r1 = assigned[0] if len(assigned) > 0 else ""
        r2 = assigned[1] if len(assigned) > 1 else ""
        rows_app.append({
            "AppID": app_id,
            "StudentName": app_name.get(app_id, ""),                         # <-- added
            "AwardStream": stream_req_by_app.get(app_id, "UNKNOWN"),
            "Reviewer1NSID": r1,
            "Reviewer1Name": rev_name.get(r1, ""),                           # <-- added
            "Reviewer2NSID": r2,
            "Reviewer2Name": rev_name.get(r2, ""),                           # <-- added
        })
    by_app_df = pd.DataFrame(rows_app)

    apps_by_rev = defaultdict(list)
    for app_id, revs_assigned in assignments_by_app.items():
        for rid in revs_assigned:
            apps_by_rev[rid].append(app_id)

    rows_rev = []
    for rid, count in sorted(loads.items(), key=lambda kv: kv[0]):
        rows_rev.append({
            "ReviewerNSID": rid,
            "ReviewerName": rev_name.get(rid, ""),                           # <-- added
            "AssignedAppCount": count,
            "AssignedAppIDs": "; ".join(str(x) for x in sorted(apps_by_rev.get(rid, []))),
        })
    by_rev_df = pd.DataFrame(rows_rev)

    load_summary = by_rev_df[["ReviewerNSID","ReviewerName","AssignedAppCount"]].copy()
    load_summary["Within10to12"] = load_summary["AssignedAppCount"].between(MIN_LOAD, MAX_LOAD)

    total_assigned = sum(loads.values())
    expected = len(app_ids) * REVIEWS_PER_APP
    exceptions_df = pd.DataFrame([{
        "AppID": None,
        "Issue": f"Total assigned reviews ({total_assigned}) != expected ({expected})",
        "AwardStream": None
    }]) if total_assigned != expected else pd.DataFrame(columns=["AppID","Issue","AwardStream"])

    # Append any previously collected exceptions
    if exceptions:
        exceptions_df = pd.concat([exceptions_df, pd.DataFrame(exceptions)], ignore_index=True)

    # Flag any out-of-range loads
    for _, row in load_summary[~load_summary["Within10to12"]].iterrows():
        exceptions_df = pd.concat([exceptions_df, pd.DataFrame([{
            "AppID": None, "Issue": f"Reviewer load out of range: {row['ReviewerNSID']} has {row['AssignedAppCount']}",
            "AwardStream": None
        }])], ignore_index=True)

    with pd.ExcelWriter(OUT_FILE, engine="openpyxl") as w:
        by_app_df.to_excel(w, index=False, sheet_name="Assignments_ByApp")
        by_rev_df.to_excel(w, index=False, sheet_name="Assignments_ByReviewer")
        load_summary.to_excel(w, index=False, sheet_name="LoadSummary")
        exceptions_df.drop_duplicates().to_excel(w, index=False, sheet_name="Exceptions")

    print("Stage 2 complete.")
    print(f"Apps: {len(app_ids)} | Reviewers: {len(loads)}")
    print(f"Total reviews assigned: {total_assigned} (expected {expected})")
    print(f"Within 10-12: {int(load_summary['Within10to12'].sum())}/{len(load_summary)}")
    print(f"Output: {OUT_FILE}")

if __name__ == "__main__":
    main()