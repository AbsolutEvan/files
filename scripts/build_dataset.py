#!/usr/bin/env python3
"""
Build the unified schools.json for the A-F Sankey from the raw AZ files.

Joins three sources on the school's state entity id, scoped to FY2022-FY2025
high schools (those with Grade 11 ELA/Math assessment results):

  - A-F letter grades   -> the grade buckets per year   (Combined A-F files)
  - Grade 11 pass rates -> segment colors (ELA/Math)     (assessment_g11_slim.csv)
  - Enrollment totals   -> ribbon width                  (Oct 1 enrollment files)

Output: data/processed/schools.json  (one object per school, with a per-year series)
"""

import os
import re
import glob
import json
import pandas as pd

RAW = os.path.join(os.path.dirname(__file__), "..", "data", "raw", "tmp")
OUT = os.path.join(os.path.dirname(__file__), "..", "data", "processed", "schools.json")
YEARS = [2022, 2023, 2024, 2025]
VALID_GRADES = {"A", "B", "C", "D", "F"}


def sid(x):
    """Normalize an entity id to a clean string ('4737.0' -> '4737')."""
    if pd.isna(x):
        return None
    s = str(x).strip()
    return s[:-2] if s.endswith(".0") else s


def normcols(df):
    """Lowercase + strip non-alphanumerics so 'School Code'=='SchoolCode'."""
    df.columns = [re.sub(r"[^a-z0-9]", "", str(c).lower()) for c in df.columns]
    return df


def load_af():
    rows = []
    for f in glob.glob(os.path.join(RAW, "*Combined*A-F*")) + glob.glob(os.path.join(RAW, "FY_2022*")):
        for sheet in pd.ExcelFile(f).sheet_names:
            if "School" not in sheet:        # model sheets only (skip Please Read / LEA)
                continue
            d = normcols(pd.read_excel(f, sheet_name=sheet, dtype=str))
            if "schoolcode" not in d.columns or "lettergrade" not in d.columns:
                continue
            d["school_id"] = d["schoolcode"].map(sid)
            d["year"] = pd.to_numeric(d["fiscalyear"], errors="coerce")
            d["charter"] = d.get("charter", "").map(
                lambda v: True if str(v).strip().lower() in ("y", "yes") else False)
            keep = d[["year", "school_id", "schoolname", "lettergrade",
                      "districtname", "county", "charter", "model"]].copy()
            rows.append(keep)
    af = pd.concat(rows, ignore_index=True)
    af = af[af.school_id.notna() & af.year.notna()]
    af["year"] = af["year"].astype(int)
    af["grade"] = af["lettergrade"].str.strip().str.upper()
    return af.drop_duplicates(["year", "school_id"])


def load_assessment():
    a = pd.read_csv(os.path.join(RAW, "assessment_g11_slim.csv"), dtype=str)
    a["school_id"] = a["School Entity ID"].map(sid)
    a["year"] = pd.to_numeric(a["FiscalYear"], errors="coerce").astype("Int64")
    a["pp"] = pd.to_numeric(a["Percent Passing"], errors="coerce")
    out = {}
    for _, r in a.iterrows():
        key = (int(r["year"]), r["school_id"])
        col = "ela_pass" if r["subject_norm"] == "ELA" else "math_pass"
        out.setdefault(key, {})[col] = None if pd.isna(r["pp"]) else float(r["pp"])
    return out  # {(year, school_id): {ela_pass, math_pass}}


def load_enrollment():
    en = []
    for f in glob.glob(os.path.join(RAW, "*Enroll*")):
        d = normcols(pd.read_excel(f, sheet_name="School by Gender", dtype=str))
        d["school_id"] = d["schoolentityid"].map(sid)
        d["year"] = pd.to_numeric(d["fiscalyear"], errors="coerce")
        d["enrollment"] = pd.to_numeric(d["total"], errors="coerce")
        en.append(d[["year", "school_id", "enrollment"]].dropna(subset=["year"]))
    en = pd.concat(en, ignore_index=True)
    en = en[en.school_id.notna()]
    en["year"] = en["year"].astype(int)
    en = en.drop_duplicates(["year", "school_id"])
    return {(r.year, r.school_id): (None if pd.isna(r.enrollment) else int(r.enrollment))
            for r in en.itertuples()}


def main():
    af = load_af()
    assess = load_assessment()
    enroll = load_enrollment()

    # Universe = high schools that have Grade 11 assessment in any year.
    universe = sorted({k[1] for k in assess})

    # Metadata: prefer most recent A-F record per school.
    meta = (af.sort_values("year").groupby("school_id")
            .agg(name=("schoolname", "last"), district=("districtname", "last"),
                 county=("county", "last"), charter=("charter", "last"),
                 model=("model", "last")))
    grade_by = {(r.year, r.school_id): r.grade for r in af.itertuples()}

    schools = []
    for sidv in universe:
        series = []
        for y in YEARS:
            g = grade_by.get((y, sidv))
            g = g if g in VALID_GRADES else None
            ass = assess.get((y, sidv), {})
            series.append({
                "year": y,
                "grade": g,
                "ela_pass": ass.get("ela_pass"),
                "math_pass": ass.get("math_pass"),
                "enrollment": enroll.get((y, sidv)),
            })
        if not any(p["grade"] for p in series):
            continue  # no letter grade in any year -> can't place in a bucket
        m = meta.loc[sidv] if sidv in meta.index else None
        schools.append({
            "school_id": sidv,
            "name": (m["name"] if m is not None and pd.notna(m["name"]) else None),
            "district": (m["district"] if m is not None else None),
            "county": (m["county"] if m is not None else None),
            "charter": bool(m["charter"]) if m is not None else None,
            "model": (m["model"].strip() if m is not None and pd.notna(m["model"]) else None),
            "years": series,
        })

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    json.dump(schools, open(OUT, "w"), indent=1)
    # Inline copy the viz loads directly (avoids needing a local web server).
    datajs = os.path.join(os.path.dirname(__file__), "..", "src", "data.js")
    with open(datajs, "w") as fh:
        fh.write("window.SCHOOLS = " + json.dumps(schools, separators=(",", ":")) + ";")

    # ---- diagnostics ----
    print(f"schools written: {len(schools)}  -> {OUT}")
    from collections import Counter
    for y in YEARS:
        gc = Counter(p["grade"] for s in schools for p in s["years"]
                     if p["year"] == y and p["grade"])
        ela = sum(1 for s in schools for p in s["years"]
                  if p["year"] == y and p["ela_pass"] is not None)
        enr = sum(1 for s in schools for p in s["years"]
                  if p["year"] == y and p["enrollment"] is not None)
        graded = sum(gc.values())
        print(f"  {y}: graded={graded:>4}  grades={dict(sorted(gc.items()))}  "
              f"ela={ela}  enrollment={enr}")
    no_grade_yr = sum(1 for s in schools for p in s["years"] if not p["grade"])
    print(f"  school-years missing a grade (gaps): {no_grade_yr}")


if __name__ == "__main__":
    main()
