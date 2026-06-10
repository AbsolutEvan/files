#!/usr/bin/env python3
"""
Slim the big Arizona assessment workbooks down to just what the Sankey needs.

We want, per high school per year:
  - Subject:     English Language Arts  OR  Mathematics
  - Test Level:  "ELA Grade 11" / "Math Grade 11"  (NOT the "Alt ..." versions)
  - Subgroup:    All Students
  - FAY Status:  All        (all students, not the FAY / Not FAY splits)
  - value:       Percent Passing
This collapses each 19-72 MB file to one row per school per subject.

USAGE
-----
    python slim_assessments_v2.py "/path/to/Assessment Data"

Drag the folder into Terminal after the script name to fill the path.
Writes  assessment_g11_slim.csv  next to the script.
"""

import sys
import glob
import os
import pandas as pd


def find_school_sheet(path):
    """Return the name of the sheet that has the school-level columns."""
    xls = pd.ExcelFile(path, engine="openpyxl")
    for sheet in xls.sheet_names:
        head = pd.read_excel(xls, sheet_name=sheet, nrows=0)
        cols = set(str(c).strip() for c in head.columns)
        if {"School Entity ID", "Percent Passing"}.issubset(cols):
            return sheet
    return None


def norm(s):
    return str(s).strip().lower()


def main():
    folder = sys.argv[1] if len(sys.argv) > 1 else "."
    files = sorted(glob.glob(os.path.join(folder, "*.xlsx")))
    # Skip Excel lock/temp files (~$foo.xlsx) and hidden files.
    files = [f for f in files
             if not os.path.basename(f).startswith(("~$", "."))]
    if not files:
        sys.exit(f"No .xlsx files found in: {folder}")

    out_frames = []
    for path in files:
        name = os.path.basename(path)
        try:
            sheet = find_school_sheet(path)
        except Exception as e:
            print(f"  !! {name}: could not open ({e}), skipping")
            continue
        if sheet is None:
            print(f"  !! {name}: no school-level sheet found, skipping")
            continue
        print(f"\n== {name}  (sheet: {sheet}) ==")
        df = pd.read_excel(path, sheet_name=sheet, engine="openpyxl", dtype=str)
        df.columns = [str(c).strip() for c in df.columns]

        level = df["Test Level"].map(norm)
        grp = df["Subgroup"].map(norm)
        keep = level.isin(["ela grade 11", "math grade 11"]) & grp.eq("all students")
        if "FAY Status" in df.columns:
            keep &= df["FAY Status"].map(norm).eq("all")

        slim = df[keep].copy()
        slim["subject_norm"] = slim["Test Level"].map(
            lambda s: "ELA" if "ela" in norm(s) else "Math"
        )
        # Safety: one row per school per subject.
        slim = slim.drop_duplicates(subset=["School Entity ID", "subject_norm"])
        print(f"   -> kept {len(slim)} rows "
              f"({(slim['subject_norm']=='ELA').sum()} ELA, "
              f"{(slim['subject_norm']=='Math').sum()} Math)")
        out_frames.append(slim)

    if not out_frames:
        sys.exit("Nothing matched the filter. Check the column labels.")

    result = pd.concat(out_frames, ignore_index=True)
    cols = [
        "FiscalYear", "School Entity ID", "SchoolName", "subject_norm",
        "Number Tested", "Percent Passing",
    ]
    result = result[[c for c in cols if c in result.columns]]

    out_path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                            "assessment_g11_slim.csv")
    result.to_csv(out_path, index=False)
    print(f"\nWROTE {out_path}  ({len(result)} rows total)")


if __name__ == "__main__":
    main()
