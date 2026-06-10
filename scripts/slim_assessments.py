#!/usr/bin/env python3
"""
Slim the big Arizona assessment workbooks down to just what the Sankey needs.

The raw assessment .xlsx files are 19-72 MB because they're long format:
one row per school x subject x test level x subgroup x FAY status. We only
want Grade 11, All Students, ELA + Math, and the "Percent Passing" value.
This collapses each file to ~one row per high school per subject.

USAGE
-----
    pip install pandas openpyxl
    python slim_assessments.py "/path/to/Assessment Data"

It scans every .xlsx in that folder, auto-detects the school-level sheet,
prints the distinct category values it finds (so we can confirm the filter
is matching the real labels), and writes:

    assessment_g11_slim.csv   (next to this script, a few hundred KB)

Upload that one small CSV to Drive (or commit it) and we're unblocked.
"""

import sys
import glob
import os
import pandas as pd

# Columns we expect in the school-level sheet (from the header you sent).
KEY_COLS = ["School Entity ID", "Percent Passing", "Subject", "Test Level", "Subgroup"]


def find_school_sheet(path):
    """Return the name of the sheet that has the school-level columns."""
    xls = pd.ExcelFile(path, engine="openpyxl")
    for sheet in xls.sheet_names:
        # Peek at just the header row of each sheet.
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
    if not files:
        sys.exit(f"No .xlsx files found in: {folder}")

    out_frames = []
    for path in files:
        name = os.path.basename(path)
        sheet = find_school_sheet(path)
        if sheet is None:
            print(f"  !! {name}: no school-level sheet found, skipping")
            continue
        print(f"\n== {name}  (sheet: {sheet}) ==")
        df = pd.read_excel(path, sheet_name=sheet, engine="openpyxl", dtype=str)
        df.columns = [str(c).strip() for c in df.columns]

        # Show the real category labels so we can trust the filter.
        for col in ["Subject", "Test Level", "Subgroup", "FAY Status"]:
            if col in df.columns:
                vals = sorted(v for v in df[col].dropna().unique())
                print(f"   {col}: {vals}")

        # Filter: Grade 11, All Students, ELA or Math.
        subj = df["Subject"].map(norm)
        level = df["Test Level"].map(norm)
        grp = df["Subgroup"].map(norm)

        is_ela = subj.str.contains("ela") | subj.str.contains("english")
        is_math = subj.str.contains("math")
        keep = (
            (is_ela | is_math)
            & level.str.contains("11")
            & grp.str.contains("all")
        )
        slim = df[keep].copy()
        slim["subject_norm"] = slim["Subject"].map(
            lambda s: "ELA" if ("ela" in norm(s) or "english" in norm(s)) else "Math"
        )
        print(f"   -> kept {len(slim)} rows")
        out_frames.append(slim)

    if not out_frames:
        sys.exit("Nothing matched the filter. Check the printed category values above.")

    result = pd.concat(out_frames, ignore_index=True)
    cols = [
        "FiscalYear", "School Entity ID", "SchoolName", "subject_norm",
        "Test Level", "Subgroup", "Number Tested", "Percent Passing",
    ]
    if "FAY Status" in result.columns:
        cols.insert(6, "FAY Status")
    result = result[[c for c in cols if c in result.columns]]

    out_path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                            "assessment_g11_slim.csv")
    result.to_csv(out_path, index=False)
    print(f"\nWROTE {out_path}  ({len(result)} rows)")


if __name__ == "__main__":
    main()
