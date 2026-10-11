#!/usr/bin/env python3
"""Clean two messy datasets (employees.xlsx and job_postings.csv) and save them as Excel.

Usage:
    pip install -r requirements.txt
    python data_prep.py

Input : data/employees.xlsx, data/job_postings.csv (never modified)
Output: output/employees_clean.xlsx, output/job_postings_clean.xlsx
        (CleanData = cleaned records, Audit = imputation and outlier flags per row)
        output/charts/*.png (missing-value heatmaps, missing rates, box plots, histograms)

Steps:
    1. Explore the data (structural problems, stray whitespace, wrong types)
    2. Find hidden missing values and visualize them
    3. Fill missing values with a justified strategy
    4. Detect outliers with IQR and cap them
    5. Check that cleaning didn't distort the distributions (mean/std, histograms)
"""
import csv
import importlib.util
import re
import sys
import unicodedata
from pathlib import Path

# Show the install command instead of a long traceback if a library is missing.
_missing = [m for m in ["pandas", "numpy", "matplotlib", "openpyxl"] if importlib.util.find_spec(m) is None]
if _missing:
    sys.exit("Missing libraries: " + ", ".join(_missing)
             + "\nInstall with: python -m pip install " + " ".join(_missing))

import matplotlib
matplotlib.use("Agg")  # Save charts to files without opening windows.
import matplotlib.pyplot as plt
from matplotlib.colors import ListedColormap
import numpy as np
import pandas as pd

BASE = Path(__file__).resolve().parent
INPUT = BASE / "data"
OUTPUT = BASE / "output"

EMP_COLS = ["Name", "Gender", "Age", "City", "Department", "Salary", "Email"]
SKILLS = ["python", "excel", "hadoop", "spark", "aws", "tableau", "big_data"]
JOB_NUMERIC = ["Rating", "min_salary", "max_salary", "avg_salary", "company_age"]
# Placeholders that mean "missing". All of them become a real NaN.
MISSING_TOKENS = {"", "?", "null", "n/a", "na", "nan", "none", "-1", "unknown",
                  "unknown / non-applicable"}
UNSPECIFIED = "Not specified"
COLOR_BEFORE, COLOR_AFTER = "#4C6E91", "#C9774A"


# Helpers
def norm_text(value, multiline=False):
    """Fix whitespace and turn hidden missing placeholders into NaN.

    multiline=True keeps the paragraph structure of job descriptions.
    """
    if value is None or value is pd.NA or pd.isna(value):
        return pd.NA
    text = unicodedata.normalize("NFKC", str(value)).replace("​", "")
    if multiline:
        lines = [re.sub(r"[\t  ]+", " ", line).strip()
                 for line in text.replace("\r\n", "\n").replace("\r", "\n").split("\n")]
        text = re.sub(r"\n{3,}", "\n\n", "\n".join(lines)).strip()
    else:
        text = re.sub(r"\s+", " ", text).strip()
    return pd.NA if text.casefold() in MISSING_TOKENS else text


def missing_mask(df, rating_zero=False):
    """Cell-level missing table. For job postings, Rating = 0 also counts as missing
    (ratings run from 1 to 5; a 0 in the skill columns is a valid value)."""
    mask = pd.DataFrame({c: df[c].map(norm_text).isna() for c in df}, index=df.index)
    if rating_zero and "Rating" in mask:
        mask["Rating"] |= pd.to_numeric(df["Rating"], errors="coerce").eq(0)
    return mask


def profile(df, rating_zero=False):
    """Step 1: per column, count unique values, missing cells (empty + hidden) and whitespace problems."""
    miss = missing_mask(df, rating_zero)
    rows = []
    for c in df:
        text = df[c].fillna("").astype(str)
        empty = df[c].isna() | text.str.strip().eq("")
        rows.append({"column": c, "unique": df[c].nunique(),
                     # Values that only differ by case or whitespace, like 'lahore'/'Lahore' or ' HR '/'HR'
                     "case_space_variants": df[c].nunique() - text.str.strip().str.casefold().nunique(),
                     "empty": int(empty.sum()),
                     "hidden_missing": int((miss[c] & ~empty).sum()),
                     "edge_spaces": int(text.ne(text.str.strip()).sum()),
                     "repeated_spaces": int(text.str.contains(r"[\t ]{2,}").sum())})
    return pd.DataFrame(rows).set_index("column")


def numeric_summary(series):
    """Mean, std, quartiles, IQR fences and outlier count."""
    s = pd.to_numeric(series, errors="coerce").dropna().astype(float)
    q1, q3 = s.quantile([.25, .75])
    iqr = q3 - q1
    lower, upper = q1 - 1.5 * iqr, q3 + 1.5 * iqr
    return {"n": len(s), "mean": s.mean(), "median": s.median(), "std": s.std(),
            "skew": s.skew(), "lower": lower, "upper": upper,
            "outliers": int(((s < lower) | (s > upper)).sum())}
