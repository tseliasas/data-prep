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


def balanced_hotdeck(series, groups=None, min_donors=20, seed=42):
    """Fill missing values by drawing from the observed distribution.

    When half the values are missing, filling them all with the median piles
    everything onto one point and shrinks the std. Instead, this picks as many
    evenly spaced quantiles as there are missing cells ('nearest' = only values
    that were actually observed) and shuffles them with a fixed seed so the
    result is reproducible. If a group has fewer than 20 donors, the whole
    column is used.
    """
    original = pd.to_numeric(series, errors="coerce").astype(float)
    result = original.copy()
    labels = pd.Series("All", index=series.index) if groups is None else groups.fillna("Unknown")
    rng = np.random.default_rng(seed)
    for group in pd.unique(labels):
        missing = labels.eq(group) & original.isna()
        if not missing.any():
            continue
        donors = original[labels.eq(group)].dropna()
        if len(donors) < min_donors:
            donors = original.dropna()
        positions = (np.arange(missing.sum()) + .5) / missing.sum()
        values = np.quantile(donors.to_numpy(), positions, method="nearest")
        result[missing] = rng.permutation(values)
    return result



# Employees
def clean_employees(path):
    sheets = pd.read_excel(path, sheet_name=None, dtype=object, keep_default_na=False)
    main = [v for v in sheets.values() if list(v.columns) == EMP_COLS]
    packed = [v for v in sheets.values()
              if len(v.columns) == 1 and str(v.columns[0]).split(",") == EMP_COLS]
    assert len(main) == 1 and len(packed) == 1, "Expected two sheets were not found."
    raw = main[0]
    d = raw.copy()

    # Scale error: on the first sheet ages run 200-600 and salaries lost their decimal point.
    # The second sheet holds the same records as CSV text, with the decimal points intact.
    # The fix is applied only if removing the point makes every value match the first sheet exactly.
    ref = pd.DataFrame([next(csv.reader([str(v)])) for v in packed[0].iloc[:, 0]], columns=EMP_COLS)
    for c in ["Name", "Gender", "City", "Department", "Email"]:
        assert d[c].astype(str).equals(ref[c].astype(str)), "Sheets are not in the same order."
    for c in ["Age", "Salary"]:
        actual = pd.to_numeric(d[c], errors="coerce")
        encoded = pd.to_numeric(ref[c].str.replace(".", "", regex=False), errors="coerce")
        assert (actual.eq(encoded) | (actual.isna() & encoded.isna())).all(), f"{c}: scale fix could not be verified."
        d[c] = pd.to_numeric(ref[c], errors="coerce")

    # Text standardization: whitespace, letter case, typos.
    for c in ["Name", "Gender", "City", "Department", "Email"]:
        d[c] = d[c].map(norm_text)
    d["City"] = d.City.str.title()  # lahore -> Lahore
    departments = {"hr": "HR", "it": "IT", "sales": "Sales", "finance": "Finance", "marketing": "Marketing"}
    d["Department"] = d.Department.map(lambda x: departments.get(str(x).casefold(), x))
    d["Email"] = d.Email.str.lower().str.replace(r"@gamil\.com$", "@gmail.com", regex=True)
    d["Email"] = d.Email.mask(~d.Email.str.fullmatch(r"[^\s@]+@[^\s@]+\.[^\s@]+", na=False), pd.NA)
    # Validity: impossible ages and salaries count as missing.
    d["Age"] = d.Age.mask(~d.Age.between(18, 120), np.nan)
    d["Salary"] = d.Salary.mask(d.Salary.le(0), np.nan)

    baseline = d.copy()  # Scale fixed, not yet imputed (used for before/after comparison).
    flags = pd.DataFrame({"source_row": d.index + 2,
                          "Age_imputed": d.Age.isna().astype(int),
                          "Salary_imputed": d.Salary.isna().astype(int),
                          "Email_repeated": d.Email.duplicated(keep=False).astype(int)})

    # Justification: how much a median fill would shrink the std.
    median_std = {"Age": d.Age.fillna(d.Age.median()).std(),
                  "Salary": d.Salary.fillna(d.Salary.median()).std()}

    # Imputation: age from all employees, salary from donors in the same department.
    d["Age"] = balanced_hotdeck(d.Age).astype("Int64")
    d["Salary"] = balanced_hotdeck(d.Salary, d.Department).round(2)
    for c in ["Name", "Gender", "City", "Department", "Email"]:
        d[c] = d[c].fillna(UNSPECIFIED)

    # Checks
    assert len(d) == len(raw) and not d.isna().any().any()
    assert d.Age.between(18, 120).all() and d.Salary.gt(0).all()
    assert d.City.nunique() == 5 and d.Department.nunique() == 5
    assert d.loc[baseline.Age.notna(), "Age"].eq(baseline.Age.dropna()).all(), "Known ages changed."
    assert np.allclose(d.loc[baseline.Salary.notna(), "Salary"], baseline.Salary.dropna()), "Known salaries changed."
    return {"raw": raw, "baseline": baseline, "clean": d, "flags": flags, "median_std": median_std}


def main():
    pd.set_option("display.width", 160)
    pd.set_option("display.max_columns", 20)

    emp = clean_employees(INPUT / "employees.xlsx")

    print("=== 1-2. Exploration and missing values (raw data) ===")
    print("\nEmployees:\n", profile(emp["raw"]))

    print("\n=== 3. Why not a median fill ===")
    for c, std in emp["median_std"].items():
        observed = emp["baseline"][c].std()
        print(f"{c}: observed std {observed:,.2f} | median fill {std:,.2f} "
              f"({(1 - std / observed) * 100:.1f}% lower) | donor fill {emp['clean'][c].std():,.2f}")
    print(f"\nEmployees: {len(emp['raw'])} -> {len(emp['clean'])} rows")


if __name__ == "__main__":
    main()
