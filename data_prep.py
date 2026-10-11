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


def same_company_values(df, column):
    """If the same company's other postings have one consistent value, return it."""
    valid = df.dropna(subset=[column, "Company Name"])
    groups = valid.groupby("Company Name")[column]
    lookup = groups.first()[groups.nunique().eq(1)]
    return df["Company Name"].map(lookup)


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


def canonical_place(value):
    """'OSAKA, Japan' -> 'Osaka, Japan'; US state codes in upper case."""
    if value is pd.NA or pd.isna(value) or value == "United States":
        return value
    pieces = str(value).rsplit(",", 1)
    if len(pieces) != 2:
        return str(value)
    city, region = pieces[0].strip().title(), pieces[1].strip()
    city = re.sub(r"\bMc ?[Ll]ean\b", "McLean", city)
    return f"{city}, {region.upper() if re.fullmatch('[A-Za-z]{2}', region) else region.title()}"


def normalize_title(value):
    """Standardize job titles: Sr -> Senior, Jr -> Junior, acronyms in upper case."""
    if value is pd.NA or pd.isna(value):
        return pd.NA
    text = str(value)
    if text.isupper():
        text = text.title()
    text = re.sub(r"\bsr\b\.?", "Senior", text, flags=re.I)
    text = re.sub(r"\bjr\b\.?", "Junior", text, flags=re.I)
    for token in ["AI", "ML", "NLP", "BI", "SQL", "AWS", "TS", "SCI", "NGS", "IT", "USA", "RFP", "ELISA"]:
        text = re.sub(r"\b" + token + r"\b", token, text, flags=re.I)
    return text



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


# Job postings
def clean_jobs(path):
    # Separator is ';'. Multi-line descriptions inside quotes are read as one record.
    raw = pd.read_csv(path, sep=";", encoding="utf-8-sig", dtype=object, keep_default_na=False)
    assert len(raw.columns) == 27, "Check the CSV separator."
    d = raw.loc[~raw.duplicated()].copy()  # Drop the 12 exact duplicates (first copy kept).

    # Text: whitespace and hidden missing values (-1, Unknown, na ...) -> NaN.
    for c in d:
        d[c] = d[c].map(lambda v: norm_text(v, c == "Job Description"))
    d["Job Title"] = d["Job Title"].map(normalize_title)
    for c in ["Location", "Headquarters"]:
        d[c] = d[c].map(canonical_place)

    # Broken headquarters codes like 'New York, 061' are fixed from the city-state pairs in Location.
    states = d.Location.str.extract(r", ([A-Z]{2})$", expand=False)
    cities = d.Location.str.replace(r", [A-Z]{2}$", "", regex=True)
    pairs = pd.DataFrame({"city": cities, "state": states}).dropna().groupby("city").state
    city_to_state = pairs.first()[pairs.nunique().eq(1)]
    broken = d.Headquarters.fillna("").str.contains(r", \d+$")
    for idx in d.index[broken]:
        city = str(d.at[idx, "Headquarters"]).rsplit(",", 1)[0]
        state = city_to_state.get(city)
        d.at[idx, "Headquarters"] = f"{city}, {state}" if state is not None else pd.NA

    # Numeric types plus scale and validity fixes.
    for c in JOB_NUMERIC + ["same_state"] + SKILLS:
        d[c] = pd.to_numeric(d[c], errors="coerce")
    positive = d.Rating[d.Rating.gt(0)]
    if positive.between(10, 50).all():          # 38 -> 3.8 (lost decimal point)
        d["Rating"] = d.Rating / 10
    d["Rating"] = d.Rating.mask(~d.Rating.between(1, 5), np.nan)   # 0 = no rating
    d["company_age"] = d.company_age.mask(d.company_age.lt(0), np.nan)
    for c in ["min_salary", "max_salary"]:
        d[c] = d[c].mask(d[c].le(0), np.nan)
    assert not d.min_salary.gt(d.max_salary).any(), "Contradictory salary range."
    d["avg_salary"] = (d.min_salary + d.max_salary) / 2   # Recompute the wrong averages.
    text_range = d["Salary Estimate"].str.extract(r"^(\d+(?:\.\d+)?)-(\d+(?:\.\d+)?)$").astype(float)
    assert np.allclose(text_range[0], d.min_salary) and np.allclose(text_range[1], d.max_salary)
    d["job_state"] = states.fillna("US")

    baseline = d.copy()  # Repaired, not yet imputed.
    flags = pd.DataFrame({"source_row": d.index + 2}, index=d.index)
    for c in ["Rating", "company_age"]:
        flags[c + "_imputed"] = d[c].isna().astype(int)

    # Imputation 1: from the same company's consistent postings, and from unambiguous Industry -> Sector pairs.
    categories = ["Headquarters", "Size", "Type of ownership", "Industry", "Sector", "Revenue"]
    for c in categories + ["Rating", "company_age"]:
        d[c] = d[c].fillna(same_company_values(d, c))
    known = d.dropna(subset=["Industry", "Sector"]).groupby("Industry").Sector
    d["Sector"] = d.Sector.fillna(d.Industry.map(known.first()[known.nunique().eq(1)]))

    # Imputation 2: median for Rating; median of the company-size group for right-skewed company_age.
    d["Rating"] = d.Rating.fillna(baseline.Rating.median()).round(2)
    age_groups = baseline.groupby("Size").company_age.agg(["median", "count"])
    age_by_size = age_groups.loc[age_groups["count"].ge(10), "median"]
    d["company_age"] = d.company_age.fillna(d.Size.map(age_by_size).fillna(baseline.company_age.median()))

    # Imputation 3: categories get an explicit label instead of the most frequent value.
    for c in categories:
        d[c] = d[c].fillna(UNSPECIFIED)
    d["job_simp"] = d.job_simp.fillna("Not classified")
    d["seniority"] = d.seniority.fillna(UNSPECIFIED).replace({"senior": "Senior", "jr": "Junior"})

    # same_state: compare the job's state with the headquarters' state; a foreign HQ counts as different.
    hq_state = d.Headquarters.str.extract(r", ([A-Z]{2})$", expand=False)
    foreign_hq = d.Headquarters.ne(UNSPECIFIED) & hq_state.isna()
    comparable = states.notna() & hq_state.notna()
    same = pd.Series(UNSPECIFIED, index=d.index, dtype=object)
    same[comparable] = np.where(states[comparable].eq(hq_state[comparable]), "Same state", "Different state")
    same[states.notna() & foreign_hq] = "Different state"
    d["same_state"] = same
    for c in SKILLS:
        assert d[c].isin([0, 1]).all(), f"{c}: invalid skill flag."
        d[c] = d[c].astype("Int64")

    # Outliers (IQR): original values are kept, capped analysis columns are added.
    for c in ["avg_salary", "company_age"]:
        s = numeric_summary(baseline[c])
        flags[c + "_iqr_outlier"] = ((baseline[c] < s["lower"]) | (baseline[c] > s["upper"])).astype(int)
        d[c + "_iqr"] = d[c].clip(max(0.0, s["lower"]), s["upper"])
    flags["state_unspecified"] = d.same_state.eq(UNSPECIFIED).astype(int)

    # Checks
    assert not d.duplicated().any() and not d.isna().any().any()
    assert d.Rating.between(1, 5).all() and d.company_age.ge(0).all()
    assert np.allclose(d.avg_salary, (d.min_salary + d.max_salary) / 2)
    assert d.job_state.str.fullmatch(r"[A-Z]{2}").all()
    assert np.allclose(d.loc[baseline.company_age.notna(), "company_age"], baseline.company_age.dropna())
    return {"raw": raw, "baseline": baseline, "clean": d, "flags": flags}


# Summaries
def compare(obj, columns, use_iqr=False):
    """Step 5: mean and std before/after cleaning, plus the pre-imputation IQR fences."""
    rows = {}
    for c in columns:
        before = numeric_summary(obj["baseline"][c])
        col = c + "_iqr" if use_iqr and c + "_iqr" in obj["clean"] else c
        after = numeric_summary(obj["clean"][col])
        rows[col] = {"mean_before": before["mean"], "mean_after": after["mean"],
                     "std_before": before["std"], "std_after": after["std"],
                     "iqr_lower": before["lower"], "iqr_upper": before["upper"],
                     "outliers": before["outliers"]}
    return pd.DataFrame(rows).T.round(3)


def main():
    pd.set_option("display.width", 160)
    pd.set_option("display.max_columns", 20)

    emp = clean_employees(INPUT / "employees.xlsx")
    jobs = clean_jobs(INPUT / "job_postings.csv")

    print("=== 1-2. Exploration and missing values (raw data) ===")
    print("\nEmployees:\n", profile(emp["raw"]))
    print("\nJob postings:\n", profile(jobs["raw"], rating_zero=True))
    print(f"\nExact duplicate postings: {jobs['raw'].duplicated().sum()}")

    print("\n=== 3. Why not a median fill ===")
    for c, std in emp["median_std"].items():
        observed = emp["baseline"][c].std()
        print(f"{c}: observed std {observed:,.2f} | median fill {std:,.2f} "
              f"({(1 - std / observed) * 100:.1f}% lower) | donor fill {emp['clean'][c].std():,.2f}")

    print("\n=== 4-5. IQR fences and before/after comparison ===")
    print(compare(emp, ["Age", "Salary"]))
    print(compare(jobs, JOB_NUMERIC, use_iqr=True))

    print(f"\nEmployees: {len(emp['raw'])} -> {len(emp['clean'])} rows")
    print(f"Job postings: {len(jobs['raw'])} -> {len(jobs['clean'])} rows")


if __name__ == "__main__":
    main()
