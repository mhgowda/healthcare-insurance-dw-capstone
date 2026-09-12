"""
STEP 1 - Data Profiling
=======================
Purpose: Look at every CSV file and report what problems exist.

Outputs (matching the project document folder structure):
  logs/profiling_summary.json   main output, referenced in the document
  output/profiling_report.csv   same data, open in Excel to explore

What we check:
  - How many rows and columns
  - How many NULL / missing values per column
  - How many duplicate primary keys
  - How many rows with negative amounts
  - How many rows with future dates (should not exist)
  - How many rows with the ORPHAN marker (broken foreign keys)
"""

import sys
import json
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import pandas as pd
from config import CSV, OUTPUT_DIR, LOGS_DIR, TODAY

# -- Which column is the primary key for each entity ----------
PK_MAP = {
    "members":     "member_id",
    "providers":   "provider_id",
    "claims":      "claim_number",
    "claim_lines": "line_id",
    "payments":    "payment_id",
    "diagnoses":   "diag_id",
    "procedures":  "proc_id",
}

# -- Columns that must be numeric -----------------------------
NUMERIC_COLS = {
    "claims":      ["allowed_amount", "paid_amount", "length_of_stay"],
    "claim_lines": ["units", "rate", "line_amount"],
    "payments":    ["paid_amount"],
}

# -- Columns that must be valid dates -------------------------
DATE_COLS = {
    "members":     ["dob", "created_at"],
    "claims":      ["admit_date", "discharge_date"],
    "claim_lines": ["service_date"],
    "payments":    ["payment_date"],
}


def profile_one_file(entity_name, filepath):
    """Read one CSV and return a quality summary dict."""
    print(f"\n--- Profiling: {entity_name} ---")
    df = pd.read_csv(filepath, low_memory=False)
    print(f"  Rows: {len(df):,}   Columns: {len(df.columns)}")

    summary = {
        "entity":          entity_name,
        "total_rows":      len(df),
        "columns":         list(df.columns),
        "missing_values":  {},
        "duplicate_pk":    {},
        "orphan_fk":       {},
        "negative_values": {},
        "non_numeric":     {},
        "future_dates":    {},
    }

    # Missing values
    for col in df.columns:
        n_null  = int(df[col].isna().sum())
        n_empty = int((df[col].astype(str).str.strip() == "").sum()) \
                  if df[col].dtype == object else 0
        total = n_null + n_empty
        if total:
            summary["missing_values"][col] = total
            print(f"  Missing '{col}': {total:,}")

    # Duplicate primary keys
    pk = PK_MAP.get(entity_name)
    if pk and pk in df.columns:
        n_dup = int(df.duplicated(subset=[pk]).sum())
        summary["duplicate_pk"][pk] = n_dup
        if n_dup:
            print(f"  Duplicate {pk}: {n_dup:,}")

    # Orphan foreign keys
    for fk in [c for c in df.columns
               if c in ("member_id", "provider_id", "claim_number", "line_id")]:
        n_orphan = int((df[fk].astype(str) == "__ORPHAN__").sum())
        if n_orphan:
            summary["orphan_fk"][fk] = n_orphan
            print(f"  Orphan FK '{fk}': {n_orphan:,}")

    # Negative or non-numeric amounts
    for col in NUMERIC_COLS.get(entity_name, []):
        if col in df.columns:
            numeric = pd.to_numeric(df[col], errors="coerce")
            n_bad   = int(numeric.isna().sum() - df[col].isna().sum())
            n_neg   = int((numeric < 0).sum())
            if n_bad:
                summary["non_numeric"][col] = n_bad
                print(f"  Non-numeric '{col}': {n_bad:,}")
            if n_neg:
                summary["negative_values"][col] = n_neg
                print(f"  Negative '{col}': {n_neg:,}")

    # Future dates
    for col in DATE_COLS.get(entity_name, []):
        if col in df.columns:
            dt = pd.to_datetime(df[col], errors="coerce")
            n_future = int((dt > TODAY).sum())
            if n_future:
                summary["future_dates"][col] = n_future
                print(f"  Future date '{col}': {n_future:,}")

    return summary


def main():
    print("=" * 50)
    print("DATA PROFILING")
    print("=" * 50)

    all_summaries = []
    flat_rows = []   # for the CSV version

    for entity, path in CSV.items():
        if not path.exists():
            print(f"  [SKIP] File not found: {path}")
            continue
        summary = profile_one_file(entity, path)
        all_summaries.append(summary)

        # Flatten for CSV
        for check_type, findings in summary.items():
            if isinstance(findings, dict):
                for col, count in findings.items():
                    flat_rows.append({
                        "entity": entity,
                        "column": col,
                        "check":  check_type,
                        "count":  count,
                    })

    # -- Save logs/profiling_summary.json (as per project document) --
    json_path = LOGS_DIR / "profiling_summary.json"
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(all_summaries, f, indent=2, default=str)

    # -- Save output/profiling_report.csv (open in Excel) ------------
    csv_path = OUTPUT_DIR / "profiling_report.csv"
    pd.DataFrame(flat_rows).to_csv(csv_path, index=False)

    print("\n" + "=" * 50)
    print(f"Profiling done.")
    print(f"  JSON -> {json_path}")
    print(f"  CSV  -> {csv_path}")


if __name__ == "__main__":
    main()
