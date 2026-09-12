"""
STEP 2 - Data Cleaning (Validation)
=====================================
Purpose: Read each CSV, identify bad rows, separate them out.
         Save the clean rows for the next step.
         Save the bad rows to output/rejected_rows/ so we can audit them.

What makes a row BAD
  - Required column is NULL/empty
  - Duplicate primary key
  - Orphan foreign key (value == "__ORPHAN__")
  - Negative amount where positive is required
  - Non-numeric value in a number column
  - Future date in an admit/service/payment date
  - Invalid ICD version (not "ICD-9" or "ICD-10")
"""

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import pandas as pd
from config import CSV, REJECTED_DIR, TODAY

TODAY_STR = str(TODAY.date())


# -- Helpers ---------------------------------------------------

def is_missing(series):
    """Returns boolean mask: True where value is null or empty string."""
    return series.isna() | (series.astype(str).str.strip() == "")

def is_future_date(series):
    """Returns boolean mask: True where date is in the future."""
    dt = pd.to_datetime(series, errors="coerce")
    return dt > TODAY

def is_numeric(series):
    """Returns boolean mask: True where value CAN be parsed as a number."""
    return pd.to_numeric(series, errors="coerce").notna()

def split_good_bad(df, bad_mask, reason_label):
    """Split a DataFrame into (good rows, bad rows).
       bad_mask: boolean Series where True = bad row."""
    good = df[~bad_mask].reset_index(drop=True)
    bad  = df[bad_mask].copy()
    bad["rejection_reason"] = reason_label
    bad = bad.reset_index(drop=True)
    return good, bad


def save_rejected(entity_name, rejected_df):
    if rejected_df.empty:
        return
    out = REJECTED_DIR / f"{entity_name}_rejected.csv"
    if out.exists():
        # Append to existing file (multiple checks may reject different rows)
        existing = pd.read_csv(out)
        combined = pd.concat([existing, rejected_df], ignore_index=True)
        combined.to_csv(out, index=False)
    else:
        rejected_df.to_csv(out, index=False)


# -- Per-entity cleaning ---------------------------------------

def clean_members(df):
    print(f"\n[CLEAN] members: {len(df):,} rows in")
    all_rejected = []

    # Rule 1: Required fields must not be null
    for col in ["member_id", "first_name", "last_name", "gender", "city", "state", "dob"]:
        bad = is_missing(df[col])
        df, rej = split_good_bad(df, bad, f"missing_{col}")
        all_rejected.append(rej)

    # Rule 2: No duplicate member_id (keep first)
    bad = df.duplicated(subset=["member_id"], keep="first")
    df, rej = split_good_bad(df, bad, "duplicate_member_id")
    all_rejected.append(rej)

    # Rule 3: DOB must not be a future date
    bad = is_future_date(df["dob"])
    df, rej = split_good_bad(df, bad, "future_dob")
    all_rejected.append(rej)

    rejected = pd.concat(all_rejected, ignore_index=True)
    save_rejected("members", rejected)
    print(f"  -> {len(df):,} valid rows  |  {len(rejected):,} rejected")
    return df


def clean_providers(df):
    print(f"\n[CLEAN] providers: {len(df):,} rows in")
    all_rejected = []

    for col in ["provider_id", "provider_name", "specialty", "city", "state", "tax_id", "license_no"]:
        bad = is_missing(df[col])
        df, rej = split_good_bad(df, bad, f"missing_{col}")
        all_rejected.append(rej)

    bad = df.duplicated(subset=["provider_id"], keep="first")
    df, rej = split_good_bad(df, bad, "duplicate_provider_id")
    all_rejected.append(rej)

    rejected = pd.concat(all_rejected, ignore_index=True)
    save_rejected("providers", rejected)
    print(f"  -> {len(df):,} valid rows  |  {len(rejected):,} rejected")
    return df


def clean_claims(df, valid_member_ids, valid_provider_ids):
    print(f"\n[CLEAN] claims: {len(df):,} rows in")
    all_rejected = []

    for col in ["claim_number", "member_id", "provider_id", "admit_date"]:
        bad = is_missing(df[col])
        df, rej = split_good_bad(df, bad, f"missing_{col}")
        all_rejected.append(rej)

    # Orphan FKs
    bad = df["member_id"].astype(str) == "__ORPHAN__"
    df, rej = split_good_bad(df, bad, "orphan_member_id")
    all_rejected.append(rej)

    bad = df["provider_id"].astype(str) == "__ORPHAN__"
    df, rej = split_good_bad(df, bad, "orphan_provider_id")
    all_rejected.append(rej)

    # FK must exist in parent tables
    bad = ~df["member_id"].isin(valid_member_ids)
    df, rej = split_good_bad(df, bad, "member_id_not_found")
    all_rejected.append(rej)

    bad = ~df["provider_id"].isin(valid_provider_ids)
    df, rej = split_good_bad(df, bad, "provider_id_not_found")
    all_rejected.append(rej)

    bad = df.duplicated(subset=["claim_number"], keep="first")
    df, rej = split_good_bad(df, bad, "duplicate_claim_number")
    all_rejected.append(rej)

    # Negative allowed_amount
    bad = pd.to_numeric(df["allowed_amount"], errors="coerce") < 0
    df, rej = split_good_bad(df, bad, "negative_allowed_amount")
    all_rejected.append(rej)

    # Non-numeric paid_amount
    bad = ~is_numeric(df["paid_amount"])
    df, rej = split_good_bad(df, bad, "non_numeric_paid_amount")
    all_rejected.append(rej)

    # Negative length_of_stay
    bad = pd.to_numeric(df["length_of_stay"], errors="coerce") < 0
    df, rej = split_good_bad(df, bad, "negative_length_of_stay")
    all_rejected.append(rej)

    # Future admit_date
    bad = is_future_date(df["admit_date"])
    df, rej = split_good_bad(df, bad, "future_admit_date")
    all_rejected.append(rej)

    rejected = pd.concat(all_rejected, ignore_index=True)
    save_rejected("claims", rejected)
    print(f"  -> {len(df):,} valid rows  |  {len(rejected):,} rejected")
    return df


def clean_claim_lines(df, valid_claim_numbers):
    print(f"\n[CLEAN] claim_lines: {len(df):,} rows in")
    all_rejected = []

    for col in ["line_id", "claim_number", "service_date", "procedure_code"]:
        bad = is_missing(df[col])
        df, rej = split_good_bad(df, bad, f"missing_{col}")
        all_rejected.append(rej)

    bad = df["claim_number"].astype(str) == "__ORPHAN__"
    df, rej = split_good_bad(df, bad, "orphan_claim_number")
    all_rejected.append(rej)

    bad = ~df["claim_number"].isin(valid_claim_numbers)
    df, rej = split_good_bad(df, bad, "claim_number_not_found")
    all_rejected.append(rej)

    bad = df.duplicated(subset=["line_id"], keep="first")
    df, rej = split_good_bad(df, bad, "duplicate_line_id")
    all_rejected.append(rej)

    bad = pd.to_numeric(df["units"], errors="coerce") <= 0
    df, rej = split_good_bad(df, bad, "invalid_units")
    all_rejected.append(rej)

    bad = ~is_numeric(df["rate"])
    df, rej = split_good_bad(df, bad, "non_numeric_rate")
    all_rejected.append(rej)

    bad = pd.to_numeric(df["line_amount"], errors="coerce") < 0
    df, rej = split_good_bad(df, bad, "negative_line_amount")
    all_rejected.append(rej)

    bad = is_future_date(df["service_date"])
    df, rej = split_good_bad(df, bad, "future_service_date")
    all_rejected.append(rej)

    rejected = pd.concat(all_rejected, ignore_index=True)
    save_rejected("claim_lines", rejected)
    print(f"  -> {len(df):,} valid rows  |  {len(rejected):,} rejected")
    return df


def clean_payments(df, valid_claim_numbers):
    print(f"\n[CLEAN] payments: {len(df):,} rows in")
    all_rejected = []

    for col in ["payment_id", "claim_number", "payment_date", "payment_method"]:
        bad = is_missing(df[col])
        df, rej = split_good_bad(df, bad, f"missing_{col}")
        all_rejected.append(rej)

    bad = df["claim_number"].astype(str) == "__ORPHAN__"
    df, rej = split_good_bad(df, bad, "orphan_claim_number")
    all_rejected.append(rej)

    bad = ~df["claim_number"].isin(valid_claim_numbers)
    df, rej = split_good_bad(df, bad, "claim_number_not_found")
    all_rejected.append(rej)

    bad = df.duplicated(subset=["payment_id"], keep="first")
    df, rej = split_good_bad(df, bad, "duplicate_payment_id")
    all_rejected.append(rej)

    bad = ~is_numeric(df["paid_amount"])
    df, rej = split_good_bad(df, bad, "non_numeric_paid_amount")
    all_rejected.append(rej)

    bad = is_future_date(df["payment_date"])
    df, rej = split_good_bad(df, bad, "future_payment_date")
    all_rejected.append(rej)

    rejected = pd.concat(all_rejected, ignore_index=True)
    save_rejected("payments", rejected)
    print(f"  -> {len(df):,} valid rows  |  {len(rejected):,} rejected")
    return df


def clean_diagnoses(df, valid_claim_numbers):
    print(f"\n[CLEAN] diagnoses: {len(df):,} rows in")
    all_rejected = []

    for col in ["diag_id", "claim_number", "diagnosis_code", "icd_version"]:
        bad = is_missing(df[col])
        df, rej = split_good_bad(df, bad, f"missing_{col}")
        all_rejected.append(rej)

    bad = df["claim_number"].astype(str) == "__ORPHAN__"
    df, rej = split_good_bad(df, bad, "orphan_claim_number")
    all_rejected.append(rej)

    bad = ~df["claim_number"].isin(valid_claim_numbers)
    df, rej = split_good_bad(df, bad, "claim_number_not_found")
    all_rejected.append(rej)

    bad = df.duplicated(subset=["diag_id"], keep="first")
    df, rej = split_good_bad(df, bad, "duplicate_diag_id")
    all_rejected.append(rej)

    # Only ICD-9 or ICD-10 allowed
    bad = ~df["icd_version"].isin(["ICD-9", "ICD-10"])
    df, rej = split_good_bad(df, bad, "invalid_icd_version")
    all_rejected.append(rej)

    rejected = pd.concat(all_rejected, ignore_index=True)
    save_rejected("diagnoses", rejected)
    print(f"  -> {len(df):,} valid rows  |  {len(rejected):,} rejected")
    return df


def clean_procedures(df, valid_line_ids):
    print(f"\n[CLEAN] procedures: {len(df):,} rows in")
    all_rejected = []

    for col in ["proc_id", "line_id", "procedure_code"]:
        bad = is_missing(df[col])
        df, rej = split_good_bad(df, bad, f"missing_{col}")
        all_rejected.append(rej)

    bad = df["line_id"].astype(str) == "__ORPHAN__"
    df, rej = split_good_bad(df, bad, "orphan_line_id")
    all_rejected.append(rej)

    bad = ~df["line_id"].isin(valid_line_ids)
    df, rej = split_good_bad(df, bad, "line_id_not_found")
    all_rejected.append(rej)

    bad = df.duplicated(subset=["proc_id"], keep="first")
    df, rej = split_good_bad(df, bad, "duplicate_proc_id")
    all_rejected.append(rej)

    rejected = pd.concat(all_rejected, ignore_index=True)
    save_rejected("procedures", rejected)
    print(f"  -> {len(df):,} valid rows  |  {len(rejected):,} rejected")
    return df


# -- Main ------------------------------------------------------

def run_cleaning():
    print("=" * 50)
    print("DATA CLEANING")
    print("=" * 50)

    # Load raw CSVs
    raw = {name: pd.read_csv(path, low_memory=False) for name, path in CSV.items()}

    # Clean in dependency order (parents first, then children)
    members_clean   = clean_members(raw["members"])
    providers_clean = clean_providers(raw["providers"])

    valid_mids  = set(members_clean["member_id"].astype(str))
    valid_pids  = set(providers_clean["provider_id"].astype(str))

    claims_clean = clean_claims(raw["claims"], valid_mids, valid_pids)
    valid_cns    = set(claims_clean["claim_number"].astype(str))

    lines_clean    = clean_claim_lines(raw["claim_lines"], valid_cns)
    payments_clean = clean_payments(raw["payments"], valid_cns)
    diag_clean     = clean_diagnoses(raw["diagnoses"], valid_cns)

    valid_lids  = set(lines_clean["line_id"].astype(str))
    proc_clean  = clean_procedures(raw["procedures"], valid_lids)

    print("\n" + "=" * 50)
    print("CLEANING COMPLETE")
    print("  Rejected rows saved to: output/rejected_rows/")
    print("=" * 50)

    # Return clean DataFrames for the next steps
    return {
        "members":     members_clean,
        "providers":   providers_clean,
        "claims":      claims_clean,
        "claim_lines": lines_clean,
        "payments":    payments_clean,
        "diagnoses":   diag_clean,
        "procedures":  proc_clean,
    }


if __name__ == "__main__":
    run_cleaning()
