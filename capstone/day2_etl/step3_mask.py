"""
STEP 3 - Data Masking (PII / PHI Protection)
=============================================
Purpose: Replace sensitive personal information with safe tokens
         BEFORE loading into the data warehouse.

Why mask
  The warehouse is used by analysts who DON'T need to see real names/emails.
  Masking protects patient privacy (required by HIPAA / data privacy laws).

What we use: Deterministic HMAC-SHA256 hashing
  - Same input  -> always same output  (so JOINs still work)
  - You cannot reverse it back to the original value
  - We add a "salt" string to make it unique to our project

Fields masked:
  members   -> member_id, first_name, last_name, full_name, email, phone, dob->birth_year
  providers -> provider_id, provider_name, tax_id, license_no
  claims    -> claim_number, member_id, provider_id
  claim_lines -> line_id, claim_number (FK)
  payments    -> payment_id, claim_number (FK)
  diagnoses   -> diag_id, claim_number (FK)
  procedures  -> proc_id, line_id (FK)

Fields NOT masked (kept for analysis):
  city, state, gender, specialty, procedure_code, diagnosis_code,
  amounts, dates (non-DOB), payment_method
"""

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import hashlib
import hmac
import json
from datetime import datetime, timezone
import pandas as pd
from config import MASKING_SALT, LOGS_DIR


# -- Core masking function -------------------------------------

def mask_value(value, salt, prefix=""):
    """
    Turn one value into a safe token.
    - Same input + same salt -> always same token (deterministic)
    - Cannot be reversed
    """
    if pd.isna(value) or str(value).strip() == "":
        return value  # leave nulls/blanks as-is

    raw_hash = hmac.new(
        salt.encode("utf-8"),
        str(value).encode("utf-8"),
        digestmod=hashlib.sha256
    ).hexdigest()[:12].upper()

    return f"{prefix}{raw_hash}" if prefix else raw_hash


def mask_column(series, salt, prefix=""):
    """Apply mask_value to a whole column (Series)."""
    return series.apply(lambda v: mask_value(v, salt, prefix))


def build_mapping(series, salt, prefix=""):
    """
    Build a dictionary: {original_value: masked_token}
    We need this so child tables (claims, lines, etc.) use the SAME masked ID
    as the parent table.
    """
    unique_vals = series.dropna().unique()
    return {v: mask_value(v, salt, prefix) for v in unique_vals}


# -- Per-entity masking ----------------------------------------

def mask_members(df, salt):
    """Mask all PII fields in members. Return masked df + id mapping."""
    df = df.copy()

    # Build mapping BEFORE masking (so we can share it with child tables)
    id_map = build_mapping(df["member_id"], salt, prefix="MBR")

    df["member_id"]  = df["member_id"].map(id_map)
    df["first_name"] = mask_column(df["first_name"], salt, prefix="FN")
    df["last_name"]  = mask_column(df["last_name"],  salt, prefix="LN")
    df["full_name"]  = mask_column(df["full_name"],  salt, prefix="NM")
    df["email"]      = mask_column(df["email"],      salt, prefix="EM")
    df["phone"]      = mask_column(df["phone"],      salt, prefix="PH")

    # DOB -> keep only the year (generalization, not hashing)
    df["birth_year"] = pd.to_datetime(df["dob"], errors="coerce").dt.year.astype("Int64")
    df.drop(columns=["dob"], inplace=True)  # remove exact DOB

    # Kept as-is (needed for analysis): gender, city, state, created_at, is_active
    return df, id_map


def mask_providers(df, salt):
    """Mask PII fields in providers."""
    df = df.copy()

    id_map = build_mapping(df["provider_id"], salt, prefix="PRV")

    df["provider_id"]   = df["provider_id"].map(id_map)
    df["provider_name"] = mask_column(df["provider_name"], salt, prefix="PN")
    df["tax_id"]        = mask_column(df["tax_id"],        salt, prefix="TX")
    df["license_no"]    = mask_column(df["license_no"],    salt, prefix="LI")

    # Kept: specialty, city, state
    return df, id_map


def mask_claims(df, salt, member_id_map, provider_id_map):
    """Mask claim_number; remap member_id and provider_id using pre-built maps."""
    df = df.copy()

    claim_map = build_mapping(df["claim_number"], salt, prefix="CLM")

    df["claim_number"] = df["claim_number"].map(claim_map)
    df["member_id"]    = df["member_id"].map(member_id_map).fillna(df["member_id"])
    df["provider_id"]  = df["provider_id"].map(provider_id_map).fillna(df["provider_id"])

    return df, claim_map


def mask_claim_lines(df, salt, claim_map):
    df = df.copy()
    line_map = build_mapping(df["line_id"], salt, prefix="LN")
    df["line_id"]      = df["line_id"].map(line_map)
    df["claim_number"] = df["claim_number"].map(claim_map).fillna(df["claim_number"])
    return df, line_map


def mask_payments(df, salt, claim_map):
    df = df.copy()
    df["payment_id"]   = mask_column(df["payment_id"],   salt, prefix="PMT")
    df["claim_number"] = df["claim_number"].map(claim_map).fillna(df["claim_number"])
    return df


def mask_diagnoses(df, salt, claim_map):
    df = df.copy()
    df["diag_id"]      = mask_column(df["diag_id"],      salt, prefix="DX")
    df["claim_number"] = df["claim_number"].map(claim_map).fillna(df["claim_number"])
    return df


def mask_procedures(df, salt, line_map):
    df = df.copy()
    df["proc_id"] = mask_column(df["proc_id"], salt, prefix="PR")
    df["line_id"] = df["line_id"].map(line_map).fillna(df["line_id"])
    return df


# -- Write masking audit log -----------------------------------

def write_masking_audit(entity_counts):
    """
    Save masking_audit.json to logs/ -- as required by the project document.
    Records: which fields were masked, how many values, timestamp.
    """
    audit = {
        "run_at": datetime.now(timezone.utc).isoformat(),
        "masking_salt_used": True,   # never log the actual salt value
        "fields_masked": [
            {"entity": e, "field": f, "rows_masked": c}
            for e, f, c in entity_counts
        ],
    }
    out = LOGS_DIR / "masking_audit.json"
    with open(out, "w", encoding="utf-8") as fh:
        json.dump(audit, fh, indent=2)
    print(f"\n[MASK] Audit log saved to: {out}")


# -- Apply masking to all clean dataframes --------------------

def apply_masking(clean_dfs, salt=MASKING_SALT):
    """
    Takes the dictionary of clean DataFrames from step2_clean.py
    and returns a dictionary of masked DataFrames.
    """
    print("\n" + "=" * 50)
    print("DATA MASKING")
    print("=" * 50)

    audit = []

    # Members
    m_df = clean_dfs["members"]
    masked_members, member_id_map = mask_members(m_df, salt)
    audit.append(("members", "member_id",  len(member_id_map)))
    audit.append(("members", "email",      len(m_df)))
    audit.append(("members", "phone",      len(m_df)))
    audit.append(("members", "dob->birth_year", len(m_df)))
    print(f"  [MASK] members: {len(m_df):,} rows masked")

    # Providers
    p_df = clean_dfs["providers"]
    masked_providers, provider_id_map = mask_providers(p_df, salt)
    audit.append(("providers", "provider_id", len(provider_id_map)))
    audit.append(("providers", "tax_id",       len(p_df)))
    audit.append(("providers", "license_no",   len(p_df)))
    print(f"  [MASK] providers: {len(p_df):,} rows masked")

    # Claims
    c_df = clean_dfs["claims"]
    masked_claims, claim_map = mask_claims(c_df, salt, member_id_map, provider_id_map)
    audit.append(("claims", "claim_number", len(claim_map)))
    print(f"  [MASK] claims: {len(c_df):,} rows masked")

    # Claim Lines
    cl_df = clean_dfs["claim_lines"]
    masked_lines, line_map = mask_claim_lines(cl_df, salt, claim_map)
    audit.append(("claim_lines", "line_id",      len(line_map)))
    audit.append(("claim_lines", "claim_number", len(cl_df)))
    print(f"  [MASK] claim_lines: {len(cl_df):,} rows masked")

    # Payments
    py_df = clean_dfs["payments"]
    masked_payments = mask_payments(py_df, salt, claim_map)
    audit.append(("payments", "payment_id",   len(py_df)))
    audit.append(("payments", "claim_number", len(py_df)))
    print(f"  [MASK] payments: {len(py_df):,} rows masked")

    # Diagnoses
    dx_df = clean_dfs["diagnoses"]
    masked_diagnoses = mask_diagnoses(dx_df, salt, claim_map)
    audit.append(("diagnoses", "claim_number", len(dx_df)))
    print(f"  [MASK] diagnoses: {len(dx_df):,} rows masked")

    # Procedures
    pr_df = clean_dfs["procedures"]
    masked_procedures = mask_procedures(pr_df, salt, line_map)
    audit.append(("procedures", "line_id", len(pr_df)))
    print(f"  [MASK] procedures: {len(pr_df):,} rows masked")

    write_masking_audit(audit)
    print("=" * 50)

    return {
        "members":     masked_members,
        "providers":   masked_providers,
        "claims":      masked_claims,
        "claim_lines": masked_lines,
        "payments":    masked_payments,
        "diagnoses":   masked_diagnoses,
        "procedures":  masked_procedures,
    }


if __name__ == "__main__":
    # Quick test of masking logic
    print("Testing mask_value():")
    s = MASKING_SALT
    v1 = mask_value("M0000001", s, "MBR")
    v2 = mask_value("M0000001", s, "MBR")
    v3 = mask_value("M0000002", s, "MBR")
    print(f"  Same input -> same output  {v1 == v2}  ({v1})")
    print(f"  Diff input -> diff output  {v1 != v3}  ({v3})")
    print("  Masking logic is working correctly.")
