"""
STEP 4 - Full ETL Load
=======================
Purpose: Load ALL clean + masked data into the SQLite warehouse.
         This is a FULL load -- it drops existing data and reloads everything.

When to use: First time, or when you want to refresh everything from scratch.

Flow:
  1. Create the star-schema tables (runs the DDL SQL file)
  2. Clean the raw CSVs (calls step2_clean)
  3. Mask the PII (calls step3_mask)
  4. Build dimension tables (dim_date, dim_location, dim_diagnosis, dim_procedure)
  5. Load fact tables (fact_claim, fact_claim_line, fact_payment)
  6. Record the load time in etl_watermark (for incremental later)
"""

import sys
from pathlib import Path

# Add both the capstone root AND the day2_etl folder to path
_capstone_root = Path(__file__).resolve().parent.parent
_etl_dir       = Path(__file__).resolve().parent
sys.path.insert(0, str(_capstone_root))
sys.path.insert(0, str(_etl_dir))

import sqlite3
from datetime import datetime, timezone

import pandas as pd

SQL_DIR = _capstone_root / "day1_sql"

from config import DB_PATH

# Import our earlier steps (direct file imports, no package needed)
from step2_clean import run_cleaning
from step3_mask  import apply_masking


# -- Database helpers ------------------------------------------

def get_connection():
    """Open (or create) the SQLite warehouse database."""
    conn = sqlite3.connect(DB_PATH)
    conn.execute("PRAGMA journal_mode=WAL")   # better write performance
    conn.execute("PRAGMA foreign_keys=ON")    # enforce FK constraints
    return conn


def run_ddl(conn):
    """Create all tables using our DDL SQL file."""
    ddl_path = SQL_DIR / "01_star_schema_ddl.sql"
    sql = ddl_path.read_text(encoding="utf-8")
    conn.executescript(sql)
    conn.commit()
    print("[DDL] Tables created / verified.")


def drop_all_data(conn):
    """Delete all rows from all warehouse tables (fresh full load)."""
    tables = [
        "fact_payment", "fact_claim_line", "fact_claim",
        "dim_member", "dim_provider", "dim_date",
        "dim_diagnosis", "dim_procedure", "dim_location",
        "etl_watermark"
    ]
    for t in tables:
        conn.execute(f"DELETE FROM {t}")
    conn.commit()
    print("[LOAD] Cleared existing data for full reload.")


def load_df(conn, table_name, df):
    """Load a DataFrame into a SQLite table. Handles SQLite's variable limit."""
    # SQLite max variables = 32766; limit chunk to avoid hitting it
    max_vars = 32000
    n_cols = len(df.columns)
    chunk = max(1, max_vars // n_cols)
    df.to_sql(table_name, conn, if_exists="append", index=False, chunksize=chunk)
    print(f"  [LOAD] {table_name}: {len(df):,} rows loaded")


def set_watermark(conn, entity, rows_loaded):
    """Record that we loaded this entity at this time."""
    now = datetime.now(timezone.utc).isoformat()
    conn.execute(
        """INSERT INTO etl_watermark (entity, last_load_dt, rows_loaded)
           VALUES (, , )
           ON CONFLICT(entity) DO UPDATE
           SET last_load_dt=excluded.last_load_dt,
               rows_loaded=excluded.rows_loaded""",
        (entity, now, rows_loaded)
    )
    conn.commit()


# -- Dimension builders ----------------------------------------

def build_dim_date(claims_df, lines_df, payments_df):
    """
    Build the date dimension by collecting all dates from fact tables
    and creating one row per calendar day between min and max date.
    """
    # Collect all date columns
    all_dates = pd.concat([
        pd.to_datetime(claims_df["admit_date"],     errors="coerce"),
        pd.to_datetime(claims_df["discharge_date"], errors="coerce"),
        pd.to_datetime(lines_df["service_date"],    errors="coerce"),
        pd.to_datetime(payments_df["payment_date"], errors="coerce"),
    ]).dropna()

    min_date = all_dates.min().normalize()
    max_date = all_dates.max().normalize()
    print(f"  [DIM_DATE] Covering {min_date.date()} to {max_date.date()}")

    dates = pd.date_range(start=min_date, end=max_date, freq="D")

    dim = pd.DataFrame({"full_date": dates})
    dim["date_key"]     = dim["full_date"].dt.strftime("%Y%m%d").astype(int)
    dim["year"]         = dim["full_date"].dt.year
    dim["quarter"]      = dim["full_date"].dt.quarter
    dim["month"]        = dim["full_date"].dt.month
    dim["month_name"]   = dim["full_date"].dt.strftime("%B")
    dim["day_of_month"] = dim["full_date"].dt.day
    dim["day_name"]     = dim["full_date"].dt.strftime("%A")
    dim["is_weekend"]   = (dim["full_date"].dt.dayofweek >= 5).astype(int)
    dim["full_date"]    = dim["full_date"].dt.strftime("%Y-%m-%d")

    return dim


def build_dim_location(members_df, providers_df):
    """Collect all unique city+state combinations."""
    cols = ["city", "state"]
    locs = pd.concat([
        members_df[cols],
        providers_df[cols],
    ]).drop_duplicates().dropna().reset_index(drop=True)
    locs["location_key"] = (locs["city"] + "|" + locs["state"]).str.upper()
    return locs[["location_key", "city", "state"]]


def build_dim_diagnosis(diagnoses_df):
    """Unique diagnosis codes."""
    dim = (diagnoses_df[["diagnosis_code", "icd_version"]]
           .drop_duplicates(subset=["diagnosis_code"])
           .reset_index(drop=True))
    return dim


def build_dim_procedure(claim_lines_df, procedures_df):
    """Unique procedure codes from both claim_lines and procedures."""
    codes = pd.concat([
        claim_lines_df[["procedure_code"]],
        procedures_df[["procedure_code"]],
    ]).drop_duplicates().dropna().reset_index(drop=True)
    return codes


# -- Fact builders ---------------------------------------------

def to_date_key(series):
    """Convert a datetime column to integer date_key (YYYYMMDD)."""
    return (pd.to_datetime(series, errors="coerce")
              .dt.strftime("%Y%m%d")
              .astype("Int64"))


def build_fact_claim(claims_df):
    df = claims_df.copy()
    df["admit_date"]      = pd.to_datetime(df["admit_date"],     errors="coerce")
    df["discharge_date"]  = pd.to_datetime(df["discharge_date"], errors="coerce")
    df["admit_date_key"]     = df["admit_date"].dt.strftime("%Y%m%d").astype("Int64")
    df["discharge_date_key"] = df["discharge_date"].dt.strftime("%Y%m%d").astype("Int64")
    df["length_of_stay"]  = pd.to_numeric(df["length_of_stay"], errors="coerce").clip(lower=0)
    df["allowed_amount"]  = pd.to_numeric(df["allowed_amount"],  errors="coerce")
    df["paid_amount"]     = pd.to_numeric(df["paid_amount"],     errors="coerce")
    df["denial_code"]     = df["denial_code"].fillna("").astype(str)
    df["is_denied"]       = (df["denial_code"].str.len() > 0).astype(int)
    return df[[
        "claim_number", "member_id", "provider_id",
        "admit_date_key", "discharge_date_key",
        "length_of_stay", "allowed_amount", "paid_amount",
        "denial_code", "is_denied"
    ]]


def build_fact_claim_line(lines_df):
    df = lines_df.copy()
    df["service_date_key"] = to_date_key(df["service_date"])
    df["units"]       = pd.to_numeric(df["units"],       errors="coerce").clip(lower=1)
    df["rate"]        = pd.to_numeric(df["rate"],        errors="coerce")
    df["line_amount"] = pd.to_numeric(df["line_amount"], errors="coerce")
    return df[[
        "line_id", "claim_number", "service_date_key",
        "procedure_code", "units", "rate", "line_amount"
    ]]


def build_fact_payment(payments_df):
    df = payments_df.copy()
    df["payment_date_key"] = to_date_key(df["payment_date"])
    df["paid_amount"]      = pd.to_numeric(df["paid_amount"], errors="coerce")
    df["payment_method"]   = df["payment_method"].astype(str).str.upper()
    df["adjustment_code"]  = df["adjustment_code"].fillna("").astype(str)
    return df[[
        "payment_id", "claim_number", "payment_date_key",
        "payment_method", "paid_amount", "adjustment_code"
    ]]


# -- Dimension: prepare members and providers for loading ------

def prepare_dim_member(members_df):
    """Select only the warehouse columns (PII already masked)."""
    df = members_df.copy()
    df["is_active"] = df["is_active"].astype(str).str.lower().isin(["true","1","yes"]).astype(int)
    df["created_at"] = pd.to_datetime(df["created_at"], errors="coerce").dt.strftime("%Y-%m-%d")
    return df[["member_id","gender","city","state","birth_year","is_active"]]


def prepare_dim_provider(providers_df):
    return providers_df[["provider_id","specialty","city","state","tax_id","license_no"]]


# -- Main ------------------------------------------------------

def run_full_etl():
    print("\n" + "=" * 55)
    print("FULL ETL LOAD")
    print("=" * 55)

    # Step 1: Clean
    clean_dfs = run_cleaning()

    # Step 2: Mask
    masked_dfs = apply_masking(clean_dfs)

    print("\n[LOAD] Connecting to database:", DB_PATH)
    conn = get_connection()

    # Step 3: Create schema + clear old data
    run_ddl(conn)
    drop_all_data(conn)

    print("\n[LOAD] Loading dimension tables ")

    # dim_date
    dim_date = build_dim_date(
        clean_dfs["claims"],
        clean_dfs["claim_lines"],
        clean_dfs["payments"],
    )
    load_df(conn, "dim_date", dim_date)
    set_watermark(conn, "dim_date", len(dim_date))

    # dim_member
    dim_member = prepare_dim_member(masked_dfs["members"])
    load_df(conn, "dim_member", dim_member)
    set_watermark(conn, "dim_member", len(dim_member))

    # dim_provider
    dim_provider = prepare_dim_provider(masked_dfs["providers"])
    load_df(conn, "dim_provider", dim_provider)
    set_watermark(conn, "dim_provider", len(dim_provider))

    # dim_location
    dim_location = build_dim_location(
        masked_dfs["members"],
        masked_dfs["providers"],
    )
    load_df(conn, "dim_location", dim_location)
    set_watermark(conn, "dim_location", len(dim_location))

    # dim_diagnosis
    dim_diag = build_dim_diagnosis(clean_dfs["diagnoses"])
    load_df(conn, "dim_diagnosis", dim_diag)
    set_watermark(conn, "dim_diagnosis", len(dim_diag))

    # dim_procedure
    dim_proc = build_dim_procedure(clean_dfs["claim_lines"], clean_dfs["procedures"])
    load_df(conn, "dim_procedure", dim_proc)
    set_watermark(conn, "dim_procedure", len(dim_proc))

    print("\n[LOAD] Loading fact tables ")

    # fact_claim
    fact_claim = build_fact_claim(masked_dfs["claims"])
    load_df(conn, "fact_claim", fact_claim)
    set_watermark(conn, "fact_claim", len(fact_claim))

    # fact_claim_line
    fact_line = build_fact_claim_line(masked_dfs["claim_lines"])
    load_df(conn, "fact_claim_line", fact_line)
    set_watermark(conn, "fact_claim_line", len(fact_line))

    # fact_payment
    fact_pay = build_fact_payment(masked_dfs["payments"])
    load_df(conn, "fact_payment", fact_pay)
    set_watermark(conn, "fact_payment", len(fact_pay))

    conn.close()

    print("\n" + "=" * 55)
    print("FULL ETL LOAD COMPLETE")
    print(f"  Warehouse: {DB_PATH}")
    print("  Next step: run step5_load_incremental.py for future loads")
    print("=" * 55)


if __name__ == "__main__":
    run_full_etl()
