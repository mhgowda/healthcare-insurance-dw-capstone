"""
STEP 5 - Incremental ETL Load
==============================
Purpose: Load ONLY new data since the last run.
         Does NOT drop or reload existing data.

How it works:
  1. Read the last_load_dt from etl_watermark for each entity
  2. Filter the CSV to rows where the date column > last_load_dt
  3. Clean, mask, and load only those new rows
  4. Update the watermark to today

Why incremental
  In real projects, CSVs or databases are updated daily.
  Re-loading 1.4 million rows every day wastes time.
  With incremental loads, you only process what's new.

Watermark columns used:
  members     -> created_at
  claims      -> admit_date
  claim_lines -> service_date
  payments    -> payment_date
  providers/diagnoses/procedures -> always refresh (small tables, no date column)
"""

import sys
from pathlib import Path

_capstone_root = Path(__file__).resolve().parent.parent
_etl_dir       = Path(__file__).resolve().parent
sys.path.insert(0, str(_capstone_root))
sys.path.insert(0, str(_etl_dir))

import sqlite3
from datetime import datetime, timezone

import pandas as pd

SQL_DIR = _capstone_root / "day1_sql"
from config import DB_PATH, CSV, MASKING_SALT

from step2_clean import (
    clean_members, clean_providers, clean_claims,
    clean_claim_lines, clean_payments, clean_diagnoses, clean_procedures,
)
from step3_mask import apply_masking, build_mapping
from step4_load_full import (
    get_connection, run_ddl,
    prepare_dim_member, prepare_dim_provider,
    build_fact_claim, build_fact_claim_line, build_fact_payment,
    build_dim_diagnosis, build_dim_procedure,
)


def get_watermark(conn, entity):
    """Read the last load timestamp for an entity. Returns None if first run."""
    row = conn.execute(
        "SELECT last_load_dt FROM etl_watermark WHERE entity = ",
        (entity,)
    ).fetchone()
    if row and row[0]:
        return pd.Timestamp(row[0])
    return None


def update_watermark(conn, entity, rows_loaded):
    now = datetime.now(timezone.utc).isoformat()
    conn.execute(
        """INSERT INTO etl_watermark (entity, last_load_dt, rows_loaded)
           VALUES (, , )
           ON CONFLICT(entity) DO UPDATE
           SET last_load_dt=excluded.last_load_dt,
               rows_loaded=etl_watermark.rows_loaded + excluded.rows_loaded""",
        (entity, now, rows_loaded)
    )
    conn.commit()


def filter_new_rows(df, date_col, watermark_dt):
    """Keep only rows where date_col > watermark_dt."""
    if watermark_dt is None:
        print(f"    No watermark found -- loading all rows.")
        return df
    dt = pd.to_datetime(df[date_col], errors="coerce")
    new_rows = df[dt > watermark_dt].reset_index(drop=True)
    print(f"    Watermark: {watermark_dt.date()}  ->  {len(new_rows):,} new rows")
    return new_rows


def insert_new(conn, table_name, df):
    """Insert new rows, ignore duplicates (INSERT OR IGNORE in SQLite)."""
    if df.empty:
        print(f"  [SKIP] {table_name}: no new rows")
        return 0

    # Use INSERT OR IGNORE to safely skip existing PKs
    cols = list(df.columns)
    placeholders = ", ".join([""] * len(cols))
    col_names    = ", ".join(cols)
    sql = f"INSERT OR IGNORE INTO {table_name} ({col_names}) VALUES ({placeholders})"

    rows = [tuple(r) for r in df.itertuples(index=False, name=None)]
    conn.executemany(sql, rows)
    conn.commit()
    print(f"  [INC] {table_name}: {len(df):,} rows attempted -> {conn.total_changes} inserted")
    return len(df)


def run_incremental_etl():
    print("\n" + "=" * 55)
    print("INCREMENTAL ETL LOAD")
    print("=" * 55)

    conn = get_connection()
    run_ddl(conn)  # ensures tables exist (safe, uses IF NOT EXISTS)

    salt = MASKING_SALT

    # -- Load current state from warehouse (for FK validation) -
    existing_members  = set(pd.read_sql("SELECT member_id FROM dim_member",   conn)["member_id"])
    existing_providers= set(pd.read_sql("SELECT provider_id FROM dim_provider",conn)["provider_id"])
    existing_claims   = set(pd.read_sql("SELECT claim_number FROM fact_claim", conn)["claim_number"])
    existing_lines    = set(pd.read_sql("SELECT line_id FROM fact_claim_line", conn)["line_id"])

    # -- Members (watermark on created_at) ---------------------
    print("\n[INC] members")
    wm_members = get_watermark(conn, "dim_member")
    raw_m = pd.read_csv(CSV["members"], low_memory=False)
    new_m = filter_new_rows(raw_m, "created_at", wm_members)

    if not new_m.empty:
        # We need raw member_ids for FK map building
        clean_m = clean_members(new_m)
        from day2_etl.step3_mask import mask_members
        masked_m, member_id_map = mask_members(clean_m, salt)
        dim_m = prepare_dim_member(masked_m)
        n = insert_new(conn, "dim_member", dim_m)
        update_watermark(conn, "dim_member", n)
    else:
        member_id_map = {}

    # -- Providers (no date col -- refresh all, ignore duplicates) -
    print("\n[INC] providers")
    raw_p = pd.read_csv(CSV["providers"], low_memory=False)
    clean_p = clean_providers(raw_p)
    from day2_etl.step3_mask import mask_providers
    masked_p, provider_id_map = mask_providers(clean_p, salt)
    dim_p = prepare_dim_provider(masked_p)
    n = insert_new(conn, "dim_provider", dim_p)
    update_watermark(conn, "dim_provider", n)

    # -- Claims (watermark on admit_date) ----------------------
    print("\n[INC] claims")
    wm_claims = get_watermark(conn, "fact_claim")
    raw_c = pd.read_csv(CSV["claims"], low_memory=False)
    new_c = filter_new_rows(raw_c, "admit_date", wm_claims)

    if not new_c.empty:
        # Use ALL valid member/provider IDs (existing + potentially new)
        all_mids = existing_members | set(member_id_map.values())
        all_pids = existing_providers | set(provider_id_map.values())

        # Clean using raw IDs (before masking)
        raw_mid_set  = set(pd.read_csv(CSV["members"],   usecols=["member_id"])["member_id"].astype(str))
        raw_pid_set  = set(pd.read_csv(CSV["providers"], usecols=["provider_id"])["provider_id"].astype(str))
        clean_c = clean_claims(new_c, raw_mid_set, raw_pid_set)

        from day2_etl.step3_mask import mask_claims
        # Rebuild member/provider maps from full dataset
        full_m_ids  = pd.read_csv(CSV["members"],   usecols=["member_id"])["member_id"]
        full_p_ids  = pd.read_csv(CSV["providers"],  usecols=["provider_id"])["provider_id"]
        from day2_etl.step3_mask import build_mapping
        full_mid_map = build_mapping(full_m_ids, salt, "MBR")
        full_pid_map = build_mapping(full_p_ids, salt, "PRV")

        masked_c, claim_map = mask_claims(clean_c, salt, full_mid_map, full_pid_map)
        fact_c = build_fact_claim(masked_c)
        n = insert_new(conn, "fact_claim", fact_c)
        update_watermark(conn, "fact_claim", n)
    else:
        claim_map = {}

    # -- Claim Lines (watermark on service_date) ---------------
    print("\n[INC] claim_lines")
    wm_lines = get_watermark(conn, "fact_claim_line")
    raw_cl = pd.read_csv(CSV["claim_lines"], low_memory=False)
    new_cl = filter_new_rows(raw_cl, "service_date", wm_lines)

    if not new_cl.empty:
        all_cns = existing_claims | set(claim_map.values())
        raw_cn_set = set(pd.read_csv(CSV["claims"], usecols=["claim_number"])["claim_number"].astype(str))
        clean_cl = clean_claim_lines(new_cl, raw_cn_set)

        from day2_etl.step3_mask import mask_claim_lines
        full_cn_series = pd.read_csv(CSV["claims"], usecols=["claim_number"])["claim_number"]
        full_claim_map = build_mapping(full_cn_series, salt, "CLM")
        masked_cl, line_map = mask_claim_lines(clean_cl, salt, full_claim_map)
        fact_cl = build_fact_claim_line(masked_cl)
        n = insert_new(conn, "fact_claim_line", fact_cl)
        update_watermark(conn, "fact_claim_line", n)
    else:
        line_map = {}

    # -- Payments (watermark on payment_date) ------------------
    print("\n[INC] payments")
    wm_pay = get_watermark(conn, "fact_payment")
    raw_py = pd.read_csv(CSV["payments"], low_memory=False)
    new_py = filter_new_rows(raw_py, "payment_date", wm_pay)

    if not new_py.empty:
        raw_cn_set = set(pd.read_csv(CSV["claims"], usecols=["claim_number"])["claim_number"].astype(str))
        clean_py = clean_payments(new_py, raw_cn_set)

        from day2_etl.step3_mask import mask_payments
        full_cn_series = pd.read_csv(CSV["claims"], usecols=["claim_number"])["claim_number"]
        full_claim_map = build_mapping(full_cn_series, salt, "CLM")
        masked_py = mask_payments(clean_py, salt, full_claim_map)
        fact_py = build_fact_payment(masked_py)
        n = insert_new(conn, "fact_payment", fact_py)
        update_watermark(conn, "fact_payment", n)

    # -- Diagnoses + Procedures (small, always refresh dims) ---
    print("\n[INC] diagnoses + procedures (dimension refresh)")
    raw_dx = pd.read_csv(CSV["diagnoses"], low_memory=False)
    clean_dx = clean_diagnoses(raw_dx, valid_claim_numbers=None)  # no FK check for dim
    dim_dx = build_dim_diagnosis(clean_dx)
    insert_new(conn, "dim_diagnosis", dim_dx)

    raw_pr = pd.read_csv(CSV["procedures"], low_memory=False)
    raw_cl_all = pd.read_csv(CSV["claim_lines"], usecols=["procedure_code"])
    dim_pr = build_dim_procedure(raw_cl_all, raw_pr[["procedure_code"]].dropna())
    insert_new(conn, "dim_procedure", dim_pr)

    conn.close()

    print("\n" + "=" * 55)
    print("INCREMENTAL ETL COMPLETE")
    print(f"  Warehouse: {DB_PATH}")
    print("=" * 55)


if __name__ == "__main__":
    run_incremental_etl()
