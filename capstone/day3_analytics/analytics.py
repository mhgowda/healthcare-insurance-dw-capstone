"""
DAY 3 - Python Analytics
=========================
Purpose: Connect to the warehouse database and answer business questions
         using Pandas + SQL.

Run this AFTER the full ETL load (step4_load_full.py).

Questions answered:
  1. Total claims summary
  2. Monthly paid trend
  3. Top 10 procedures
  4. Provider performance
  5. Denial rate breakdown
  6. Length of stay distribution
  7. Member utilisation (top spenders)
  8. Geographic breakdown by state
  9. Payment method breakdown
"""

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import sqlite3
import pandas as pd
from config import DB_PATH, OUTPUT_DIR


def connect():
    """Open a read-only connection to the warehouse."""
    return sqlite3.connect(DB_PATH)


def run_query(conn, sql, description=""):
    """Run a SQL query and return a DataFrame."""
    df = pd.read_sql_query(sql, conn)
    if description:
        print(f"\n{'-'*50}")
        print(f"  {description}")
        print(f"{'-'*50}")
        print(df.to_string(index=False))
    return df


# -- 1. Total claims summary -----------------------------------

def q1_total_summary(conn):
    sql = """
    SELECT
        COUNT(*)                        AS total_claims,
        ROUND(SUM(allowed_amount), 2)   AS total_allowed,
        ROUND(SUM(paid_amount), 2)       AS total_paid,
        ROUND(AVG(paid_amount), 2)       AS avg_paid_per_claim,
        SUM(is_denied)                  AS total_denied,
        ROUND(100.0 * SUM(is_denied) / COUNT(*), 2) AS denial_rate_pct
    FROM fact_claim
    """
    return run_query(conn, sql, "Q1 - Total Claims Summary")


# -- 2. Monthly paid trend -------------------------------------

def q2_monthly_trend(conn):
    sql = """
    SELECT
        d.year,
        d.month,
        d.month_name,
        COUNT(fc.claim_number)          AS claims,
        ROUND(SUM(fc.paid_amount), 2)   AS total_paid,
        ROUND(AVG(fc.paid_amount), 2)   AS avg_paid
    FROM fact_claim fc
    JOIN dim_date d ON fc.admit_date_key = d.date_key
    GROUP BY d.year, d.month, d.month_name
    ORDER BY d.year, d.month
    """
    return run_query(conn, sql, "Q2 - Monthly Paid Trend")


# -- 3. Top 10 procedures --------------------------------------

def q3_top_procedures(conn):
    sql = """
    SELECT
        procedure_code,
        COUNT(*)                        AS times_used,
        SUM(units)                      AS total_units,
        ROUND(SUM(line_amount), 2)       AS total_billed,
        ROUND(AVG(line_amount), 2)       AS avg_per_line
    FROM fact_claim_line
    GROUP BY procedure_code
    ORDER BY times_used DESC
    LIMIT 10
    """
    return run_query(conn, sql, "Q3 - Top 10 Procedures by Utilisation")


# -- 4. Provider performance -----------------------------------

def q4_provider_performance(conn):
    sql = """
    SELECT
        dp.provider_id,
        dp.specialty,
        dp.state,
        COUNT(fc.claim_number)          AS total_claims,
        ROUND(SUM(fc.paid_amount), 2)   AS total_paid,
        ROUND(AVG(fc.paid_amount), 2)   AS avg_paid,
        ROUND(AVG(fc.length_of_stay), 1) AS avg_los,
        SUM(fc.is_denied)               AS denied_claims,
        ROUND(100.0 * SUM(fc.is_denied) / COUNT(*), 2) AS denial_rate_pct
    FROM fact_claim fc
    JOIN dim_provider dp ON fc.provider_id = dp.provider_id
    GROUP BY dp.provider_id, dp.specialty, dp.state
    ORDER BY total_paid DESC
    LIMIT 15
    """
    return run_query(conn, sql, "Q4 - Provider Performance (Top 15)")


# -- 5. Denial rate by code ------------------------------------

def q5_denial_rate(conn):
    sql = """
    SELECT
        CASE WHEN denial_code = '' OR denial_code IS NULL THEN 'NO DENIAL'
             ELSE denial_code END AS denial_reason,
        COUNT(*)                   AS claim_count,
        ROUND(100.0 * COUNT(*) / (SELECT COUNT(*) FROM fact_claim), 2) AS pct_of_total
    FROM fact_claim
    GROUP BY denial_reason
    ORDER BY claim_count DESC
    """
    return run_query(conn, sql, "Q5 - Denial Rate Breakdown")


# -- 6. Length of stay distribution ---------------------------

def q6_los_distribution(conn):
    sql = """
    SELECT
        length_of_stay,
        COUNT(*)                        AS claim_count,
        ROUND(AVG(paid_amount), 2)       AS avg_paid
    FROM fact_claim
    WHERE length_of_stay >= 0 AND length_of_stay <= 30
    GROUP BY length_of_stay
    ORDER BY length_of_stay
    """
    return run_query(conn, sql, "Q6 - Length of Stay Distribution (0-30 days)")


# -- 7. Member utilisation -------------------------------------

def q7_member_utilisation(conn):
    sql = """
    SELECT
        dm.member_id,
        dm.gender,
        dm.state,
        dm.birth_year,
        COUNT(fc.claim_number)          AS total_claims,
        ROUND(SUM(fc.paid_amount), 2)   AS total_paid,
        ROUND(AVG(fc.length_of_stay), 1) AS avg_los
    FROM fact_claim fc
    JOIN dim_member dm ON fc.member_id = dm.member_id
    GROUP BY dm.member_id, dm.gender, dm.state, dm.birth_year
    ORDER BY total_paid DESC
    LIMIT 20
    """
    return run_query(conn, sql, "Q7 - Top 20 Members by Total Paid")


# -- 8. Geographic breakdown -----------------------------------

def q8_geographic(conn):
    sql = """
    SELECT
        dm.state,
        COUNT(fc.claim_number)          AS total_claims,
        ROUND(SUM(fc.paid_amount), 2)   AS total_paid,
        ROUND(AVG(fc.paid_amount), 2)   AS avg_paid,
        SUM(fc.is_denied)               AS denied_claims
    FROM fact_claim fc
    JOIN dim_member dm ON fc.member_id = dm.member_id
    GROUP BY dm.state
    ORDER BY total_paid DESC
    LIMIT 15
    """
    return run_query(conn, sql, "Q8 - Geographic Breakdown by State")


# -- 9. Payment method breakdown -------------------------------

def q9_payment_methods(conn):
    sql = """
    SELECT
        payment_method,
        COUNT(*)                        AS count,
        ROUND(SUM(paid_amount), 2)       AS total_paid,
        ROUND(AVG(paid_amount), 2)       AS avg_payment
    FROM fact_payment
    GROUP BY payment_method
    ORDER BY total_paid DESC
    """
    return run_query(conn, sql, "Q9 - Payment Method Breakdown")


# -- Run all + save results ------------------------------------

def run_all_analytics():
    print("\n" + "=" * 55)
    print("ANALYTICS - Running all queries against the warehouse")
    print(f"  Database: {DB_PATH}")
    print("=" * 55)

    if not DB_PATH.exists():
        print("\nERROR: Warehouse database not found.")
        print("  Please run step4_load_full.py first.")
        return

    conn = connect()

    results = {
        "q1_summary":          q1_total_summary(conn),
        "q2_monthly_trend":    q2_monthly_trend(conn),
        "q3_top_procedures":   q3_top_procedures(conn),
        "q4_provider_perf":    q4_provider_performance(conn),
        "q5_denial_rate":      q5_denial_rate(conn),
        "q6_los_dist":         q6_los_distribution(conn),
        "q7_member_util":      q7_member_utilisation(conn),
        "q8_geographic":       q8_geographic(conn),
        "q9_payment_methods":  q9_payment_methods(conn),
    }

    conn.close()

    # Save each result as CSV
    for name, df in results.items():
        path = OUTPUT_DIR / f"{name}.csv"
        df.to_csv(path, index=False)

    print(f"\n[ANALYTICS] All results saved to: {OUTPUT_DIR}")
    print("  Open the CSV files to review the output.")

    return results


if __name__ == "__main__":
    run_all_analytics()
