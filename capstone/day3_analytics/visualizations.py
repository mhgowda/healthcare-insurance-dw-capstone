"""
DAY 3 - Matplotlib Visualizations
===================================
Purpose: Create charts from the warehouse data and save them as PNG images.
Output:  output/charts/  (8 chart files)

Charts:
  01 - Monthly paid amount (bar chart)
  02 - Top 10 procedures (horizontal bar)
  03 - Provider performance: total paid (horizontal bar)
  04 - Denial rate breakdown (bar chart)
  05 - Length of stay distribution (histogram/bar)
  06 - Member utilisation scatter (claims vs paid)
  07 - Geographic breakdown by state (bar)
  08 - Payment method breakdown (pie chart)

Run AFTER step4_load_full.py
"""

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

# Use non-interactive backend (saves files without opening a window)
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import sqlite3
import pandas as pd

from config import DB_PATH, CHARTS_DIR


def connect():
    return sqlite3.connect(DB_PATH)


def query(conn, sql):
    return pd.read_sql_query(sql, conn)


def save_chart(fig, filename, title=""):
    path = CHARTS_DIR / filename
    fig.savefig(path, dpi=120, bbox_inches="tight")
    plt.close(fig)
    print(f"  Saved: {path.name}")


# -- Chart 1: Monthly Paid Amount -----------------------------

def chart_monthly_paid(conn):
    df = query(conn, """
        SELECT d.year, d.month, d.month_name,
               ROUND(SUM(fc.paid_amount)/1000000, 2) AS paid_millions
        FROM fact_claim fc
        JOIN dim_date d ON fc.admit_date_key = d.date_key
        GROUP BY d.year, d.month, d.month_name
        ORDER BY d.year, d.month
    """)
    if df.empty:
        print("  [SKIP] No data for monthly paid chart")
        return

    df["label"] = df["month_name"].str[:3] + " " + df["year"].astype(str)

    fig, ax = plt.subplots(figsize=(14, 5))
    ax.bar(df["label"], df["paid_millions"], color="#4C72B0", alpha=0.85)
    ax.set_title("Monthly Total Paid Amount", fontsize=14, fontweight="bold")
    ax.set_xlabel("Month")
    ax.set_ylabel("Paid Amount (Millions Rs)")
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"Rs{x:.1f}M"))
    plt.xticks(rotation=60, ha="right", fontsize=7)
    plt.tight_layout()
    save_chart(fig, "01_monthly_paid.png")


# -- Chart 2: Top 10 Procedures -------------------------------

def chart_top_procedures(conn):
    df = query(conn, """
        SELECT procedure_code, COUNT(*) AS times_used
        FROM fact_claim_line
        GROUP BY procedure_code
        ORDER BY times_used DESC
        LIMIT 10
    """)
    if df.empty:
        print("  [SKIP] No data for procedures chart")
        return

    fig, ax = plt.subplots(figsize=(10, 6))
    ax.barh(df["procedure_code"], df["times_used"], color="#55A868", alpha=0.85)
    ax.set_title("Top 10 Procedures by Utilisation", fontsize=14, fontweight="bold")
    ax.set_xlabel("Number of Claim Lines")
    ax.invert_yaxis()
    ax.xaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"{x:,.0f}"))
    plt.tight_layout()
    save_chart(fig, "02_top_procedures.png")


# -- Chart 3: Provider Performance ----------------------------

def chart_provider_performance(conn):
    df = query(conn, """
        SELECT dp.provider_id, dp.specialty,
               ROUND(SUM(fc.paid_amount)/1000, 1) AS paid_thousands,
               ROUND(100.0 * SUM(fc.is_denied) / COUNT(*), 1) AS denial_rate
        FROM fact_claim fc
        JOIN dim_provider dp ON fc.provider_id = dp.provider_id
        GROUP BY dp.provider_id, dp.specialty
        ORDER BY paid_thousands DESC
        LIMIT 15
    """)
    if df.empty:
        print("  [SKIP] No data for provider chart")
        return

    fig, axes = plt.subplots(1, 2, figsize=(16, 6))

    # Left: Total Paid
    ax1 = axes[0]
    ax1.barh(df["provider_id"], df["paid_thousands"], color="#C44E52", alpha=0.85)
    ax1.set_title("Top 15 Providers - Total Paid", fontsize=12)
    ax1.set_xlabel("Total Paid (Thousands Rs)")
    ax1.invert_yaxis()
    ax1.tick_params(axis="y", labelsize=7)

    # Right: Denial Rate
    ax2 = axes[1]
    colours = ["#C44E52" if r > 10 else "#55A868" for r in df["denial_rate"].fillna(0)]
    ax2.barh(df["provider_id"], df["denial_rate"].fillna(0), color=colours, alpha=0.85)
    ax2.set_title("Top 15 Providers - Denial Rate %", fontsize=12)
    ax2.set_xlabel("Denial Rate (%)")
    ax2.invert_yaxis()
    ax2.tick_params(axis="y", labelsize=7)

    fig.suptitle("Provider Performance", fontsize=14, fontweight="bold")
    plt.tight_layout()
    save_chart(fig, "03_provider_performance.png")


# -- Chart 4: Denial Rate Breakdown ---------------------------

def chart_denial_rate(conn):
    df = query(conn, """
        SELECT
            CASE WHEN denial_code = '' OR denial_code IS NULL THEN 'NO DENIAL'
                 ELSE denial_code END AS denial_code,
            COUNT(*) AS count
        FROM fact_claim
        GROUP BY denial_code
        ORDER BY count DESC
    """)
    if df.empty:
        print("  [SKIP] No data for denial chart")
        return

    fig, ax = plt.subplots(figsize=(8, 5))
    colours = ["#4C72B0" if c == "NO DENIAL" else "#C44E52"
               for c in df["denial_code"]]
    ax.bar(df["denial_code"].astype(str), df["count"], color=colours, alpha=0.85)
    ax.set_title("Claims by Denial Code", fontsize=14, fontweight="bold")
    ax.set_xlabel("Denial Code")
    ax.set_ylabel("Number of Claims")
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"{x:,.0f}"))
    plt.xticks(rotation=30, ha="right")
    plt.tight_layout()
    save_chart(fig, "04_denial_rate.png")


# -- Chart 5: Length of Stay Distribution ---------------------

def chart_los(conn):
    df = query(conn, """
        SELECT length_of_stay, COUNT(*) AS count
        FROM fact_claim
        WHERE length_of_stay >= 0 AND length_of_stay <= 30
        GROUP BY length_of_stay
        ORDER BY length_of_stay
    """)
    if df.empty:
        print("  [SKIP] No data for LOS chart")
        return

    fig, ax = plt.subplots(figsize=(12, 5))
    ax.bar(df["length_of_stay"], df["count"], color="#8172B2", alpha=0.85, width=0.8)
    ax.set_title("Length of Stay Distribution (0-30 Days)", fontsize=14, fontweight="bold")
    ax.set_xlabel("Days in Hospital")
    ax.set_ylabel("Number of Claims")
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"{x:,.0f}"))
    plt.tight_layout()
    save_chart(fig, "05_los_distribution.png")


# -- Chart 6: Member Utilisation Scatter ----------------------

def chart_member_utilisation(conn):
    df = query(conn, """
        SELECT dm.gender,
               COUNT(fc.claim_number)        AS total_claims,
               ROUND(SUM(fc.paid_amount),2)  AS total_paid
        FROM fact_claim fc
        JOIN dim_member dm ON fc.member_id = dm.member_id
        GROUP BY dm.member_id, dm.gender
        ORDER BY total_paid DESC
        LIMIT 200
    """)
    if df.empty:
        print("  [SKIP] No data for member utilisation chart")
        return

    fig, ax = plt.subplots(figsize=(10, 6))
    for gender, colour, marker in [("male","#4C72B0","o"), ("female","#DD8452","^")]:
        sub = df[df["gender"] == gender]
        ax.scatter(sub["total_claims"], sub["total_paid"] / 1000,
                   label=gender.capitalize(), color=colour, alpha=0.6, s=40, marker=marker)

    ax.set_title("Member Utilisation: Claims vs Total Paid (Top 200)", fontsize=13, fontweight="bold")
    ax.set_xlabel("Total Claims")
    ax.set_ylabel("Total Paid (Thousands Rs)")
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"Rs{x:.0f}K"))
    ax.legend()
    plt.tight_layout()
    save_chart(fig, "06_member_utilisation.png")


# -- Chart 7: Geographic Breakdown ----------------------------

def chart_geographic(conn):
    df = query(conn, """
        SELECT dm.state,
               ROUND(SUM(fc.paid_amount)/1000000, 2) AS paid_millions
        FROM fact_claim fc
        JOIN dim_member dm ON fc.member_id = dm.member_id
        GROUP BY dm.state
        ORDER BY paid_millions DESC
        LIMIT 15
    """)
    if df.empty:
        print("  [SKIP] No data for geographic chart")
        return

    fig, ax = plt.subplots(figsize=(12, 5))
    ax.bar(df["state"], df["paid_millions"], color="#64B5CD", alpha=0.85)
    ax.set_title("Total Paid Amount by State (Top 15)", fontsize=14, fontweight="bold")
    ax.set_xlabel("State")
    ax.set_ylabel("Total Paid (Millions Rs)")
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"Rs{x:.1f}M"))
    plt.xticks(rotation=45, ha="right")
    plt.tight_layout()
    save_chart(fig, "07_geographic.png")


# -- Chart 8: Payment Method Pie -------------------------------

def chart_payment_method(conn):
    df = query(conn, """
        SELECT payment_method, ROUND(SUM(paid_amount)/1000000, 2) AS paid_millions
        FROM fact_payment
        WHERE payment_method IS NOT NULL AND payment_method != 'nan'
        GROUP BY payment_method
        ORDER BY paid_millions DESC
    """)
    if df.empty:
        print("  [SKIP] No data for payment method chart")
        return

    fig, ax = plt.subplots(figsize=(7, 7))
    ax.pie(
        df["paid_millions"],
        labels=df["payment_method"],
        autopct="%1.1f%%",
        startangle=140,
        colors=["#4C72B0","#DD8452","#55A868","#C44E52","#8172B2"],
    )
    ax.set_title("Payment Method Breakdown (by Total Paid)", fontsize=13, fontweight="bold")
    plt.tight_layout()
    save_chart(fig, "08_payment_methods.png")


# -- Run all ---------------------------------------------------

def run_all_charts():
    print("\n" + "=" * 55)
    print("VISUALIZATIONS - Generating charts")
    print(f"  Database : {DB_PATH}")
    print(f"  Output   : {CHARTS_DIR}")
    print("=" * 55)

    if not DB_PATH.exists():
        print("\nERROR: Warehouse database not found.")
        print("  Please run step4_load_full.py first.")
        return

    conn = connect()

    chart_monthly_paid(conn)
    chart_top_procedures(conn)
    chart_provider_performance(conn)
    chart_denial_rate(conn)
    chart_los(conn)
    chart_member_utilisation(conn)
    chart_geographic(conn)
    chart_payment_method(conn)

    conn.close()

    print(f"\n[VIZ] All charts saved to: {CHARTS_DIR}")


if __name__ == "__main__":
    run_all_charts()
