# Healthcare Insurance Claims – Data Warehouse & Analytics
## Beginner-Friendly Project (Python + SQLite)

---

## What is this project about?

A national health insurance company collects data about:
- **Members** – people who have insurance
- **Providers** – doctors and hospitals
- **Claims** – when a member visits a doctor and the insurer is billed
- **Claim Lines** – individual services inside a claim (e.g., one blood test, one X-ray)
- **Payments** – money paid by the insurer for a claim
- **Diagnoses** – medical codes (ICD codes) saying what illness was treated
- **Procedures** – medical procedure codes (CPT codes) for what was done

The data comes in 7 CSV files with ~1.4 million rows total.
Some rows are intentionally bad (missing values, wrong formats, fake IDs) — like real-world data.

**Our job:**
1. Profile the data — understand what's clean and what's not
2. Clean the data — remove/flag bad rows
3. Build a Data Warehouse (Star Schema) — organised for analysis
4. Mask private information (names, emails, phone, DOB etc.)
5. Load data into the warehouse
6. Run SQL analytics and Python/Matplotlib charts

---

## Star Schema Explained (Simple Version)

```
            dim_date
               |
dim_member — fact_claim — dim_provider
               |
         fact_claim_line — dim_procedure
               |
         fact_payment
               |
         dim_diagnosis
         dim_location
```

- **Fact tables** hold the numbers (amounts, counts, durations)
- **Dimension tables** hold the descriptions (who, what, where, when)

---

## Project Structure

```
capstone/
├── README.md                  ← You are here
├── requirements.txt           ← Python packages needed
│
├── day1_sql/
│   ├── 01_star_schema_ddl.sql     ← CREATE TABLE statements
│   └── 02_analytical_queries.sql  ← Business questions answered in SQL
│
├── day2_etl/
│   ├── step1_profile.py       ← Check data quality
│   ├── step2_clean.py         ← Remove bad rows
│   ├── step3_mask.py          ← Hide private information
│   ├── step4_load_full.py     ← Load everything into the warehouse
│   └── step5_load_incremental.py ← Load only new data next time
│
├── day3_analytics/
│   ├── analytics.py           ← Python analysis using Pandas + SQL
│   └── visualizations.py     ← Matplotlib charts
│
├── output/
│   ├── warehouse.db           ← SQLite database (auto-created)
│   ├── profiling_report.csv   ← Data quality findings
│   ├── rejected_rows/         ← Bad rows removed during cleaning
│   ├── masking_audit.csv      ← Log of what was masked
│   └── charts/                ← PNG chart images
```

---

## How to Run

### Step 1 – Install packages
```
pip install pandas matplotlib
```

### Step 2 – Profile the data (understand quality issues)
```
python day2_etl/step1_profile.py
```

### Step 3 – Clean the data
```
python day2_etl/step2_clean.py
```

### Step 4 – Mask PII
```
python day2_etl/step3_mask.py
```

### Step 5 – Full ETL load (first time)
```
python day2_etl/step4_load_full.py
```

### Step 6 – Incremental load (subsequent runs)
```
python day2_etl/step5_load_incremental.py
```

### Step 7 – Analytics + Charts
```
python day3_analytics/analytics.py
python day3_analytics/visualizations.py
```

---

## Technologies Used

| Tool       | What it does in this project                        |
|------------|-----------------------------------------------------|
| Python     | Main programming language                           |
| Pandas     | Read CSVs, clean and transform data                 |
| SQLite     | Local database for the warehouse (no setup needed)  |
| SQL        | Create tables, query data, analytical reports       |
| Matplotlib | Draw charts and graphs                              |

---

## Snowflake Note

If your trainer asks you to use Snowflake instead of SQLite:
- The DDL SQL in `day1_sql/01_star_schema_ddl.sql` works in Snowflake with minor changes
- Replace `CREATE TABLE IF NOT EXISTS` with Snowflake syntax
- Use Snowflake's web UI (Snowsight) to run the SQL
- Load CSVs via Snowflake stages or the Python `snowflake-connector-python` package
