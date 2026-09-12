# Healthcare Insurance Claims – Data Warehouse & Analytics
### Midterm Capstone Project | 3-Day Sprint

---

## Project Overview

A national health insurance company collects data about members, providers,
claims, payments, diagnoses, and procedures across ~1.4 million records.

This project builds an end-to-end data warehouse solution:

- **Day 1** – Star schema design + analytical SQL
- **Day 2** – ETL pipeline (full + incremental) with data masking
- **Day 3** – Python analytics + Matplotlib visualizations

---

## Tech Stack

| Tool | Purpose |
|---|---|
| Python 3.10+ | ETL, analytics, visualizations |
| Pandas | Data cleaning, transformation |
| SQLite | Local data warehouse (no setup needed) |
| SQL | DDL, analytical queries |
| Matplotlib | Charts and graphs |
| Jupyter Notebook | Interactive analytics |

---

## Project Structure

```
capstone/
├── config.py                        Central config (paths, salt)
├── requirements.txt
│
├── day1_sql/
│   ├── 01_star_schema_ddl.sql       Star schema CREATE TABLE statements
│   └── 02_analytical_queries.sql    10 business SQL queries
│
├── day2_etl/
│   ├── step1_profile.py             Data quality profiling
│   ├── step2_clean.py               Validation + error logging
│   ├── step3_mask.py                PII/PHI masking
│   ├── step4_load_full.py           Full ETL load
│   └── step5_load_incremental.py    Incremental ETL load
│
├── day3_analytics/
│   ├── analytics.py                 9 analytical queries via Python
│   ├── visualizations.py            8 Matplotlib charts
│   └── analytics_notebook.ipynb    Jupyter notebook
│
└── output/
    ├── charts/                      8 PNG chart files
    ├── rejected_rows/               Bad rows removed during cleaning
    └── *.csv                        Query result files
```

---

## How to Run

```bash
# 1. Install dependencies
pip install pandas matplotlib

# 2. Profile the data
python capstone/day2_etl/step1_profile.py

# 3. Clean the data
python capstone/day2_etl/step2_clean.py

# 4. Mask PII
python capstone/day2_etl/step3_mask.py

# 5. Full ETL load
python capstone/day2_etl/step4_load_full.py

# 6. Run analytics
python capstone/day3_analytics/analytics.py

# 7. Generate charts
python capstone/day3_analytics/visualizations.py

# 8. Open notebook
jupyter notebook capstone/day3_analytics/analytics_notebook.ipynb
```

---

## Star Schema

```
           dim_date
               |
dim_member - fact_claim - dim_provider
               |
        fact_claim_line - dim_procedure
               |
         fact_payment
               |
          dim_diagnosis
          dim_location
```

---

## Data Quality Summary

| Entity | Total Rows | Valid | Rejected |
|---|---|---|---|
| members | 260,000 | 232,207 | 27,793 |
| providers | 40,000 | 38,424 | 1,576 |
| claims | 240,000 | 208,640 | 31,360 |
| claim_lines | 760,000 | 645,733 | 114,267 |
| payments | 100,000 | 59,615 | 40,385 |
| diagnoses | 60,000 | 39,741 | 20,259 |
| procedures | 40,000 | 24,844 | 15,156 |

> Note: Source CSV files are not included in this repo due to size (~500MB).
> Run `generate_insurance_capstone_dataset.py` to regenerate them.
