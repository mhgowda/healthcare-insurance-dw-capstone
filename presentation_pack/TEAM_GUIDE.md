# Healthcare Insurance Claims – Data Warehouse & Analytics
### Complete Project Guide for Team Members

---

## What is this project?

A national health insurance company has data about lakhs of patients, doctors,
hospitals, and medical bills. This data is scattered across 7 separate CSV files
and is messy — it has missing values, wrong formats, fake IDs, and incorrect dates.

Our job is to:
1. **Understand the mess** — find all data quality problems
2. **Clean it** — remove the bad records
3. **Protect private information** — hide names, emails, phone numbers, dates of birth
4. **Build a Data Warehouse** — a well-organised database designed for analysis
5. **Answer business questions** — using SQL and Python
6. **Create charts** — to visualise the findings

---

## GitHub Repository

**https://github.com/mhgowda/healthcare-insurance-dw-capstone**

Clone it:
```
git clone https://github.com/mhgowda/healthcare-insurance-dw-capstone.git
```

> NOTE: The source CSV files are NOT in the repo (too large — 500MB).
> You need them locally in:  insurance_capstone_dataset_v1/data/merged/
> Get them from the team member who has the original dataset files.

---

## Source Data (7 CSV files, ~1.4 million rows total)

| File                    | What it contains                              | Rows     |
|-------------------------|-----------------------------------------------|----------|
| members_merged.csv      | Patient details — name, DOB, city, state      | 2,60,000 |
| providers_merged.csv    | Doctors and hospitals — name, specialty, city | 40,000   |
| claims_merged.csv       | Each insurance bill raised                    | 2,40,000 |
| claim_lines_merged.csv  | Individual services inside each bill          | 7,60,000 |
| payments_merged.csv     | Payments made on each bill                    | 1,00,000 |
| diagnoses_merged.csv    | ICD medical codes linked to bills             | 60,000   |
| procedures_merged.csv   | CPT procedure codes linked to bill lines      | 40,000   |

The data has intentional problems to make it realistic:
- Missing values       — blank names, empty emails
- Orphan FKs          — rows with __ORPHAN__ where a valid ID should be
- Negative amounts     — bills with negative money values
- Future dates         — birth dates in the future, future admit dates
- Duplicate IDs        — same claim number appearing twice
- Wrong data types     — text like "ERR" in a number column

---

## Project Structure

```
project_5/
|
+-- insurance_capstone_dataset_v1/
|   +-- data/merged/          <- 7 source CSV files (DO NOT MODIFY)
|   +-- logs/
|       +-- profiling_summary.json   <- auto-created by step1
|       +-- masking_audit.json       <- auto-created by step3
|
+-- capstone/                 <- YOUR ENTIRE PROJECT LIVES HERE
    +-- config.py             <- file paths and settings
    +-- requirements.txt      <- Python packages needed
    +-- README.md
    |
    +-- day1_sql/
    |   +-- 01_star_schema_ddl.sql     <- CREATE TABLE statements
    |   +-- 02_analytical_queries.sql  <- 10 business SQL queries
    |
    +-- day2_etl/
    |   +-- step1_profile.py           <- data quality checks
    |   +-- step2_clean.py             <- remove bad rows
    |   +-- step3_mask.py              <- hide private information
    |   +-- step4_load_full.py         <- build + load warehouse
    |   +-- step5_load_incremental.py  <- load only new rows next time
    |
    +-- day3_analytics/
    |   +-- analytics.py               <- 9 business questions answered
    |   +-- visualizations.py          <- 8 charts saved as PNG
    |   +-- analytics_notebook.ipynb   <- Jupyter notebook (for demo)
    |
    +-- output/
        +-- warehouse.db               <- SQLite database (auto-created)
        +-- profiling_report.csv       <- data quality findings
        +-- q1_summary.csv ... q9_     <- query result files
        +-- rejected_rows/             <- bad rows saved here
        +-- charts/                    <- 8 PNG chart images
```

---

## The Complete Flow — Step by Step

```
Step 1: Read CSVs -> Count problems (missing, negative, future, orphan, duplicate)
        |
        v
Step 2: Remove bad rows -> Save rejected rows to output/rejected_rows/
        |
        v
Step 3: Mask private info -> Replace names/emails/IDs with safe tokens
        |
        v
Step 4: Build warehouse -> Create star schema tables -> Load all clean data
        |
        v
Step 5: (Next run only) Load only new rows -> Using watermark/timestamp
        |
        v
Step 6: Run analytics -> Answer 9 business questions via Python + SQL
        |
        v
Step 7: Generate 8 charts -> Save to output/charts/
```

---

## What is a Star Schema?

It is the standard design for a data warehouse. You have:
- Fact tables      — store numbers (amounts, counts, days)
- Dimension tables — store descriptions (who, what, where, when)

```
           dim_date
               |
dim_member -- fact_claim -- dim_provider
               |
       fact_claim_line -- dim_procedure
               |
         fact_payment
               |
          dim_diagnosis
          dim_location
```

Our tables:

| Table           | Type      | Contains                             |
|-----------------|-----------|--------------------------------------|
| fact_claim      | Fact      | One row per insurance claim          |
| fact_claim_line | Fact      | One row per service line in a claim  |
| fact_payment    | Fact      | One row per payment made             |
| dim_member      | Dimension | Patient details (masked)             |
| dim_provider    | Dimension | Doctor/hospital details (masked)     |
| dim_date        | Dimension | Calendar breakdown for every date    |
| dim_diagnosis   | Dimension | Unique ICD diagnosis codes           |
| dim_procedure   | Dimension | Unique CPT procedure codes           |
| dim_location    | Dimension | Unique city + state combinations     |

---

## What is Data Masking?

Private information (PII/PHI) must never be visible to analysts.
Before loading into the warehouse we replace sensitive fields with safe tokens.

How it works:
- We use HMAC-SHA256 hashing to convert a value into a random-looking code
- The same input ALWAYS gives the same output — so JOINs between tables still work
- It CANNOT be reversed back to the original value

Example:
```
member_id : M0000001        ->  MBR3F9A2C1D4E
email     : alice@gmail.com ->  EMAIF29A3B1C2D
phone     : 9876543210      ->  PH8A2B3C4D
DOB       : 1985-03-10      ->  1985  (only year kept)
```

What is masked vs what is kept:

| Field                    | Action         | Reason                          |
|--------------------------|----------------|---------------------------------|
| member_id, provider_id   | Masked token   | Cannot expose real IDs          |
| first_name, last_name    | Masked token   | Personal information            |
| email, phone             | Masked token   | Personal information            |
| Date of birth            | Year only      | Age analysis still possible     |
| tax_id, license_no       | Masked token   | Sensitive provider info         |
| claim_number             | Masked token   | Cannot expose real claim ref    |
| city, state              | Kept as-is     | Needed for geographic analysis  |
| gender, specialty        | Kept as-is     | Needed for demographic analysis |
| Amounts, dates           | Kept as-is     | Core analytics data             |

---

## Data Quality Results

| Entity      | Total In | Valid (kept) | Rejected  | Main reasons                              |
|-------------|----------|--------------|-----------|-------------------------------------------|
| members     | 2,60,000 | 2,32,207     | 27,793    | Missing name, future DOB, duplicate ID    |
| providers   | 40,000   | 38,424       | 1,576     | Missing specialty, duplicate ID           |
| claims      | 2,40,000 | 2,08,640     | 31,360    | Orphan member_id, negative amount         |
| claim_lines | 7,60,000 | 6,45,733     | 1,14,267  | Orphan claim_number, negative units       |
| payments    | 1,00,000 | 59,615       | 40,385    | Orphan claim_number, non-numeric amount   |
| diagnoses   | 60,000   | 39,741       | 20,259    | Orphan claim_number, invalid ICD version  |
| procedures  | 40,000   | 24,844       | 15,156    | Orphan line_id, missing procedure code    |

---

## Business Questions Answered

| #  | Question                    | Answer from data                              |
|----|-----------------------------|-----------------------------------------------|
| Q1 | Total claims summary        | 2,08,640 claims — Rs 15.68 billion total paid |
| Q2 | Monthly paid trend          | ~Rs 540 million per month, consistent         |
| Q3 | Top procedures              | CPT35289 used 417 times — highest volume      |
| Q4 | Provider performance        | Top providers by paid amount + denial rate    |
| Q5 | Denial rate                 | 49.92% — nearly half of all claims denied     |
| Q6 | Length of stay distribution | Flat distribution 0 to 29 days                |
| Q7 | Top spending members        | By member_id, gender, state                   |
| Q8 | Geographic breakdown        | Punjab and Odisha have highest claims         |
| Q9 | Payment methods             | NEFT, UPI, RTGS, Cheque — nearly equal split  |

---

## Charts Generated (8 PNG files in output/charts/)

| File                        | What it shows                              |
|-----------------------------|--------------------------------------------|
| 01_monthly_paid.png         | Bar chart — monthly payout trend           |
| 02_top_procedures.png       | Horizontal bar — top 10 procedures         |
| 03_provider_performance.png | Side-by-side bars — paid + denial rate     |
| 04_denial_rate.png          | Bar chart — claims by denial code          |
| 05_los_distribution.png     | Bar chart — length of stay 0-30 days       |
| 06_member_utilisation.png   | Scatter plot — claims vs paid per member   |
| 07_geographic.png           | Bar chart — total paid by state            |
| 08_payment_methods.png      | Pie chart — payment method split           |

---

## Team Split (5 Members)

Member 1 — Data Profiling + Star Schema Design
  Files   : day1_sql/01_star_schema_ddl.sql, day2_etl/step1_profile.py
  Command : python capstone/day2_etl/step1_profile.py
  Present : Data quality findings + explain the star schema diagram

Member 2 — Data Cleaning + Error Logging
  Files   : day2_etl/step2_clean.py, output/rejected_rows/
  Command : python capstone/day2_etl/step2_clean.py
  Present : Rejection counts per entity + walk through cleaning rules

Member 3 — Data Masking + Full ETL Load
  Files   : day2_etl/step3_mask.py, day2_etl/step4_load_full.py
  Command : python capstone/day2_etl/step3_mask.py
            python capstone/day2_etl/step4_load_full.py
  Present : What was masked + how masking works + warehouse loading

Member 4 — Incremental ETL + Analytical SQL
  Files   : day2_etl/step5_load_incremental.py, day1_sql/02_analytical_queries.sql
  Command : python capstone/day2_etl/step5_load_incremental.py
  Present : Watermark concept + run 3 SQL queries live

Member 5 — Python Analytics + Charts + Notebook
  Files   : day3_analytics/analytics.py, visualizations.py, notebook.ipynb
  Command : python capstone/day3_analytics/analytics.py
            python capstone/day3_analytics/visualizations.py
  Present : Open notebook live + walk through each chart + optimization summary

---

## How to Run (from scratch)

1. Install packages (only needed once):
   pip install pandas matplotlib

2. Run in this exact order:
   python capstone/day2_etl/step1_profile.py
   python capstone/day2_etl/step2_clean.py
   python capstone/day2_etl/step3_mask.py
   python capstone/day2_etl/step4_load_full.py
   python capstone/day3_analytics/analytics.py
   python capstone/day3_analytics/visualizations.py

3. Open the notebook:
   pip install notebook
   jupyter notebook capstone/day3_analytics/analytics_notebook.ipynb

---

## Technologies Used

| Tool       | Purpose                                  |
|------------|------------------------------------------|
| Python     | Main programming language                |
| Pandas     | Data cleaning and transformation         |
| SQLite     | Local data warehouse (no setup needed)   |
| SQL        | CREATE TABLE, analytical queries         |
| Matplotlib | Charts and graphs                        |
| Jupyter    | Interactive notebook for presentation    |
