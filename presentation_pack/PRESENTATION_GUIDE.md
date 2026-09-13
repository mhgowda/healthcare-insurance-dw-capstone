# Presentation Speaker Guide
## Healthcare Insurance Claims Data Warehouse & Analytics
### Slide-by-Slide Notes for Group 5

---

> **How to use this file:**
> Read each section before presenting that slide.
> The "What to say" part is a script you can speak naturally.
> The "Questions the guide might ask" part helps you prepare for evaluation.

---

## Slide 1 — Title Slide
**Title:** Healthcare Insurance Claims — Data Warehouse & Analytics

### What to say
"Good morning / Good afternoon everyone.
We are Group 5, and today we are presenting our Midterm Capstone Project on
Healthcare Insurance Claims Data Warehousing and Analytics.

This project covers three major areas:
First — SQL schema design and analytical querying.
Second — building a complete ETL pipeline with data masking.
Third — Python analytics and visualisation using Matplotlib.

We worked on approximately 1.4 million rows of insurance data across 7 source files.
Let us walk you through what we built."

### Key numbers to mention
- 1.4 million source rows
- 7 CSV files
- 9 warehouse tables
- 8 charts

---

## Slide 2 — Project Overview
**Title:** Project Overview

### What to say
"Let me start with the business problem.

A national health insurance company receives claims from members — that means
insured patients — and from providers — that means doctors and hospitals.
All this data comes as 7 separate CSV files and it is messy — it has missing
values, wrong formats, and broken references.

The company needs a Data Warehouse to answer questions like:
- How much money did we pay out each month?
- Which providers have the highest denial rates?
- Which states generate the most claims?
- Are there any fraud patterns?

AND they need all of this without exposing patient names, emails, phone numbers,
or dates of birth to the analysts.

Our solution has four parts:
1. Star Schema Design — how we organised the warehouse tables.
2. ETL Pipeline — how we moved data from the CSVs into the warehouse.
3. Data Masking — how we protected patient privacy.
4. Analytics and Charts — how we answered the business questions.

The technology we used is entirely local — Python, Pandas, SQLite, and Matplotlib.
No cloud, no Docker, no complex setup."

### Questions the guide might ask
Q: Why did you choose SQLite instead of PostgreSQL?
A: For a learning project, SQLite requires zero setup — it is a single file.
   The skills transfer directly to PostgreSQL or Snowflake in production.

Q: What does OLTP mean?
A: Online Transaction Processing. It is the live system where data is created
   in real time — like a hospital billing system. We took extracts (CSV files)
   from that system and built a Data Warehouse on top.

---

## Slide 3 — Source Data (7 CSV Files)
**Title:** Source Data — 7 CSV Files

### What to say
"We had 7 source CSV files totalling about 1.4 million rows.

Let me walk through each one:

- members_merged.csv — this has 2,60,000 rows. It contains patient details:
  their name, email, phone, date of birth, city, state, and whether they are active.

- providers_merged.csv — 40,000 rows. Doctors and hospitals with their
  specialty, location, tax ID, and licence number.

- claims_merged.csv — 2,40,000 rows. Each row is one insurance claim.
  It has the claim amount, dates, and whether it was denied.

- claim_lines_merged.csv — 7,60,000 rows. This is the largest file.
  Each claim is broken into individual services — like 'blood test' or 'X-ray'.

- payments_merged.csv — 1,00,000 rows. Payments made against claims.

- diagnoses_merged.csv — 60,000 rows. ICD medical codes attached to claims.
  ICD stands for International Classification of Diseases.

- procedures_merged.csv — 40,000 rows. CPT codes for specific medical procedures.
  CPT stands for Current Procedural Terminology.

Now, about 83 percent of these rows are clean.
The remaining 17 percent were intentionally made bad — to test our ETL pipeline.
The bad rows include things like blank names, negative billing amounts,
dates in the future, and broken references between tables."

### Questions the guide might ask
Q: What is the difference between claims and claim_lines?
A: A claim is the overall bill submitted by a provider for one visit or episode.
   A claim_line is one specific service inside that claim.
   For example: one hospital visit (claim) might include a blood test, an X-ray,
   and a consultation — each is a separate claim_line.

Q: What does ICD mean?
A: International Classification of Diseases. It is a global standard for
   classifying medical conditions. For example ICD10.01 might represent a
   specific type of fever. Every diagnosis gets an ICD code.

Q: What does CPT mean?
A: Current Procedural Terminology. It is a standard code for medical procedures.
   For example CPT35289 might represent a specific type of blood test.

---

## Slide 4 — Entity Relationship Diagram
**Title:** Entity Relationship Diagram

### What to say
"This slide shows the Entity Relationship Diagram — or ER diagram — of our
source system. This is the structure of the data BEFORE we transformed it
into the warehouse.

An ER diagram shows tables as boxes and relationships as lines between them.

The top row has three tables:
- MEMBERS on the left — these are the patients.
- CLAIMS in the middle — the insurance bills.
- PAYMENTS on the right — money paid for those bills.

You can see that MEMBERS connects to CLAIMS — one member can have many claims.
CLAIMS connects to PAYMENTS — one claim can have many payments.

The bottom row has four tables:
- PROVIDERS on the left — doctors and hospitals.
- CLAIM_LINES — the individual services inside each claim.
- DIAGNOSES — ICD codes attached to claims.
- PROCEDURES — CPT codes attached to claim lines.

The badges on each field show PK for Primary Key — the unique identifier —
and FK for Foreign Key — the column that links to another table.

For example, member_id appears in both MEMBERS (as PK) and CLAIMS (as FK).
That link is what gets broken when we have orphan foreign keys in our data."

### Questions the guide might ask
Q: What is a Primary Key?
A: A Primary Key is a column (or set of columns) that uniquely identifies
   each row in a table. No two rows can have the same Primary Key value.
   For example, every claim has a unique claim_number.

Q: What is a Foreign Key?
A: A Foreign Key is a column that refers to the Primary Key of another table.
   It creates a relationship between two tables. For example, member_id in
   the claims table is a Foreign Key that points to the members table.

Q: What is an orphan foreign key?
A: An orphan FK is a row where the FK column contains a value that does not
   exist in the parent table. In our data it appeared as the literal string
   '__ORPHAN__'. For example, a claim with member_id = '__ORPHAN__' means
   the claim exists but the patient record is missing. We reject all orphan rows.

---

## Slide 5 — Star Schema Design
**Title:** Star Schema — Data Warehouse Design

### What to say
"The star schema is the design we used for the Data Warehouse. It is different
from the ER diagram on the previous slide — that was how the raw data is structured.
This is how we REORGANISED it for fast analytics.

A star schema has two types of tables:

FACT tables — these store the numbers. The things you measure.
Amounts paid, number of claims, length of stay.

DIMENSION tables — these store the descriptions. The context.
Who was the member? Which provider? What date? What location?

We have three fact tables:
- fact_claim — one row per insurance claim
- fact_claim_line — one row per service inside a claim
- fact_payment — one row per payment made

And six dimension tables:
- dim_member — patient details (with PII masked)
- dim_provider — doctor/hospital details (with PII masked)
- dim_date — a full calendar breakdown for every date
- dim_diagnosis — unique ICD codes
- dim_procedure — unique CPT codes
- dim_location — city and state combinations

The connecting lines show foreign key relationships between the tables.

The reason we use this design is simple — analytical queries are very fast
and the SQL is easy to write. You just join the fact table to the dimension
you need."

### Questions the guide might ask
Q: Why is it called a 'star' schema?
A: Because when you draw it, the fact table is in the centre and the
   dimension tables surround it on all sides — like points of a star.

Q: What is dim_date and why do we need a separate table for dates?
A: dim_date is a calendar dimension. Every date from the earliest admit_date
   to the latest payment_date gets one row, with columns like year, quarter,
   month, month_name, day_name, and is_weekend.
   This lets us do GROUP BY year or GROUP BY quarter without any date
   arithmetic functions in our queries.

Q: What is the difference between a star schema and the original ER diagram?
A: The ER diagram shows the OLTP structure — designed for fast inserts and updates.
   The star schema is redesigned for analytics — designed for fast reads and
   aggregations. We denormalised some tables and separated concerns more cleanly.

---

## Slide 6 — ETL Pipeline Architecture
**Title:** ETL Pipeline Architecture

### What to say
"ETL stands for Extract, Transform, Load. It is the process of taking raw data,
cleaning and reshaping it, and loading it into the warehouse.

Our pipeline has 6 steps shown left to right on the slide:

Step 1 — Source CSVs. We start with 7 files, 1.4 million rows.

Step 2 — Profile. We read every file and count all quality problems.
This produces a JSON report so we know exactly what is wrong before touching anything.

Step 3 — Clean and Validate. We remove the bad rows — orphans, negatives,
future dates, missing values, type mismatches, duplicates.
Rejected rows are saved to audit files.

Step 4 — Mask PII. We hash the sensitive fields before loading.
Names, emails, IDs all become tokens. DOB becomes birth year only.

Step 5 — Full ETL Load. We build the schema and load everything.
This is the complete initial load.

Step 6 — Incremental Load. For future runs, we only load new rows
using a watermark mechanism. No need to reload everything every day.

The bottom half of the slide shows the difference:
Full Load drops all data and reloads everything — used for the first time or
when you need a complete refresh.
Incremental Load reads the last_load_dt timestamp and only processes rows
newer than that. Much faster for daily runs."

### Questions the guide might ask
Q: What is a watermark in incremental loading?
A: A watermark is a timestamp stored in a control table (etl_watermark).
   After each run, we save the current time for each entity.
   On the next run, we filter the CSV to rows where the date column
   is greater than that saved timestamp. Only new rows are processed.

Q: In what order do you load the tables and why?
A: Dimensions first, then facts. Because the fact tables have foreign keys
   that reference the dimension tables. If we loaded facts first, the FK
   references would point to rows that do not exist yet.
   For the same reason, we clean members before claims — so we can validate
   that every member_id in a claim actually exists.

Q: What happens to rejected rows?
A: They are saved to CSV files in output/rejected_rows/ with a column called
   rejection_reason that explains exactly why each row failed.
   This creates a full audit trail and allows the team to fix and resubmit
   those rows in a future run.

---

## Slide 7 — Data Quality Challenges
**Title:** Data Quality Challenges

### What to say
"This slide shows the six types of bad data we found in the source files.

Let me go through each one:

1. Missing Values — some required fields like first_name, email, and city
   were left blank. We reject any row that is missing a required field.

2. Orphan Foreign Keys — the string '__ORPHAN__' appeared in columns like
   member_id in claims. This means the claim references a member that does
   not exist. We reject all orphan rows.

3. Negative Amounts — financial columns like allowed_amount and paid_amount
   had negative values. A bill of minus 5000 rupees makes no business sense.
   We reject all negative amount rows.

4. Future Dates — some admit dates and dates of birth were set in the future.
   A patient cannot be admitted to a hospital in 2031. We reject these rows.

5. Duplicate Primary Keys — the same claim_number or member_id appeared more
   than once. We keep the first occurrence and reject all duplicates.

6. Type Mismatches — numeric columns like paid_amount contained text values
   like 'ERR' or 'MISSING'. SQL cannot add these up. We reject them.

The bottom of each card shows actual numbers from our profiling:
For example, we found 30,000 orphan member_ids in the claims file,
and 22,880 future dates of birth in the members file."

### Questions the guide might ask
Q: How did you detect type mismatches programmatically?
A: We used pandas pd.to_numeric(column, errors='coerce').
   This tries to convert every value to a number.
   Any value that cannot be converted becomes NaN (null).
   We then counted how many new nulls appeared — those were the type mismatches.

Q: Why are there so many orphan foreign keys?
A: They were deliberately injected into the dataset as bad rows.
   In real life, orphan FKs happen when records are deleted from the parent
   table without cascading the delete to child records.

Q: What percentage of rows were bad overall?
A: 17.9 percent — about 250,796 rows out of 1.4 million failed one or more checks.

---

## Slide 8 — Data Cleaning Results
**Title:** Data Cleaning Results

### What to say
"This table shows the final results of our cleaning step.

For each of the 7 entities, we show:
- Total input rows from the CSV
- How many passed validation
- How many were rejected
- The rejection percentage
- The main reasons for rejection

The key numbers to note:

members — 10.7% rejected — mainly missing names and future dates of birth.

providers — only 3.9% rejected — providers are relatively clean.

claims — 13.1% rejected — mainly orphan member_id and negative amounts.

claim_lines — 15.0% rejected — largest file, 1,14,267 bad rows.

payments — 40.4% rejected — the highest rejection rate. Almost all bad payments
had orphan claim_numbers or the text 'ERR' in the paid_amount column.

diagnoses — 33.8% rejected — orphan claim_numbers and invalid ICD version codes.

procedures — 37.9% rejected — orphan line_ids.

Overall total: 12,49,204 valid rows out of 14,00,000 input rows.
That is an 82.1% acceptance rate, or 17.9% rejection rate.

All rejected rows are saved in output/rejected_rows/ with a reason column."

### Questions the guide might ask
Q: Why is the payment rejection rate so high at 40.4%?
A: About 40% of payment rows were deliberately injected as bad data.
   The main issue was that paid_amount contained the string 'ERR'
   and claim_number contained '__ORPHAN__'. Both are type mismatch
   and orphan FK issues respectively.

Q: What did you do with the rejected rows?
A: We saved them to CSV files with a rejection_reason column.
   They are not lost — they can be fixed and reloaded in a future ETL run.

Q: Why do you clean entities in a specific order?
A: Because of dependency. Claims have a FK to members.
   If we cleaned claims first, we could not validate whether member_ids
   in claims actually exist in the members table.
   So we must clean members first, then use the valid member IDs to
   validate claims, and so on down the dependency chain.

---

## Slide 9 — Data Masking
**Title:** Data Masking — PII / PHI Protection

### What to say
"Data masking is one of the most important parts of this project.

PII stands for Personally Identifiable Information — names, emails, phone numbers.
PHI stands for Protected Health Information — medical diagnoses, treatment dates.

Healthcare data is among the most sensitive personal information in existence.
If an analyst needs to study claim costs and denial rates, they do NOT need to
know that patient 'Alice Smith' with email 'alice@gmail.com' made those claims.
They only need to know that member 'MBR3F9A2C1D4E' made those claims.

We use HMAC-SHA256 hashing with a secret salt string.

Why deterministic? Because the same input always gives the same output.
So when we mask member_id 'M0000001' to 'MBR3F9A2C1D4E' in dim_member,
that same token 'MBR3F9A2C1D4E' also appears in fact_claim.
This means JOIN operations between the two tables still work correctly.

Why irreversible? Because SHA-256 is a one-way function.
You cannot reverse 'MBR3F9A2C1D4E' back to 'M0000001' without the salt.

Date of birth is handled differently — we do not hash it.
We reduce it to birth year only.
So '1985-03-10' becomes just '1985'.
The analyst can still calculate age bands — 18-34, 35-50, 50+ — but
cannot identify the individual person.

Fields we kept unmasked:
city, state, gender, specialty, procedure codes, diagnosis codes, amounts.
These are needed for geographic and clinical analytics and do not identify individuals."

### Questions the guide might ask
Q: What is HMAC-SHA256?
A: HMAC stands for Hash-based Message Authentication Code.
   SHA-256 is the hashing algorithm used inside it.
   We feed the original value PLUS a secret salt string into the algorithm.
   It produces a fixed-length hexadecimal output.
   The salt makes it specific to our project — even if someone gets the
   masked data, they cannot reverse it without knowing our exact salt value.

Q: What is deterministic masking and why is it important?
A: Deterministic means the same input always gives the same output.
   This is essential because we need foreign key relationships to still work.
   If member_id 'M0000001' mapped to a different token each time it appeared,
   the JOIN between dim_member and fact_claim would break.

Q: Why not just delete the PII fields entirely?
A: Some PII fields ARE needed for analysis — city and state for geographic
   analysis, gender for demographic analysis.
   Deleting everything would remove analytical value.
   Masking keeps the relationship and allows grouping and joining
   while removing the ability to identify individuals.

---

## Slide 10 — Analytical SQL Queries
**Title:** Analytical SQL Queries

### What to say
"We built 10 analytical SQL queries to answer business questions from the warehouse.

Let me go through them:

Q1 — Total summary — how many claims, how much was paid, what is the denial rate?
Answer: 208,640 claims, Rs 15.68 billion paid, 49.92% denial rate.

Q2 — Monthly trend — how does paid amount change month by month?
Answer: Consistent about Rs 540 million per month over 30 months.

Q3 — Top procedures — which medical procedures are used most often?
Answer: CPT35289 with 417 uses is the most utilised procedure.

Q4 — Provider scorecard — which providers paid most and which had highest denials?
Answer: Some providers have over 60% denial rate — worth investigating for fraud.

Q5 — Denial breakdown — which denial codes appear most?
Answer: D001, D002, D003 each account for about 16.6% of all claims.
        50% of claims are paid with no denial.

Q6 — Length of stay — how long do patients stay in hospital?
Answer: Flat distribution from 0 to 29 days. Average is about 14-15 days.

Q7 — Member utilisation — who are the highest-cost members?
Answer: Top member spent Rs 785,828 across 6 claims.

Q8 — Geographic — which states generate the most claims?
Answer: Punjab leads with Rs 876 million, followed by Odisha.

Q9 — Payment methods — how do providers prefer to receive payment?
Answer: NEFT, UPI, RTGS, and Cheque are all roughly equal at about 25% each.

Q10 — Diagnosis cost — which ICD codes lead to the highest costs?
Answer: Links diagnosis codes to claim amounts for cost analysis.

All queries are stored in day1_sql/02_analytical_queries.sql and
also executed via Python in day3_analytics/analytics.py."

### Questions the guide might ask
Q: Walk me through one SQL query in detail.
A: Take Q4 — Provider Performance.
   We SELECT the provider_id, specialty, and state from dim_provider.
   We JOIN to fact_claim on provider_id.
   We GROUP BY each provider and calculate total claims, total paid,
   average paid, average length of stay, and the denial rate as a percentage.
   ORDER BY total_paid DESC LIMIT 20 gives us the top 20.

Q: Why is the denial rate 49.92%? Is that normal?
A: No, it is unusually high. In real-world insurance, typical denial rates
   are 10-20%. A 50% denial rate would be a major red flag — it could indicate
   systematic coding errors, policy mismatches, or potentially fraudulent claims.
   Our data has this because bad rows were injected with denial codes D001/D002/D003.

---

## Slide 11 — Key Business Metrics
**Title:** Key Business Metrics

### What to say
"This slide shows the six most important KPIs — Key Performance Indicators —
from our warehouse after the full ETL load.

Total Claims: 2,08,640 claims loaded after cleaning.

Total Paid: Rs 15.68 billion. That is the total insurance payout.

Average Paid per Claim: Rs 75,156. This is remarkably consistent across all claims.

Denial Rate: 49.92 percent — nearly half of all claims were denied.
This is a critical business insight.

Top State: Punjab with the highest total insurance expenditure.

Warehouse Tables: We have 9 tables total — 3 fact tables and 6 dimension tables.

These numbers come directly from the live warehouse.db file after the ETL
pipeline completed. They are real query results, not estimates."

### Questions the guide might ask
Q: How did you verify that these numbers are correct?
A: We ran the Q1 summary SQL query directly and checked it against
   the source CSV row counts. The numbers are consistent with the
   cleaning results — 208,640 valid claims from 240,000 input claims.

---

## Slide 12 — Visualisations Part 1
**Title:** Analytics Visualisations — Part 1 of 2

### What to say
"Part 1 of our charts shows four visualisations:

Top left — Monthly Paid Amount.
A bar chart showing total insurance payouts by month.
Key insight: The line is almost perfectly flat at about Rs 540 million
per month across 30 months. There is no seasonal pattern.

Top right — Top 10 Procedures.
A horizontal bar chart showing the most frequently billed procedure codes.
Key insight: CPT35289 is the most used procedure at 417 claim lines.
The top 10 procedures together represent a significant share of all services billed.

Bottom left — Denial Rate by Code.
A bar chart showing how claims are distributed across denial codes.
The blue bar is NO DENIAL — paid claims.
The orange/red bars are denial codes D001, D002, D003.
Key insight: The denied and paid portions are almost exactly equal at 50% each.

Bottom right — Length of Stay Distribution.
A bar chart showing how many claims had each length of stay from 0 to 29 days.
Key insight: The distribution is flat — there is no clustering around a particular
number of days. This suggests the data was generated uniformly, not from real patterns."

---

## Slide 13 — Visualisations Part 2
**Title:** Analytics Visualisations — Part 2 of 2

### What to say
"Part 2 shows four more visualisations:

Top left — Provider Performance.
Two side-by-side bars for each of the top 15 providers:
Left bar shows total paid amount, right bar shows denial rate percentage.
Key insight: There is no consistent relationship between high payment and
high denial rate — some high-volume providers have low denials.

Top right — Paid Amount by State.
A bar chart showing top 15 states by total insurance spend.
Key insight: Punjab and Odisha lead significantly.
This analysis is possible only because we kept city and state unmasked.

Bottom left — Member Utilisation Scatter Plot.
Each dot represents one member — X axis is number of claims, Y axis is total paid.
Blue dots are male members, orange triangles are female members.
Key insight: High-spend members cluster at 6-7 claims. Gender does not predict spend.

Bottom right — Payment Method Pie Chart.
Shows the split of total paid amount across NEFT, UPI, RTGS, and Cheque.
Key insight: All four methods are approximately equal at about 25% each.
No single payment method dominates."

### Questions the guide might ask
Q: How were the charts generated?
A: Using Matplotlib and Python in day3_analytics/visualizations.py.
   Each chart connects to the SQLite warehouse, runs a query via pandas,
   and saves the result as a PNG file in output/charts/.
   The same charts are also available interactively in the Jupyter notebook.

Q: Why does the LOS (Length of Stay) distribution look flat?
A: Because the data was synthetically generated with a uniform random distribution
   between 0 and 29 days. Real hospital data would show clustering — most stays
   would be 0-3 days with a long tail for complex cases.

---

## Slide 14 — Optimisation
**Title:** Optimisation

### What to say
"We made three types of optimisation in this project:

First — Database Indexes.
We created 7 indexes on the foreign key columns and commonly filtered columns.
An index is like the index at the back of a textbook — instead of reading every
page to find a topic, you go straight to the right page.
Without indexes, SQLite scans every row in the table for every JOIN.
With indexes, it goes directly to the matching rows.
The index on is_denied for example makes the denial rate query much faster.

Second — Vectorised Pandas Operations.
Pandas processes entire columns at once using optimised C code underneath.
A Python for-loop over 760,000 claim_line rows would take minutes.
pd.to_numeric(col, errors='coerce') on the same 760,000 rows takes under one second.
We used vectorised operations for all type conversions, date parsing, and filtering.

Third — Incremental Loading.
Once the warehouse is built, we do not need to reload all 1.4 million rows every day.
The watermark mechanism means each daily run only processes rows newer than
the last successful load. Instead of 3-5 minutes, incremental runs take seconds.

The masking impact note at the bottom shows that deterministic masking —
our choice of masking technique — has zero negative impact on analytics.
JOINs still work perfectly because the same token always maps to the same value."

### Questions the guide might ask
Q: What is the difference between a vectorised operation and a loop?
A: A loop in Python processes one item at a time in Python's interpreter.
   A vectorised operation processes the entire array in one C function call.
   For 1 million rows, a loop might take 10 seconds. The vectorised
   equivalent takes 0.01 seconds. That is a 1000x speedup.

Q: How does an index work in SQLite?
A: When you create an index on a column, SQLite builds a separate B-tree
   data structure sorted by that column. When a query filters or joins on
   that column, SQLite uses the B-tree to find matching rows in O(log n) time
   instead of scanning every row in O(n) time.

---

## Slide 15 — Project Deliverables
**Title:** Project Deliverables

### What to say
"This slide is our deliverables checklist — every output required by the
capstone project document, with the file location and status.

Every item is marked Complete.

The key deliverables to highlight:

For Day 1 SQL work:
- Star schema DDL — the CREATE TABLE statements with indexes
- Analytical SQL queries — 10 business questions

For Day 2 ETL work:
- Data profiling report saved as JSON in the logs folder
- Full ETL load script
- Incremental ETL load script
- Masking policy and audit log
- Error logs — 7 rejection CSV files

For Day 3 Analytics:
- Python analytics script
- 8 Matplotlib charts saved as PNG
- Jupyter notebook for interactive use

And finally this presentation, the project report,
the README for the GitHub repository, and the repository itself."

### Questions the guide might ask
Q: Where is the masking audit log?
A: logs/masking_audit.json. It records the timestamp, which fields were masked,
   and how many values were processed for each field.

Q: Where are the rejected rows stored?
A: output/rejected_rows/ — seven files, one per entity.
   Each file has the rejected rows with a rejection_reason column.

---

## Slide 16 — Thank You
**Title:** Thank You / Questions

### What to say
"That concludes our presentation on the Healthcare Insurance Claims
Data Warehouse and Analytics project.

To summarise what we built:
- A complete 9-table star schema warehouse
- A 5-step ETL pipeline with profiling, cleaning, masking, and loading
- Deterministic HMAC-SHA256 masking protecting all PII and PHI
- 10 analytical SQL queries
- 8 Matplotlib charts
- A Jupyter notebook for interactive analytics
- This presentation and a full 20-page project report

All code is available on our GitHub repository:
github.com/mhgowda/healthcare-insurance-dw-capstone

We are happy to answer any questions."

---

## General Questions to Prepare For

**Q: What is a Data Warehouse?**
A: A Data Warehouse is a database designed specifically for analytical queries
   rather than live transaction processing. It is optimised for reads and
   aggregations, uses denormalised schemas like the star schema, and typically
   stores historical data for trend analysis.

**Q: What is the difference between OLTP and OLAP?**
A: OLTP (Online Transaction Processing) is the live system — like a hospital
   billing system. It is optimised for fast inserts and updates, many small transactions.
   OLAP (Online Analytical Processing) is the warehouse — optimised for reading
   large volumes of historical data and calculating aggregates.

**Q: Why do you need 7 tables instead of one big table?**
A: Normalisation reduces redundancy. If we had one big table, the member's name
   would be repeated on every single claim row. With separate tables, member
   details are stored once and referenced by a key. This saves storage and ensures
   consistency — if a member changes their city, we update one row, not thousands.

**Q: What would you do differently if this were a real production project?**
A: We would use Snowflake instead of SQLite for scale, IDMC for visual ETL mapping,
   implement role-based masking with different levels for analysts and auditors,
   add automated tests with pytest, and connect a Power BI or Tableau dashboard
   for business users who do not write SQL.

**Q: How did you handle the __ORPHAN__ marker?**
A: During the cleaning step (step2_clean.py), for every entity that has a
   foreign key, we check if the FK column contains the string '__ORPHAN__'.
   Any row with an orphan FK is immediately rejected and saved to the
   rejection file with reason 'orphan_<column_name>'. We also check whether
   the FK value exists in the parent entity's valid rows — if it does not,
   we reject it with reason '<column_name>_not_found'.

---

## Quick Reference — Key Numbers

| Metric | Value |
|---|---|
| Source files | 7 CSVs |
| Total input rows | 14,00,000 |
| Valid rows loaded | 12,49,204 |
| Rejected rows | 2,50,796 |
| Rejection rate | 17.9% |
| Warehouse tables | 9 (3 fact + 6 dim) |
| Valid claims | 2,08,640 |
| Total paid | Rs 15.68 Billion |
| Avg per claim | Rs 75,156 |
| Denial rate | 49.92% |
| Top state | Punjab |
| Charts generated | 8 |
| SQL queries | 10 |
