# Presentation Guide — Member 1
## Slides 1, 2, 3  |  Project Intro + Source Data
### Healthcare Insurance Claims DW & Analytics — Group 5

---

> **Your job:** Open the presentation, introduce the team, explain the
> business problem, and walk through the source data.
> Estimated time: 4-5 minutes

---

## Slide 1 — Title Slide

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
- 8 charts generated

---

## Slide 2 — Project Overview

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

### Questions you may be asked
Q: Why did you choose SQLite instead of PostgreSQL?
A: For a learning project, SQLite requires zero setup — it is a single file.
   The skills transfer directly to PostgreSQL or Snowflake in production.

Q: What does OLTP mean?
A: Online Transaction Processing. It is the live system where data is created
   in real time — like a hospital billing system. We took extracts (CSV files)
   from that system and built a Data Warehouse on top.

---

## Slide 3 — Source Data (7 CSV Files)

### What to say
"We had 7 source CSV files totalling about 1.4 million rows.

Let me walk through each one:

- members_merged.csv — 2,60,000 rows. Patient details:
  name, email, phone, date of birth, city, state, active status.

- providers_merged.csv — 40,000 rows. Doctors and hospitals:
  specialty, location, tax ID, licence number.

- claims_merged.csv — 2,40,000 rows. Each row is one insurance claim:
  the claim amount, dates, and whether it was denied.

- claim_lines_merged.csv — 7,60,000 rows. This is the largest file.
  Each claim is broken into individual services — like a blood test or X-ray.

- payments_merged.csv — 1,00,000 rows. Payments made against claims.

- diagnoses_merged.csv — 60,000 rows. ICD medical codes attached to claims.
  ICD stands for International Classification of Diseases.

- procedures_merged.csv — 40,000 rows. CPT codes for medical procedures.
  CPT stands for Current Procedural Terminology.

About 83 percent of these rows are clean.
The remaining 17 percent were intentionally made bad to test our ETL pipeline.
Bad rows include blank names, negative billing amounts, future dates, and
broken references between tables."

### Questions you may be asked
Q: What is the difference between claims and claim_lines?
A: A claim is the overall bill for one visit. A claim_line is one specific
   service inside that claim. One hospital visit (claim) might include a blood
   test, an X-ray, and a consultation — each is a separate claim_line.

Q: What does ICD mean?
A: International Classification of Diseases. A global standard for classifying
   medical conditions. For example ICD10.01 might represent a specific type of fever.

Q: What does CPT mean?
A: Current Procedural Terminology. A standard code for medical procedures.
   For example CPT35289 might represent a specific type of blood test.

Q: What is a Data Warehouse?
A: A database designed for analytical queries rather than live transaction
   processing. Optimised for reads and aggregations, stores historical data
   for trend analysis.

---

## Your Key Numbers to Memorise

| File | Rows |
|---|---|
| members | 2,60,000 |
| providers | 40,000 |
| claims | 2,40,000 |
| claim_lines | 7,60,000 (largest) |
| payments | 1,00,000 |
| diagnoses | 60,000 |
| procedures | 40,000 |
| **TOTAL** | **~14,00,000** |

- 83% clean rows, 17% bad rows
- After you finish Slide 3, hand over to **Member 2** for the ER Diagram.
