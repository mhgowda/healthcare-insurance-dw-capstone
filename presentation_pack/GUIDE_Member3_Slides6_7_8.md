# Presentation Guide — Member 3
## Slides 6, 7, 8  |  ETL Pipeline + Data Quality + Cleaning Results
### Healthcare Insurance Claims DW & Analytics — Group 5

---

> **Your job:** Explain the full ETL pipeline, all 6 data quality issues,
> and present the cleaning results with exact numbers.
> Estimated time: 5-6 minutes

---

## Slide 6 — ETL Pipeline Architecture

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
Full Load drops all data and reloads everything.
Incremental Load reads the last_load_dt timestamp and only processes rows
newer than that. Much faster for daily runs."

### Questions you may be asked
Q: What is a watermark in incremental loading?
A: A timestamp stored in a control table called etl_watermark.
   After each run, we save the current time for each entity.
   On the next run, we filter the CSV to rows where the date column
   is greater than that saved timestamp. Only new rows are processed.

Q: In what order do you load the tables and why?
A: Dimensions first, then facts. Fact tables have foreign keys referencing
   dimension tables. If we loaded facts first, the FKs would point to
   rows that do not exist yet.
   We also clean members before claims so we can validate every member_id.

Q: What happens to rejected rows?
A: Saved to CSV files in output/rejected_rows/ with a rejection_reason column.
   This creates a full audit trail. They can be fixed and reloaded later.

---

## Slide 7 — Data Quality Challenges

### What to say
"This slide shows the six types of bad data we found in the source files.

1. Missing Values — required fields like first_name, email, and city were blank.
   We reject any row that is missing a required field.

2. Orphan Foreign Keys — the string '__ORPHAN__' appeared in member_id in claims.
   This means the claim references a member that does not exist.
   We reject all orphan rows.

3. Negative Amounts — financial columns had negative values.
   A bill of minus 5000 rupees makes no business sense.
   We reject all negative amount rows.

4. Future Dates — some admit dates and dates of birth were set in the future.
   A patient cannot be admitted to a hospital in 2031.
   We reject these rows.

5. Duplicate Primary Keys — the same claim_number appeared more than once.
   We keep the first occurrence and reject all duplicates.

6. Type Mismatches — numeric columns like paid_amount had text values like 'ERR'.
   SQL cannot add 'ERR' to calculate totals. We reject them.

The bottom of each card shows actual numbers:
- 30,000 orphan member_ids in the claims file
- 22,880 future dates of birth in the members file"

### Questions you may be asked
Q: How did you detect type mismatches programmatically?
A: We used pandas pd.to_numeric(column, errors='coerce').
   Values that cannot convert to a number become NaN (null).
   We counted how many new nulls appeared — those were the type mismatches.

Q: Why are there so many orphan foreign keys?
A: They were deliberately injected as bad rows.
   In real life, orphan FKs happen when parent records are deleted without
   cascading the delete to child records.

Q: What percentage of rows were bad overall?
A: 17.9 percent — about 2,50,796 rows out of 14,00,000 failed one or more checks.

---

## Slide 8 — Data Cleaning Results

### What to say
"This table shows the final results of our cleaning step.

For each of the 7 entities we show total input, valid rows, rejected rows,
rejection percentage, and the main reasons.

Key numbers:
- members: 10.7% rejected — missing names and future DOBs.
- providers: only 3.9% rejected — relatively clean.
- claims: 13.1% rejected — orphan member_id and negative amounts.
- claim_lines: 15.0% rejected — largest file, 1,14,267 bad rows.
- payments: 40.4% rejected — the highest rate. Almost all had orphan
  claim_numbers or 'ERR' in the paid_amount column.
- diagnoses: 33.8% rejected — orphan claim_numbers and invalid ICD versions.
- procedures: 37.9% rejected — orphan line_ids.

Overall: 12,49,204 valid rows out of 14,00,000. That is 82.1% acceptance.
All rejected rows are saved in output/rejected_rows/ with a reason column."

### Questions you may be asked
Q: Why is the payment rejection rate so high at 40.4%?
A: About 40% of payment rows were deliberately injected as bad data.
   paid_amount contained 'ERR' and claim_number contained '__ORPHAN__'.

Q: What did you do with the rejected rows?
A: Saved to CSV files with a rejection_reason column. They can be fixed
   and reloaded in a future ETL run.

Q: Why clean entities in a specific order?
A: Dependency. Claims have a FK to members. We must clean members first
   so we can validate member_ids in claims. Same chain for all entities.

---

## Your Key Numbers to Memorise

| Entity | Rejected | Rate |
|---|---|---|
| members | 27,793 | 10.7% |
| providers | 1,576 | 3.9% |
| claims | 31,360 | 13.1% |
| claim_lines | 1,14,267 | 15.0% |
| payments | 40,385 | 40.4% |
| diagnoses | 20,259 | 33.8% |
| procedures | 15,156 | 37.9% |
| **TOTAL** | **2,50,796** | **17.9%** |

- After you finish Slide 8, hand over to **Member 4** for Data Masking.
