# Presentation Guide — Member 2
## Slides 4, 5  |  ER Diagram + Star Schema Design
### Healthcare Insurance Claims DW & Analytics — Group 5

---

> **Your job:** Explain how the source data is structured (ER diagram) and
> how we redesigned it into the warehouse (star schema).
> Estimated time: 4-5 minutes

---

## Slide 4 — Entity Relationship Diagram

### What to say
"This slide shows the Entity Relationship Diagram — or ER diagram — of our
source system. This is the structure of the data BEFORE we transformed it
into the warehouse.

An ER diagram shows tables as boxes and relationships as lines between them.

The top row has three tables:
- MEMBERS on the left — these are the patients.
- CLAIMS in the middle — the insurance bills.
- PAYMENTS on the right — money paid for those bills.

MEMBERS connects to CLAIMS — one member can have many claims.
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

### Questions you may be asked
Q: What is a Primary Key?
A: A column that uniquely identifies each row in a table. No two rows can
   have the same Primary Key. Every claim has a unique claim_number.

Q: What is a Foreign Key?
A: A column that refers to the Primary Key of another table. It creates a
   relationship. member_id in claims is a FK pointing to the members table.

Q: What is an orphan foreign key?
A: A row where the FK column contains a value that does not exist in the
   parent table. In our data it appeared as the string '__ORPHAN__'.
   A claim with member_id = '__ORPHAN__' means the patient record is missing.
   We reject all orphan rows during cleaning.

---

## Slide 5 — Star Schema Design

### What to say
"The star schema is the design we used for the Data Warehouse. It is different
from the ER diagram on the previous slide — that was how the raw data is structured.
This is how we REORGANISED it for fast analytics.

A star schema has two types of tables:

FACT tables — store the numbers. The things you measure.
Amounts paid, number of claims, length of stay.

DIMENSION tables — store the descriptions. The context.
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

The reason we use this design: analytical queries are very fast and the SQL
is easy to write. You just join the fact table to the dimension you need."

### Questions you may be asked
Q: Why is it called a 'star' schema?
A: Because the fact table is in the centre and dimension tables surround it
   on all sides — like points of a star.

Q: What is dim_date and why do we need a separate table for dates?
A: dim_date is a calendar dimension. Every date gets one row with columns
   like year, quarter, month, month_name, day_name, and is_weekend.
   This lets us GROUP BY year or quarter without any date arithmetic functions.

Q: What is the difference between the star schema and the original ER diagram?
A: The ER diagram is the OLTP structure — designed for fast inserts and updates.
   The star schema is redesigned for analytics — fast reads and aggregations.
   We denormalised some tables and separated concerns more cleanly.

Q: What is OLTP vs OLAP?
A: OLTP (Online Transaction Processing) is the live billing system.
   OLAP (Online Analytical Processing) is the warehouse — optimised for
   reading large volumes of historical data and calculating aggregates.

---

## Your Key Numbers to Memorise

| Tables | Count |
|---|---|
| Fact tables | 3 (fact_claim, fact_claim_line, fact_payment) |
| Dimension tables | 6 (member, provider, date, diagnosis, procedure, location) |
| Total warehouse tables | 9 |
| Indexes created | 7 |

- After you finish Slide 5, hand over to **Member 3** for the ETL Pipeline.
