# Presentation Guide — Member 5
## Slides 12, 13, 14, 15, 16  |  Charts + Optimisation + Deliverables + Close
### Healthcare Insurance Claims DW & Analytics — Group 5

---

> **Your job:** Walk through all 8 charts, explain the optimisations,
> present the deliverables checklist, and close the presentation.
> Estimated time: 5-6 minutes

---

## Slide 12 — Visualisations Part 1
### (Monthly Paid | Top Procedures | Denial Rate | Length of Stay)

### What to say
"Part 1 of our charts shows four visualisations.

Top left — Monthly Paid Amount.
A bar chart showing total insurance payouts by month.
Key insight: Almost perfectly flat at about Rs 540 million per month
across 30 months. There is no seasonal pattern in the data.

Top right — Top 10 Procedures.
A horizontal bar chart showing the most frequently billed procedure codes.
Key insight: CPT35289 is the most used procedure at 417 claim lines.
The top 10 procedures together represent a significant share of all services.

Bottom left — Denial Rate by Code.
A bar chart showing claims distributed across denial codes.
The blue bar is NO DENIAL — paid claims.
The orange bars are denial codes D001, D002, D003.
Key insight: Denied and paid portions are almost exactly equal at 50% each.

Bottom right — Length of Stay Distribution.
How many claims had each length of stay from 0 to 29 days.
Key insight: The distribution is completely flat — no clustering.
This is because the data was synthetically generated with a uniform
random distribution, not from real hospital patterns."

---

## Slide 13 — Visualisations Part 2
### (Provider Performance | Geographic | Member Utilisation | Payment Methods)

### What to say
"Part 2 shows four more visualisations.

Top left — Provider Performance.
Two side-by-side bars for the top 15 providers:
Left bar = total paid, Right bar = denial rate percentage.
Key insight: No consistent relationship — some high-volume providers
have low denials, some have very high denials over 60%.

Top right — Paid Amount by State.
Top 15 states by total insurance spend.
Key insight: Punjab and Odisha lead significantly.
This analysis is possible ONLY because we kept city and state unmasked.
Masking city would have made geographic analysis impossible.

Bottom left — Member Utilisation Scatter Plot.
Each dot = one member. X axis = number of claims, Y axis = total paid.
Blue dots = male, orange triangles = female.
Key insight: High-spend members cluster at 6-7 claims.
Gender does not predict spend — both genders appear equally.

Bottom right — Payment Method Pie Chart.
Split of total paid across NEFT, UPI, RTGS, and Cheque.
Key insight: All four methods are approximately 25% each.
No single payment method dominates."

### Questions you may be asked
Q: How were the charts generated?
A: Using Matplotlib in day3_analytics/visualizations.py.
   Each chart connects to the SQLite warehouse, runs a SQL query via pandas,
   and saves the result as a PNG file in output/charts/.
   The same charts are also interactive in the Jupyter notebook.

Q: Why does the LOS distribution look flat?
A: The data was synthetically generated with a uniform random distribution
   0 to 29 days. Real hospital data would show clustering — most stays
   0-3 days with a long tail for complex cases.

Q: Why could you do geographic analysis if data was masked?
A: We specifically kept city and state unmasked because they do not identify
   individuals by themselves. Names, emails, phone, DOB — those were masked.
   City and state are needed for geographic analysis and are safe to keep.

---

## Slide 14 — Optimisation

### What to say
"We made three types of optimisation.

First — Database Indexes.
We created 7 indexes on foreign key columns and commonly filtered columns.
An index is like the index at the back of a textbook — instead of reading
every page to find a topic, you go straight to the right page.
Without indexes, SQLite scans every row for every JOIN.
With indexes on member_id, provider_id, admit_date_key, and is_denied,
lookups are O(log n) instead of O(n).

Second — Vectorised Pandas Operations.
Pandas processes entire columns at once using C code internally.
A Python for-loop over 760,000 claim_line rows would take several minutes.
pd.to_numeric(col, errors='coerce') on the same rows takes under one second.
We used vectorised operations for all type conversions, date parsing, and filtering.

Third — Incremental Loading.
Once the warehouse is built, we do not reload 1.4 million rows every day.
The watermark mechanism means each daily run processes only new rows.
Instead of 3-5 minutes, incremental runs take seconds.

The bottom note confirms that deterministic masking has zero negative impact
on analytics — JOINs still work because the same token always maps to the same value."

### Questions you may be asked
Q: What is the difference between a vectorised operation and a loop?
A: A loop processes one row at a time in Python's interpreter.
   A vectorised operation processes the entire column in one C function call.
   For 1 million rows, a loop might take 10 seconds.
   The vectorised equivalent takes 0.01 seconds. That is a 1000x speedup.

Q: How does an index work in SQLite?
A: SQLite builds a B-tree data structure sorted by the indexed column.
   Queries filtering on that column use the B-tree to find matching rows
   in O(log n) time instead of scanning every row in O(n) time.

---

## Slide 15 — Project Deliverables

### What to say
"This slide is our deliverables checklist — every output required by the
capstone project document. Every item is marked Complete.

Key deliverables:

SQL work: Star schema DDL with indexes, 10 analytical SQL queries.

ETL work: Data profiling JSON report, full load script, incremental load script,
masking audit log, 7 rejection audit CSV files.

Analytics: Python analytics script, 8 Matplotlib charts as PNG files,
Jupyter notebook for interactive use.

Final outputs: This presentation, a 20-page project report in Word format,
README and Team Guide on GitHub."

---

## Slide 16 — Thank You / Close

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
- This presentation and a 20-page project report

All code is available on our GitHub repository:
github.com/mhgowda/healthcare-insurance-dw-capstone

We are happy to answer any questions from the team."

---

## General Questions Any Member May Get

**Q: What would you do differently in production?**
A: Use Snowflake instead of SQLite, IDMC for visual ETL mapping,
   role-based masking with different levels for analysts vs auditors,
   connect Power BI or Tableau for business users, add automated tests.

**Q: How did you handle the __ORPHAN__ marker?**
A: During step2_clean.py, for every entity with a FK column, we check if
   the column equals '__ORPHAN__'. Any such row is rejected with reason
   'orphan_<column_name>'. We also check if the FK value exists in the
   parent entity's valid rows — if not, rejected as '<column>_not_found'.

**Q: What is the difference between OLTP and OLAP?**
A: OLTP (Online Transaction Processing) is the live billing system —
   optimised for fast inserts and updates.
   OLAP (Online Analytical Processing) is the warehouse — optimised for
   reading large volumes of historical data and calculating aggregates.

---

## Quick Reference — All Key Numbers

| Metric | Value |
|---|---|
| Total input rows | 14,00,000 |
| Valid rows loaded | 12,49,204 |
| Rejected rows | 2,50,796 (17.9%) |
| Warehouse tables | 9 (3 fact + 6 dim) |
| Valid claims | 2,08,640 |
| Total paid | Rs 15.68 Billion |
| Avg per claim | Rs 75,156 |
| Denial rate | 49.92% |
| Top state | Punjab (Rs 876M) |
| Charts | 8 PNG files |
| SQL queries | 10 |
| GitHub | github.com/mhgowda/healthcare-insurance-dw-capstone |
