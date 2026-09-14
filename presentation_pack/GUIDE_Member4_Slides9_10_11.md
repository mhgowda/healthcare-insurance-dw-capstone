# Presentation Guide — Member 4
## Slides 9, 10, 11  |  Data Masking + SQL Queries + KPI Results
### Healthcare Insurance Claims DW & Analytics — Group 5

---

> **Your job:** Explain data masking (privacy protection), walk through
> the 10 SQL queries, and present the live KPI results from the warehouse.
> Estimated time: 5-6 minutes

---

## Slide 9 — Data Masking

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
When we mask member_id 'M0000001' to 'MBR3F9A2C1D4E' in dim_member,
that same token 'MBR3F9A2C1D4E' also appears in fact_claim.
This means JOIN operations between the two tables still work correctly.

Why irreversible? Because SHA-256 is a one-way function.
You cannot reverse 'MBR3F9A2C1D4E' back to 'M0000001' without the salt.

Date of birth is handled differently — we do not hash it.
We reduce it to birth year only.
'1985-03-10' becomes just '1985'.
The analyst can still calculate age bands but cannot identify the individual.

Fields kept unmasked: city, state, gender, specialty, procedure codes,
diagnosis codes, amounts. These are needed for analytics and do not
identify individuals."

### Questions you may be asked
Q: What is HMAC-SHA256?
A: HMAC = Hash-based Message Authentication Code. SHA-256 is the hashing
   algorithm. We feed the original value PLUS a secret salt string into it.
   It produces a fixed-length hexadecimal output.
   Without the salt, an attacker cannot reverse the masked data.

Q: What is deterministic masking and why is it important?
A: Deterministic means the same input always gives the same output.
   This is essential so foreign key relationships still work after masking.
   If member_id 'M0000001' mapped to a different token each time,
   the JOIN between dim_member and fact_claim would break.

Q: Why not just delete the PII fields entirely?
A: Some fields ARE needed — city and state for geographic analysis,
   gender for demographic analysis. Deleting everything removes analytical value.
   Masking keeps the relationship while removing the ability to identify people.

---

## Slide 10 — Analytical SQL Queries

### What to say
"We built 10 analytical SQL queries to answer business questions from the warehouse.

Q1 — Total summary: 208,640 claims, Rs 15.68 billion paid, 49.92% denial rate.

Q2 — Monthly trend: Consistent about Rs 540 million per month over 30 months.

Q3 — Top procedures: CPT35289 with 417 uses is the most utilised procedure.

Q4 — Provider scorecard: Some providers have over 60% denial rate —
     worth investigating for potential fraud.

Q5 — Denial breakdown: D001, D002, D003 each account for about 16.6% of claims.
     50% of claims are paid with no denial.

Q6 — Length of stay: Flat distribution 0 to 29 days. Average ~14-15 days.

Q7 — Member utilisation: Top member spent Rs 785,828 across 6 claims.

Q8 — Geographic: Punjab leads with Rs 876 million, followed by Odisha.

Q9 — Payment methods: NEFT, UPI, RTGS, Cheque all roughly equal at ~25% each.

Q10 — Diagnosis cost: Links ICD codes to claim amounts for cost analysis.

All queries are stored in day1_sql/02_analytical_queries.sql and
also executed via Python in day3_analytics/analytics.py."

### Questions you may be asked
Q: Walk me through one SQL query in detail.
A: Q4 — Provider Performance:
   SELECT provider_id, specialty, state from dim_provider.
   JOIN to fact_claim on provider_id.
   GROUP BY each provider and calculate total claims, total paid,
   average paid, average length of stay, and denial rate as a percentage.
   ORDER BY total_paid DESC LIMIT 20 gives the top 20 providers.

Q: Why is the denial rate 49.92%? Is that normal?
A: No, it is unusually high. Real-world insurance denial rates are 10-20%.
   50% denial would be a major red flag — systematic coding errors,
   policy mismatches, or potentially fraudulent claims.
   Our dataset has this because bad rows were injected with D001/D002/D003.

---

## Slide 11 — Key Business Metrics

### What to say
"This slide shows the eight most important KPIs from our warehouse after
the full ETL load. These are real query results from the live database.

Total Claims: 2,08,640 claims loaded after cleaning.

Total Paid: Rs 15.68 billion — the total insurance payout.

Average Paid per Claim: Rs 75,156 — remarkably consistent across all claims.

Denial Rate: 49.92 percent — nearly half of all claims were denied.
This is a critical business insight that warrants immediate investigation.

Top State: Punjab with the highest total insurance expenditure.

Top Procedure: CPT35289 — most utilised procedure code.

Payment Methods: 4 types — NEFT, UPI, RTGS, Cheque — all equal share.

Warehouse Tables: 9 total — 3 fact tables and 6 dimension tables."

### Questions you may be asked
Q: How did you verify these numbers are correct?
A: We ran the Q1 summary SQL query directly and cross-checked it against
   source CSV row counts. 208,640 is consistent with cleaning results —
   208,640 valid claims from 240,000 input claims.

Q: What does the 49.92% denial rate mean for the business?
A: It means almost half of all claims are being rejected. In real insurance
   operations this would trigger an audit — either providers are submitting
   incorrect claim codes or there is a systematic policy issue.

---

## Your Key Numbers to Memorise

| Metric | Value |
|---|---|
| Valid claims loaded | 2,08,640 |
| Total paid amount | Rs 15.68 Billion |
| Average paid per claim | Rs 75,156 |
| Denial rate | 49.92% |
| Top state | Punjab (Rs 876M) |
| Top procedure | CPT35289 (417 uses) |

- After you finish Slide 11, hand over to **Member 5** for charts and wrap-up.
