-- ============================================================
-- DAY 1 – Analytical SQL Queries
-- Run these AFTER loading the warehouse (step4_load_full.py)
-- ============================================================

-- ── Q1. How many total claims do we have? ───────────────────
SELECT
    COUNT(*)                 AS total_claims,
    SUM(allowed_amount)      AS total_allowed,
    SUM(paid_amount)         AS total_paid,
    ROUND(AVG(paid_amount),2) AS avg_paid_per_claim
FROM fact_claim;


-- ── Q2. Monthly paid amount trend ───────────────────────────
-- Which months had the highest insurance payouts?
SELECT
    d.year,
    d.month,
    d.month_name,
    COUNT(fc.claim_number)   AS number_of_claims,
    ROUND(SUM(fc.paid_amount),2) AS total_paid
FROM fact_claim fc
JOIN dim_date d ON fc.admit_date_key = d.date_key
GROUP BY d.year, d.month, d.month_name
ORDER BY d.year, d.month;


-- ── Q3. Top 10 most expensive diagnosis codes ───────────────
-- Which ICD codes cost the most in total?
SELECT
    dx.diagnosis_code,
    dx.icd_version,
    COUNT(DISTINCT fcl.claim_number) AS claims_with_this_diagnosis,
    ROUND(SUM(fcl.line_amount), 2)   AS total_billed
FROM fact_claim_line fcl
JOIN dim_diagnosis dx ON fcl.procedure_code = dx.diagnosis_code
GROUP BY dx.diagnosis_code, dx.icd_version
ORDER BY total_billed DESC
LIMIT 10;


-- ── Q4. Top 10 procedures by how often they are used ────────
SELECT
    procedure_code,
    COUNT(*)           AS times_used,
    SUM(units)         AS total_units,
    ROUND(SUM(line_amount),2) AS total_billed
FROM fact_claim_line
GROUP BY procedure_code
ORDER BY times_used DESC
LIMIT 10;


-- ── Q5. Provider performance – who processes most claims? ───
SELECT
    dp.provider_id,
    dp.specialty,
    dp.state,
    COUNT(fc.claim_number)          AS total_claims,
    ROUND(SUM(fc.paid_amount),2)    AS total_paid,
    ROUND(AVG(fc.paid_amount),2)    AS avg_paid,
    ROUND(AVG(fc.length_of_stay),1) AS avg_length_of_stay,
    SUM(fc.is_denied)               AS denied_claims,
    ROUND(100.0 * SUM(fc.is_denied) / COUNT(*), 2) AS denial_rate_pct
FROM fact_claim fc
JOIN dim_provider dp ON fc.provider_id = dp.provider_id
GROUP BY dp.provider_id, dp.specialty, dp.state
ORDER BY total_paid DESC
LIMIT 20;


-- ── Q6. Denial rate by denial code ──────────────────────────
SELECT
    CASE
        WHEN denial_code IS NULL OR denial_code = '' THEN 'NO DENIAL'
        ELSE denial_code
    END                  AS denial_reason,
    COUNT(*)             AS claim_count,
    ROUND(100.0 * COUNT(*) / (SELECT COUNT(*) FROM fact_claim), 2) AS pct_of_total
FROM fact_claim
GROUP BY denial_reason
ORDER BY claim_count DESC;


-- ── Q7. Length of Stay distribution ─────────────────────────
SELECT
    length_of_stay       AS days_in_hospital,
    COUNT(*)             AS number_of_claims,
    ROUND(AVG(paid_amount),2) AS avg_paid
FROM fact_claim
WHERE length_of_stay >= 0
GROUP BY length_of_stay
ORDER BY length_of_stay;


-- ── Q8. Member utilisation – who uses insurance the most? ───
SELECT
    dm.member_id,
    dm.gender,
    dm.state,
    dm.birth_year,
    COUNT(fc.claim_number)       AS total_claims,
    ROUND(SUM(fc.paid_amount),2) AS total_paid
FROM fact_claim fc
JOIN dim_member dm ON fc.member_id = dm.member_id
GROUP BY dm.member_id, dm.gender, dm.state, dm.birth_year
ORDER BY total_paid DESC
LIMIT 20;


-- ── Q9. Geographic analysis – which state has most claims? ──
SELECT
    dm.state,
    COUNT(fc.claim_number)       AS total_claims,
    ROUND(SUM(fc.paid_amount),2) AS total_paid,
    ROUND(AVG(fc.paid_amount),2) AS avg_paid
FROM fact_claim fc
JOIN dim_member dm ON fc.member_id = dm.member_id
GROUP BY dm.state
ORDER BY total_paid DESC
LIMIT 15;


-- ── Q10. Payment method breakdown ───────────────────────────
SELECT
    payment_method,
    COUNT(*)                     AS payment_count,
    ROUND(SUM(paid_amount),2)    AS total_paid,
    ROUND(AVG(paid_amount),2)    AS avg_per_payment
FROM fact_payment
GROUP BY payment_method
ORDER BY total_paid DESC;
