-- ============================================================
-- DAY 1 – Star Schema DDL
-- Healthcare Insurance Claims Data Warehouse
--
-- Run this in SQLite (used by step4_load_full.py) OR
-- copy into Snowflake / PostgreSQL with minor adjustments.
-- ============================================================

-- ── DIMENSION: dim_date ──────────────────────────────────────
-- Stores every calendar date we need, with useful breakdowns
CREATE TABLE IF NOT EXISTS dim_date (
    date_key      INTEGER PRIMARY KEY,  -- e.g. 20240115  (YYYYMMDD)
    full_date     TEXT    NOT NULL,     -- e.g. 2024-01-15
    year          INTEGER NOT NULL,
    quarter       INTEGER NOT NULL,     -- 1, 2, 3, or 4
    month         INTEGER NOT NULL,     -- 1 to 12
    month_name    TEXT    NOT NULL,     -- January, February …
    day_of_month  INTEGER NOT NULL,
    day_name      TEXT    NOT NULL,     -- Monday, Tuesday …
    is_weekend    INTEGER NOT NULL      -- 1 = weekend, 0 = weekday
);

-- ── DIMENSION: dim_member ────────────────────────────────────
-- Who has insurance? PII fields are MASKED before loading.
CREATE TABLE IF NOT EXISTS dim_member (
    member_id  TEXT PRIMARY KEY,   -- masked token (not real ID)
    gender     TEXT,
    city       TEXT,               -- kept for geographic analysis
    state      TEXT,               -- kept for geographic analysis
    birth_year INTEGER,            -- DOB reduced to year only (masked)
    is_active  INTEGER             -- 1 = active member
);

-- ── DIMENSION: dim_provider ─────────────────────────────────
-- Doctors, hospitals, clinics. tax_id and license_no are MASKED.
CREATE TABLE IF NOT EXISTS dim_provider (
    provider_id   TEXT PRIMARY KEY,  -- masked
    specialty     TEXT,
    city          TEXT,
    state         TEXT,
    tax_id        TEXT,              -- masked
    license_no    TEXT               -- masked
);

-- ── DIMENSION: dim_location ─────────────────────────────────
-- All unique city+state combinations (for geo analysis)
CREATE TABLE IF NOT EXISTS dim_location (
    location_key TEXT PRIMARY KEY,  -- "CITY|STATE"
    city         TEXT,
    state        TEXT
);

-- ── DIMENSION: dim_diagnosis ─────────────────────────────────
-- Unique diagnosis (ICD) codes
CREATE TABLE IF NOT EXISTS dim_diagnosis (
    diagnosis_code TEXT PRIMARY KEY,
    icd_version    TEXT   -- "ICD-9" or "ICD-10"
);

-- ── DIMENSION: dim_procedure ─────────────────────────────────
-- Unique procedure (CPT) codes
CREATE TABLE IF NOT EXISTS dim_procedure (
    procedure_code TEXT PRIMARY KEY
);

-- ── FACT: fact_claim ─────────────────────────────────────────
-- One row per insurance claim. The main fact table.
CREATE TABLE IF NOT EXISTS fact_claim (
    claim_number       TEXT PRIMARY KEY,  -- masked
    member_id          TEXT,              -- FK → dim_member
    provider_id        TEXT,              -- FK → dim_provider
    admit_date_key     INTEGER,           -- FK → dim_date
    discharge_date_key INTEGER,           -- FK → dim_date
    length_of_stay     INTEGER,
    allowed_amount     REAL,
    paid_amount        REAL,
    denial_code        TEXT,
    is_denied          INTEGER  -- 1 = denied, 0 = paid
);

-- ── FACT: fact_claim_line ────────────────────────────────────
-- One row per service line inside a claim
CREATE TABLE IF NOT EXISTS fact_claim_line (
    line_id          TEXT PRIMARY KEY,   -- masked
    claim_number     TEXT,               -- FK → fact_claim
    service_date_key INTEGER,            -- FK → dim_date
    procedure_code   TEXT,               -- FK → dim_procedure
    units            INTEGER,
    rate             REAL,
    line_amount      REAL
);

-- ── FACT: fact_payment ───────────────────────────────────────
-- One row per payment made on a claim
CREATE TABLE IF NOT EXISTS fact_payment (
    payment_id        TEXT PRIMARY KEY,  -- masked
    claim_number      TEXT,              -- FK → fact_claim
    payment_date_key  INTEGER,           -- FK → dim_date
    payment_method    TEXT,
    paid_amount       REAL,
    adjustment_code   TEXT
);

-- ── ETL CONTROL TABLE ────────────────────────────────────────
-- Tracks the last time each entity was loaded (for incremental ETL)
CREATE TABLE IF NOT EXISTS etl_watermark (
    entity       TEXT PRIMARY KEY,
    last_load_dt TEXT,     -- ISO datetime string
    rows_loaded  INTEGER
);

-- ── INDEXES (speeds up JOINs and filters) ───────────────────
CREATE INDEX IF NOT EXISTS idx_fc_member     ON fact_claim (member_id);
CREATE INDEX IF NOT EXISTS idx_fc_provider   ON fact_claim (provider_id);
CREATE INDEX IF NOT EXISTS idx_fc_admit      ON fact_claim (admit_date_key);
CREATE INDEX IF NOT EXISTS idx_fc_denied     ON fact_claim (is_denied);
CREATE INDEX IF NOT EXISTS idx_fcl_claim     ON fact_claim_line (claim_number);
CREATE INDEX IF NOT EXISTS idx_fcl_svcdate   ON fact_claim_line (service_date_key);
CREATE INDEX IF NOT EXISTS idx_fp_claim      ON fact_payment (claim_number);
CREATE INDEX IF NOT EXISTS idx_fp_date       ON fact_payment (payment_date_key);
