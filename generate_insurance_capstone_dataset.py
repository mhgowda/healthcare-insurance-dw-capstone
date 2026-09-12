#!/usr/bin/env python3
"""
National Health Insurer Capstone Dataset Generator (Masking Variation Ready) — v1.0 FAST
- Deterministic totals: ~1.4M merged rows (1.2M clean + 0.2M imperfect)
- Vectorized generation (no per-row Faker in hot loops)
- Indian locale Faker pools (names, cities, providers)
- Bad-data injection (missing, negatives, future dates, dup keys, orphan FKs, type mismatches)
- Shuffle after concatenating clean+bad per entity, to avoid "tail-trimming" easy-clean artifacts
- Optional Parquet output, and a --quick scaler for fast smoke tests

Usage:
  python generate_healthcare_capstone_dataset.py --out ./healthcare_capstone_v1 --seed 23
  python generate_healthcare_capstone_dataset.py --out ./healthcare_capstone_v1_quick --seed 23 --quick 0.01
  python generate_healthcare_capstone_dataset.py --out ./healthcare_capstone_v1 --seed 23 --parquet
"""

from __future__ import annotations
import argparse, json, hashlib
from pathlib import Path
from datetime import datetime, timedelta
import numpy as np
import pandas as pd
from faker import Faker
from sklearn.utils import shuffle as sk_shuffle

# ------------------------- Helpers -------------------------

def build_args():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", required=True)
    ap.add_argument("--seed", type=int, default=23)
    ap.add_argument("--quick", type=float, default=1.0, help="Scale factor (0<quick<=1]. e.g., 0.01 for 1% rows")
    ap.add_argument("--parquet", action="store_true")
    return ap.parse_args()

def ensure_dirs(root: Path):
    (root / "data" / "merged").mkdir(parents=True, exist_ok=True)
    (root / "logs").mkdir(parents=True, exist_ok=True)

def build_faker(seed: int):
    Faker.seed(seed)
    np.random.seed(seed)
    return Faker("en_IN")

def rand_dates(rng, n, start: datetime, end: datetime):
    delta = int((end - start).total_seconds())
    offs = rng.integers(0, max(1,delta+1), size=n, dtype=np.int64)
    # return as pandas datetime64[ns]
    return (start + pd.to_timedelta(offs, unit="s")).astype("datetime64[ns]")

def sample_from_pool(rng, pool, size):
    idx = rng.integers(0, len(pool), size=size)
    return np.array(pool, dtype=object)[idx]

def make_email(firsts, lasts, rng):
    domains = np.array(["gmail.com","yahoo.co.in","outlook.com","rediffmail.com"], dtype=object)
    dn = domains[rng.integers(0, len(domains), size=len(firsts))]
    num = rng.integers(10, 9999, size=len(firsts))
    base = np.char.lower(firsts.astype(str)) + "." + np.char.lower(lasts.astype(str)) + num.astype(str)
    return base + "@" + dn

def write_merged(df_clean, df_bad, path: Path, seed: int, parquet=False):
    merged = sk_shuffle(pd.concat([df_clean, df_bad], ignore_index=True), random_state=seed).reset_index(drop=True)
    merged.to_csv(path, index=False)
    if parquet:
        merged.to_parquet(path.with_suffix(".parquet"), index=False)
    return len(merged)

def tokenized(series: pd.Series, salt: str = "hc_mask_salt") -> pd.Series:
    """Deterministic tokenization suitable for joins (not used in raw, but handy if you want a masked preview)."""
    # Returns hex-like stable tokens same length per value
    def tok(x):
        h = hashlib.sha256((str(x)+salt).encode("utf-8")).hexdigest()
        return h[:16].upper()
    return series.astype(str).map(tok)

# ------------------------- Main -------------------------

def main():
    args = build_args()
    if not (0 < args.quick <= 1.0):
        raise SystemExit("--quick must be in (0,1]. Example: 0.01 for 1%.")
    root = Path(args.out).resolve()
    ensure_dirs(root)
    rng = np.random.default_rng(args.seed)
    fake = build_faker(args.seed)

    data_dir = root / "data" / "merged"
    logs_dir = root / "logs"

    # ---------- TARGET COUNTS (clean, bad) ----------
    COUNTS = {
        "members":       (230_000, 30_000),
        "providers":     (35_000,  5_000),
        "claims":        (210_000, 30_000),
        "claim_lines":   (650_000, 110_000),
        "payments":      (60_000,  40_000),
        "diagnoses":     (40_000,  20_000),
        "procedures":    (25_000,  15_000),
    }
    # scale by quick
    COUNTS = {k: (max(1,int(v[0]*args.quick)), max(1,int(v[1]*args.quick))) for k,v in COUNTS.items()}

    # ---------- SMALL FAKE POOLS ----------
    pool_size = 5000
    first_male  = [fake.first_name_male() for _ in range(pool_size)]
    first_fem   = [fake.first_name_female() for _ in range(pool_size)]
    last_names  = [fake.last_name() for _ in range(pool_size)]
    cities      = [fake.city() for _ in range(pool_size)]
    states      = [fake.state() for _ in range(500)]
    providers_pool = [fake.company() for _ in range(6000)]
    specialties = ["Cardiology","Orthopedics","Pediatrics","General Medicine","Dermatology","ENT","OBGYN","Neurology","Oncology","Urology"]
    diag_pool   = [f"ICD{rng.integers(10,99)}.{rng.integers(0,99):02d}" for _ in range(5000)]
    proc_pool   = [f"CPT{rng.integers(10000,99999)}" for _ in range(5000)]

    # ---------- Members ----------
    m_clean, m_bad = COUNTS["members"]
    genders = np.array(["male","female"])
    g = genders[rng.integers(0,2,size=m_clean)]
    fn = np.where(g == "male", sample_from_pool(rng, first_male, m_clean), sample_from_pool(rng, first_fem, m_clean))
    ln = sample_from_pool(rng, last_names, m_clean)
    full = np.char.add(np.char.add(fn, " "), ln)
    email = make_email(fn, ln, rng)
    city = sample_from_pool(rng, cities, m_clean)
    state = sample_from_pool(rng, states, m_clean)
    phone = np.array([fake.phone_number() for _ in range(m_clean)], dtype=object)
    dob = (datetime.now() - pd.to_timedelta(rng.integers(18*365, 85*365, size=m_clean), unit="D")).astype("datetime64[ns]")
    start = datetime.now() - timedelta(days=365*3); end = datetime.now() - timedelta(days=1)
    created = rand_dates(rng, m_clean, start, end)

    members_clean = pd.DataFrame({
        "member_id": [f"M{i:07d}" for i in range(1, m_clean+1)],
        "first_name": fn, "last_name": ln, "full_name": full,
        "email": email, "gender": g, "city": city, "state": state,
        "phone": phone, "dob": dob, "created_at": created, "is_active": True
    })

    mb_n = m_bad
    g2 = genders[rng.integers(0,2,size=mb_n)]
    fn2 = np.where(g2 == "male", sample_from_pool(rng, first_male, mb_n), sample_from_pool(rng, first_fem, mb_n))
    ln2 = sample_from_pool(rng, last_names, mb_n)
    email2 = make_email(fn2, ln2, rng)
    city2 = sample_from_pool(rng, cities, mb_n)
    state2 = sample_from_pool(rng, states, mb_n)
    dob2 = (datetime.now() + pd.to_timedelta(rng.integers(1, 1200, size=mb_n), unit="D")).astype("datetime64[ns]")  # future dates
    created2 = rand_dates(rng, mb_n, start, end)

    members_bad = pd.DataFrame({
        "member_id": [f"M{i:07d}" for i in range(m_clean+1, m_clean+mb_n+1)],
        "first_name": fn2, "last_name": ln2, "full_name": np.char.add(np.char.add(fn2, " "), ln2),
        "email": email2, "gender": g2, "city": city2, "state": state2,
        "phone": "NOT_A_NUMBER", "dob": dob2, "created_at": created2, "is_active": True
    })
    if mb_n > 5:
        members_bad.loc[members_bad.sample(frac=0.35, random_state=101).index, ["city","email","first_name"]] = None
        members_bad.loc[members_bad.sample(frac=0.20, random_state=102).index, "member_id"] = members_clean["member_id"].sample(n=int(mb_n*0.20), replace=True, random_state=103).values

    # ---------- Providers ----------
    p_clean, p_bad = COUNTS["providers"]
    prov_names = sample_from_pool(rng, providers_pool, p_clean)
    providers_clean = pd.DataFrame({
        "provider_id": [f"P{i:06d}" for i in range(1, p_clean+1)],
        "provider_name": prov_names,
        "specialty": sample_from_pool(rng, specialties, p_clean),
        "city": sample_from_pool(rng, cities, p_clean),
        "state": sample_from_pool(rng, states, p_clean),
        "tax_id": [f"{rng.integers(100000000, 999999999)}" for _ in range(p_clean)],
        "license_no": [f"LIC{rng.integers(100000, 999999)}" for _ in range(p_clean)]
    })
    providers_bad = pd.DataFrame({
        "provider_id": [f"P{i:06d}" for i in range(p_clean+1, p_clean+p_bad+1)],
        "provider_name": sample_from_pool(rng, providers_pool, p_bad),
        "specialty": sample_from_pool(rng, specialties + [None], p_bad),
        "city": sample_from_pool(rng, cities, p_bad),
        "state": sample_from_pool(rng, states, p_bad),
        "tax_id": "XXXXX",
        "license_no": "BADLIC"
    })
    if p_bad > 5:
        providers_bad.loc[providers_bad.sample(frac=0.25, random_state=104).index, "provider_id"] = providers_clean["provider_id"].sample(n=int(p_bad*0.25), replace=True, random_state=105).values

    # ---------- Claims ----------
    c_clean, c_bad = COUNTS["claims"]
    mem_pool = np.array(members_clean["member_id"].values, dtype=object)
    prov_pool = np.array(providers_clean["provider_id"].values, dtype=object)
    claim_ids_clean = np.array([f"CL{i:08d}" for i in range(1, c_clean+1)], dtype=object)
    claim_ids_bad   = np.array([f"CL{i:08d}" for i in range(c_clean+1, c_clean+c_bad+1)], dtype=object)

    admit = rand_dates(rng, c_clean, datetime.now()-timedelta(days=900), datetime.now()-timedelta(days=1))
    los = np.clip(rng.integers(0, 30, size=c_clean), 0, None)
    discharge = (pd.to_datetime(admit) + pd.to_timedelta(los, unit="D")).astype("datetime64[ns]")
    allowed = np.round(rng.uniform(500, 200000, size=c_clean), 2)
    paid    = np.round(allowed * rng.uniform(0.5, 1.0, size=c_clean), 2)

    claims_clean = pd.DataFrame({
        "claim_number": claim_ids_clean,
        "member_id": mem_pool[rng.integers(0, len(mem_pool), size=c_clean)],
        "provider_id": prov_pool[rng.integers(0, len(prov_pool), size=c_clean)],
        "admit_date": admit,
        "discharge_date": discharge,
        "length_of_stay": los,
        "allowed_amount": allowed,
        "paid_amount": paid,
        "denial_code": sample_from_pool(rng, ["", "", "", "D001","D002","D003"], c_clean)  # mostly empty
    })

    admit_b = rand_dates(rng, c_bad, datetime.now()-timedelta(days=900), datetime.now()-timedelta(days=1))
    los_b = rng.integers(-5, 60, size=c_bad)  # negatives + extremes
    discharge_b = (pd.to_datetime(admit_b) + pd.to_timedelta(los_b, unit="D")).astype("datetime64[ns]")
    allowed_b = -np.abs(np.round(rng.uniform(10, 5000, size=c_bad), 2))
    paid_b    = "TEXT"

    claims_bad = pd.DataFrame({
        "claim_number": claim_ids_bad,
        "member_id": "__ORPHAN__",
        "provider_id": "__ORPHAN__",
        "admit_date": admit_b,
        "discharge_date": discharge_b,
        "length_of_stay": los_b,
        "allowed_amount": allowed_b,
        "paid_amount": paid_b,
        "denial_code": sample_from_pool(rng, ["D001","D002", None], c_bad)
    })
    if c_bad > 5:
        idx = claims_bad.sample(frac=0.25, random_state=106).index
        claims_bad.loc[idx, "admit_date"] = pd.to_datetime(claims_bad.loc[idx, "admit_date"]) + pd.to_timedelta(rng.integers(1, 365, size=len(idx)), unit="D")
        idx2 = claims_bad.sample(frac=0.25, random_state=107).index
        claims_bad.loc[idx2, "claim_number"] = claims_clean["claim_number"].sample(n=len(idx2), replace=True, random_state=108).values

    # ---------- Claim Lines ----------
    l_clean, l_bad = COUNTS["claim_lines"]
    claim_fk_pool = claim_ids_clean
    units = rng.integers(1, 10, size=l_clean)
    rate = np.round(rng.uniform(50, 20000, size=l_clean), 2)
    line_amt = np.round(units * rate, 2)
    svc_dates = rand_dates(rng, l_clean, datetime.now()-timedelta(days=900), datetime.now()-timedelta(days=1))

    claim_lines_clean = pd.DataFrame({
        "line_id": [f"LN{i:09d}" for i in range(1, l_clean+1)],
        "claim_number": claim_fk_pool[rng.integers(0, len(claim_fk_pool), size=l_clean)],
        "service_date": svc_dates,
        "procedure_code": sample_from_pool(rng, proc_pool, l_clean),
        "units": units,
        "rate": rate,
        "line_amount": line_amt
    })
    claim_lines_bad = pd.DataFrame({
        "line_id": [f"LN{i:09d}" for i in range(l_clean+1, l_clean+l_bad+1)],
        "claim_number": "__ORPHAN__",
        "service_date": rand_dates(rng, l_bad, datetime.now()-timedelta(days=900), datetime.now()-timedelta(days=1)),
        "procedure_code": sample_from_pool(rng, proc_pool + [None], l_bad),
        "units": -np.abs(rng.integers(1, 5, size=l_bad)),
        "rate": "MISSING",
        "line_amount": -np.abs(np.round(rng.uniform(1, 999, size=l_bad), 2))
    })
    if l_bad > 5:
        idx = claim_lines_bad.sample(frac=0.25, random_state=109).index
        claim_lines_bad.loc[idx, "service_date"] = pd.to_datetime(claim_lines_bad.loc[idx, "service_date"]) + pd.to_timedelta(rng.integers(1, 365, size=len(idx)), unit="D")
        idx2 = claim_lines_bad.sample(frac=0.25, random_state=110).index
        claim_lines_bad.loc[idx2, "line_id"] = claim_lines_clean["line_id"].sample(n=len(idx2), replace=True, random_state=111).values

    # ---------- Payments ----------
    py_clean, py_bad = COUNTS["payments"]
    payments_clean = pd.DataFrame({
        "payment_id": [f"PM{i:08d}" for i in range(1, py_clean+1)],
        "claim_number": claim_fk_pool[rng.integers(0, len(claim_fk_pool), size=py_clean)],
        "payment_date": rand_dates(rng, py_clean, datetime.now()-timedelta(days=900), datetime.now()-timedelta(days=1)),
        "payment_method": sample_from_pool(rng, ["NEFT","RTGS","UPI","CHEQUE"], py_clean),
        "paid_amount": np.round(rng.uniform(200, 150000, size=py_clean), 2),
        "adjustment_code": sample_from_pool(rng, ["","ADJ01","ADJ02","ADJ03"], py_clean)
    })
    payments_bad = pd.DataFrame({
        "payment_id": [f"PM{i:08d}" for i in range(py_clean+1, py_clean+py_bad+1)],
        "claim_number": "__ORPHAN__",
        "payment_date": rand_dates(rng, py_bad, datetime.now()-timedelta(days=900), datetime.now()-timedelta(days=1)),
        "payment_method": sample_from_pool(rng, ["NEFT","RTGS","UPI","CHEQUE", None], py_bad),
        "paid_amount": "ERR",
        "adjustment_code": None
    })
    if py_bad > 5:
        idx = payments_bad.sample(frac=0.25, random_state=112).index
        payments_bad.loc[idx, "payment_date"] = pd.to_datetime(payments_bad.loc[idx, "payment_date"]) + pd.to_timedelta(rng.integers(1, 365, size=len(idx)), unit="D")
        idx2 = payments_bad.sample(frac=0.25, random_state=113).index
        payments_bad.loc[idx2, "payment_id"] = payments_clean["payment_id"].sample(n=len(idx2), replace=True, random_state=114).values

    # ---------- Diagnoses ----------
    d_clean, d_bad = COUNTS["diagnoses"]
    diagnoses_clean = pd.DataFrame({
        "diag_id": [f"DX{i:07d}" for i in range(1, d_clean+1)],
        "claim_number": claim_fk_pool[rng.integers(0, len(claim_fk_pool), size=d_clean)],
        "diagnosis_code": sample_from_pool(rng, diag_pool, d_clean),
        "icd_version": sample_from_pool(rng, ["ICD-10","ICD-9"], d_clean)
    })
    diagnoses_bad = pd.DataFrame({
        "diag_id": [f"DX{i:07d}" for i in range(d_clean+1, d_clean+d_bad+1)],
        "claim_number": "__ORPHAN__",
        "diagnosis_code": sample_from_pool(rng, diag_pool + [None], d_bad),
        "icd_version": "X"
    })
    if d_bad > 5:
        idx = diagnoses_bad.sample(frac=0.25, random_state=115).index
        diagnoses_bad.loc[idx, "diag_id"] = diagnoses_clean["diag_id"].sample(n=len(idx), replace=True, random_state=116).values

    # ---------- Procedures ----------
    pr_clean, pr_bad = COUNTS["procedures"]
    procedures_clean = pd.DataFrame({
        "proc_id": [f"PR{i:07d}" for i in range(1, pr_clean+1)],
        "line_id": claim_lines_clean["line_id"].sample(n=pr_clean, replace=True, random_state=117).values,
        "procedure_code": sample_from_pool(rng, proc_pool, pr_clean)
    })
    procedures_bad = pd.DataFrame({
        "proc_id": [f"PR{i:07d}" for i in range(pr_clean+1, pr_clean+pr_bad+1)],
        "line_id": "__ORPHAN__",
        "procedure_code": sample_from_pool(rng, proc_pool + [None], pr_bad)
    })
    if pr_bad > 5:
        idx = procedures_bad.sample(frac=0.25, random_state=118).index
        procedures_bad.loc[idx, "proc_id"] = procedures_clean["proc_id"].sample(n=len(idx), replace=True, random_state=119).values

    # ---------- Write merged (shuffle) ----------
    totals = {}
    totals["members"]       = write_merged(members_clean, members_bad, data_dir / "members_merged.csv", args.seed, args.parquet)
    totals["providers"]     = write_merged(providers_clean, providers_bad, data_dir / "providers_merged.csv", args.seed, args.parquet)
    totals["claims"]        = write_merged(claims_clean, claims_bad, data_dir / "claims_merged.csv", args.seed, args.parquet)
    totals["claim_lines"]   = write_merged(claim_lines_clean, claim_lines_bad, data_dir / "claim_lines_merged.csv", args.seed, args.parquet)
    totals["payments"]      = write_merged(payments_clean, payments_bad, data_dir / "payments_merged.csv", args.seed, args.parquet)
    totals["diagnoses"]     = write_merged(diagnoses_clean, diagnoses_bad, data_dir / "diagnoses_merged.csv", args.seed, args.parquet)
    totals["procedures"]    = write_merged(procedures_clean, procedures_bad, data_dir / "procedures_merged.csv", args.seed, args.parquet)

    # ---------- Logs ----------
    profiling = {
        "seed": args.seed,
        "entities": {
            k: {"clean": int(COUNTS[k][0]), "bad": int(COUNTS[k][1]), "total": int(totals[k])}
            for k in COUNTS
        },
        "grand_totals": {
            "clean_sum": int(sum(COUNTS[k][0] for k in COUNTS)),
            "bad_sum": int(sum(COUNTS[k][1] for k in COUNTS)),
            "merged_sum": int(sum(totals.values()))
        },
        "notes": [
            "Clean+bad are concatenated and shuffled before write.",
            "Indian Faker pools used; no per-row Faker inside big loops.",
            "Bad rows include missing values, negatives, future dates, duplicate keys, orphan references, and type mismatches.",
            "Masking to be applied in ETL (student task). Fields with PII/PHI are present by design."
        ]
    }
    (logs_dir / "profiling_summary.json").write_text(json.dumps(profiling, indent=2))
    (logs_dir / "generation_config.json").write_text(json.dumps({
        "version": "healthcare-v1.0-fast",
        "faker_locale": "en_IN",
        "quick": args.quick,
        "write_parquet": bool(args.parquet)
    }, indent=2))

    print("Done. Files at:", root)

if __name__ == "__main__":
    main()
