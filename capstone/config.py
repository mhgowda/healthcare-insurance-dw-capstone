"""
config.py
---------
Central place for all file paths and settings.
Edit DATA_DIR if your CSV files are in a different location.
"""
from pathlib import Path

# -- Paths ------------------------------------------------------------------
# Root of the capstone folder
BASE_DIR = Path(__file__).resolve().parent

# Where the source CSV files live
DATA_DIR = BASE_DIR.parent / "insurance_capstone_dataset_v1" / "data" / "merged"

# logs/ folder -- matches the folder structure in the project document
LOGS_DIR = BASE_DIR.parent / "insurance_capstone_dataset_v1" / "logs"

# Output folder (database, charts, query results)
OUTPUT_DIR   = BASE_DIR / "output"
REJECTED_DIR = OUTPUT_DIR / "rejected_rows"
CHARTS_DIR   = OUTPUT_DIR / "charts"

# SQLite database file
DB_PATH = OUTPUT_DIR / "warehouse.db"

# Create output directories if they don't exist
for _d in (OUTPUT_DIR, REJECTED_DIR, CHARTS_DIR, LOGS_DIR):
    _d.mkdir(parents=True, exist_ok=True)

# -- Source CSV files --------------------------------------------------------
CSV = {
    "members":     DATA_DIR / "members_merged.csv",
    "providers":   DATA_DIR / "providers_merged.csv",
    "claims":      DATA_DIR / "claims_merged.csv",
    "claim_lines": DATA_DIR / "claim_lines_merged.csv",
    "payments":    DATA_DIR / "payments_merged.csv",
    "diagnoses":   DATA_DIR / "diagnoses_merged.csv",
    "procedures":  DATA_DIR / "procedures_merged.csv",
}

# -- Masking salt (change this to something secret in real projects) ---------
MASKING_SALT = "capstone_project_salt_2024"

# -- Today's date (used to detect future dates) -----------------------------
import pandas as pd
TODAY = pd.Timestamp("today").normalize()
