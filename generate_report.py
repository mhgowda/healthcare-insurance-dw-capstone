"""
generate_report.py
------------------
Generates a professional ~20-page Word document report for the
Healthcare Insurance Claims Data Warehouse project.

Run:  python generate_report.py
Out:  Healthcare_Insurance_DW_Report.docx
"""

from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import datetime

OUT = "Healthcare_Insurance_DW_Report.docx"

# ── Colours ──────────────────────────────────────────────────────────────────
NAVY   = RGBColor(0x0D, 0x1B, 0x3E)
BLUE   = RGBColor(0x1A, 0x6F, 0xC4)
TEAL   = RGBColor(0x00, 0xBF, 0xA5)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
LGRAY  = RGBColor(0xF0, 0xF4, 0xFA)
MGRAY  = RGBColor(0xD0, 0xD8, 0xE8)
DARK   = RGBColor(0x12, 0x1A, 0x2E)
GREEN  = RGBColor(0x2E, 0x7D, 0x32)
RED    = RGBColor(0xC6, 0x28, 0x28)
ORANGE = RGBColor(0xE6, 0x5C, 0x00)

doc = Document()

# ── Page margins ─────────────────────────────────────────────────────────────
for section in doc.sections:
    section.top_margin    = Cm(2.0)
    section.bottom_margin = Cm(2.0)
    section.left_margin   = Cm(2.5)
    section.right_margin  = Cm(2.5)

# ── Style helpers ─────────────────────────────────────────────────────────────

def set_cell_bg(cell, hex_color: str):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd  = OxmlElement("w:shd")
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  hex_color)
    tcPr.append(shd)


def cell_text(cell, text, bold=False, sz=10, color=None, align=WD_ALIGN_PARAGRAPH.LEFT):
    p = cell.paragraphs[0]
    p.alignment = align
    run = p.add_run(text)
    run.bold = bold
    run.font.size = Pt(sz)
    if color:
        run.font.color.rgb = color


def heading1(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(18)
    p.paragraph_format.space_after  = Pt(6)
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(18)
    run.font.color.rgb = NAVY
    # bottom border
    pPr = p._p.get_or_add_pPr()
    pb  = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"),  "single")
    bottom.set(qn("w:sz"),   "6")
    bottom.set(qn("w:space"),"1")
    bottom.set(qn("w:color"),"00BFA5")
    pb.append(bottom)
    pPr.append(pb)
    return p


def heading2(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after  = Pt(4)
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(13)
    run.font.color.rgb = BLUE
    return p


def heading3(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after  = Pt(2)
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(11)
    run.font.color.rgb = DARK
    return p


def body(text, space_after=6):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.space_before = Pt(0)
    run = p.add_run(text)
    run.font.size = Pt(10.5)
    run.font.color.rgb = DARK
    return p


def bullet(text, level=0):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.space_after  = Pt(3)
    p.paragraph_format.left_indent  = Inches(0.25 * (level + 1))
    run = p.add_run(text)
    run.font.size = Pt(10.5)
    run.font.color.rgb = DARK
    return p


def info_box(label, content):
    """Grey shaded box for callouts."""
    t = doc.add_table(rows=1, cols=1)
    t.style = "Table Grid"
    cell = t.cell(0, 0)
    set_cell_bg(cell, "EEF2FA")
    cell.paragraphs[0].clear()
    p  = cell.paragraphs[0]
    r1 = p.add_run(label + "  ")
    r1.bold = True; r1.font.size = Pt(10); r1.font.color.rgb = NAVY
    r2 = p.add_run(content)
    r2.font.size = Pt(10); r2.font.color.rgb = DARK
    doc.add_paragraph()


def make_table(headers, rows, col_widths=None, header_bg="0D1B3E"):
    t = doc.add_table(rows=1+len(rows), cols=len(headers))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.LEFT
    # Header row
    hdr = t.rows[0]
    for i, h in enumerate(headers):
        c = hdr.cells[i]
        set_cell_bg(c, header_bg)
        cell_text(c, h, bold=True, sz=10, color=WHITE, align=WD_ALIGN_PARAGRAPH.CENTER)
        if col_widths:
            c.width = Inches(col_widths[i])
    # Data rows
    for ri, row in enumerate(rows):
        bg = "FFFFFF" if ri % 2 == 0 else "EEF2FA"
        for ci, val in enumerate(row):
            c = t.rows[ri+1].cells[ci]
            set_cell_bg(c, bg)
            cell_text(c, str(val), sz=10, color=DARK)
            if col_widths:
                c.width = Inches(col_widths[ci])
    doc.add_paragraph()
    return t


def page_break():
    doc.add_page_break()


# ══════════════════════════════════════════════════════════════════════════════
# COVER PAGE
# ══════════════════════════════════════════════════════════════════════════════

# Title block
p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(60)
p.paragraph_format.space_after  = Pt(6)
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run("Healthcare Insurance Claims")
run.bold = True; run.font.size = Pt(28); run.font.color.rgb = NAVY

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(4)
run = p.add_run("Data Warehouse & Analytics")
run.bold = True; run.font.size = Pt(24); run.font.color.rgb = BLUE

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(40)
run = p.add_run("Project Report")
run.font.size = Pt(16); run.font.color.rgb = DARK

# Divider line
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(40)
run = p.add_run("─" * 60)
run.font.size = Pt(8); run.font.color.rgb = TEAL

# Details table
t = doc.add_table(rows=6, cols=2)
t.style = "Table Grid"
t.alignment = WD_TABLE_ALIGNMENT.CENTER
details = [
    ("Project Type",    "Midterm Capstone - 3-Day Sprint"),
    ("Organisation",    "National Health Insurance Analytics"),
    ("Technology",      "Python  |  Pandas  |  SQLite  |  SQL  |  Matplotlib"),
    ("Submitted by",    "Group 5"),
    ("Batch",           "2024-25"),
    ("Date",            datetime.date.today().strftime("%B %d, %Y")),
]
for i, (k, v) in enumerate(details):
    bg = "EEF2FA" if i % 2 == 0 else "FFFFFF"
    lc = t.rows[i].cells[0]
    rc = t.rows[i].cells[1]
    set_cell_bg(lc, "0D1B3E")
    set_cell_bg(rc, bg)
    cell_text(lc, k, bold=True, sz=10, color=WHITE)
    cell_text(rc, v, sz=10, color=DARK)
    lc.width = Inches(2.0)
    rc.width = Inches(4.5)

doc.add_paragraph()
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run("GitHub: https://github.com/mhgowda/healthcare-insurance-dw-capstone")
run.font.size = Pt(9); run.font.color.rgb = BLUE
run.font.italic = True

page_break()

# ══════════════════════════════════════════════════════════════════════════════
# TABLE OF CONTENTS  (manual)
# ══════════════════════════════════════════════════════════════════════════════
heading1("Table of Contents")

toc_items = [
    ("1",   "Executive Summary",                               "3"),
    ("2",   "Project Overview",                                "4"),
    ("2.1", "Business Scenario",                              "4"),
    ("2.2", "Project Objectives",                             "4"),
    ("2.3", "Technology Stack",                               "4"),
    ("3",   "Source Data Description",                        "5"),
    ("3.1", "Dataset Overview",                               "5"),
    ("3.2", "Data Quality Issues",                            "6"),
    ("4",   "Star Schema Design",                             "7"),
    ("4.1", "Dimensional Model",                              "7"),
    ("4.2", "Fact Tables",                                    "8"),
    ("4.3", "Dimension Tables",                               "8"),
    ("4.4", "Indexes",                                        "9"),
    ("5",   "ETL Pipeline",                                   "9"),
    ("5.1", "Pipeline Architecture",                          "9"),
    ("5.2", "Full Load",                                      "10"),
    ("5.3", "Incremental Load",                               "10"),
    ("5.4", "Audit and Rejected Rows",                        "11"),
    ("6",   "Data Masking",                                   "11"),
    ("6.1", "Why Masking",                                    "11"),
    ("6.2", "Masking Technique",                              "12"),
    ("6.3", "Masking Policy Table",                           "12"),
    ("7",   "Data Cleaning Results",                          "13"),
    ("8",   "Analytical SQL Queries",                         "14"),
    ("9",   "Python Analytics and Visualizations",            "15"),
    ("9.1", "Analytics Results",                              "15"),
    ("9.2", "Charts Generated",                               "16"),
    ("10",  "Optimization",                                   "17"),
    ("11",  "Deliverables Checklist",                         "18"),
    ("12",  "Conclusions and Future Scope",                   "19"),
]

for num, title, pg in toc_items:
    p = doc.add_paragraph()
    p.paragraph_format.space_after  = Pt(2)
    p.paragraph_format.space_before = Pt(0)
    indent = Inches(0.3) if "." in num else Inches(0)
    p.paragraph_format.left_indent = indent
    r1 = p.add_run(f"{num}  {title}")
    r1.font.size = Pt(10.5)
    r1.font.color.rgb = DARK
    r1.bold = ("." not in num)
    r2 = p.add_run(f"  {'.' * (60 - len(num) - len(title) - 4)}  {pg}")
    r2.font.size = Pt(10)
    r2.font.color.rgb = RGBColor(0x88, 0x88, 0x88)

page_break()

# ══════════════════════════════════════════════════════════════════════════════
# 1. EXECUTIVE SUMMARY
# ══════════════════════════════════════════════════════════════════════════════
heading1("1.  Executive Summary")

body(
    "This report documents the design, development, and results of an end-to-end "
    "Healthcare Insurance Claims Data Warehouse project. The project was completed "
    "as a three-day structured sprint covering SQL schema design, ETL pipeline "
    "development with data masking, and Python analytics with visualization."
)
body(
    "The source data consists of seven CSV files totalling approximately 1.4 million "
    "rows, representing real-world insurance operations including members, providers, "
    "claims, payments, diagnoses, and procedures. The data intentionally contains "
    "quality issues such as missing values, orphan foreign keys, negative amounts, "
    "future dates, duplicate keys, and type mismatches."
)
body(
    "After full ETL processing, 1,249,204 valid rows were loaded into a star-schema "
    "warehouse. Key findings include a 49.92% claim denial rate, Rs 15.68 billion "
    "in total paid claims, and Punjab and Odisha states generating the highest "
    "insurance spend. All personally identifiable information was masked using "
    "deterministic HMAC-SHA256 tokenisation before warehouse loading."
)

info_box("Key Result:", "208,640 valid claims loaded | Rs 15.68B total paid | 49.92% denial rate | 8 charts generated")

page_break()

# ══════════════════════════════════════════════════════════════════════════════
# 2. PROJECT OVERVIEW
# ══════════════════════════════════════════════════════════════════════════════
heading1("2.  Project Overview")
heading2("2.1  Business Scenario")
body(
    "A national health insurance company collects and processes millions of "
    "records annually. Data about members (insured patients), providers (doctors "
    "and hospitals), claims (insurance bills), and payments flows in from multiple "
    "source systems as flat CSV extracts. This data is siloed, messy, and contains "
    "sensitive personal and medical information."
)
body(
    "The company needs a centralised Data Warehouse to answer business questions "
    "about costs, provider performance, claim denial patterns, and geographic "
    "utilisation — while fully protecting patient privacy (PII/PHI)."
)

heading2("2.2  Project Objectives")
for obj in [
    "Design a healthcare claims star schema (fact + dimension tables)",
    "Develop a Full and Incremental ETL pipeline from CSV to warehouse",
    "Implement deterministic data masking to protect PII/PHI",
    "Profile, validate, and clean the source data",
    "Execute analytical SQL queries for business KPIs",
    "Build Python/Pandas analytics and Matplotlib visualisations",
    "Optimise the warehouse with indexes and vectorised operations",
]:
    bullet(obj)

heading2("2.3  Technology Stack")
make_table(
    ["Tool", "Version", "Purpose in This Project"],
    [
        ("Python",       "3.10+",   "Main language — ETL, analytics, report generation"),
        ("Pandas",       "2.3+",    "Data cleaning, transformation, vectorised operations"),
        ("SQLite",       "Built-in","Local data warehouse — no setup or Docker required"),
        ("SQL",          "Standard","DDL (schema creation) and analytical queries"),
        ("Matplotlib",   "3.8+",    "8 charts saved as PNG files"),
        ("python-docx",  "1.x",     "This report generation"),
        ("python-pptx",  "1.x",     "Presentation generation"),
        ("Jupyter",      "Optional","Interactive analytics notebook"),
    ],
    col_widths=[1.5, 1.0, 4.0],
)

page_break()

# ══════════════════════════════════════════════════════════════════════════════
# 3. SOURCE DATA
# ══════════════════════════════════════════════════════════════════════════════
heading1("3.  Source Data Description")
heading2("3.1  Dataset Overview")
body(
    "The source data consists of seven CSV files generated using the Indian Faker "
    "locale. The dataset intentionally mixes clean rows with imperfect rows to "
    "simulate real-world conditions. Total dataset size is approximately 1.4 million "
    "rows across all files."
)

make_table(
    ["CSV File", "Description", "Rows", "Primary Key"],
    [
        ("members_merged.csv",     "Patient master — name, DOB, email, city, state",    "2,60,000", "member_id"),
        ("providers_merged.csv",   "Doctor/hospital master — specialty, city, tax_id",  "40,000",   "provider_id"),
        ("claims_merged.csv",      "Insurance claim headers — amount, date, denial",    "2,40,000", "claim_number"),
        ("claim_lines_merged.csv", "Service lines inside each claim",                   "7,60,000", "line_id"),
        ("payments_merged.csv",    "Payment transactions on claims",                    "1,00,000", "payment_id"),
        ("diagnoses_merged.csv",   "ICD diagnosis codes per claim",                     "60,000",   "diag_id"),
        ("procedures_merged.csv",  "CPT procedure codes per claim line",                "40,000",   "proc_id"),
    ],
    col_widths=[2.2, 2.8, 0.9, 1.3],
)

info_box("Total:", "7 files  |  ~14,00,000 rows  |  1,200,000 clean + 200,000 imperfect")

heading2("3.2  Data Quality Issues")
body(
    "The following categories of data quality problems were identified during the "
    "profiling phase (step1_profile.py) and handled during cleaning (step2_clean.py):"
)

make_table(
    ["Issue Type", "Description", "Affected Columns", "Example"],
    [
        ("Missing Values",    "NULL or blank required fields",
         "first_name, email, city, diagnosis_code",  "first_name = NULL"),
        ("Orphan FKs",        "__ORPHAN__ marker — parent record missing",
         "member_id in claims, claim_number in payments", "member_id = '__ORPHAN__'"),
        ("Negative Amounts",  "Invalid negative financial values",
         "allowed_amount, paid_amount, units",       "allowed_amount = -500"),
        ("Future Dates",      "Timestamps in the future (impossible)",
         "admit_date, service_date, dob",            "dob = 2028-05-01"),
        ("Duplicate Keys",    "Same primary key repeated",
         "claim_number, member_id",                  "Two rows with CL00001234"),
        ("Type Mismatches",   "Text where numeric is required",
         "paid_amount, rate, length_of_stay",        "paid_amount = 'ERR'"),
    ],
    col_widths=[1.5, 1.8, 2.2, 1.7],
)

page_break()

# ══════════════════════════════════════════════════════════════════════════════
# 4. STAR SCHEMA DESIGN
# ══════════════════════════════════════════════════════════════════════════════
heading1("4.  Star Schema Design")

heading2("4.1  Dimensional Model")
body(
    "A star schema is the standard design pattern for data warehouses. It consists "
    "of Fact Tables (which store measurable numeric data — amounts, counts, durations) "
    "surrounded by Dimension Tables (which store descriptive context — who, what, "
    "where, when). Analytical queries join from a fact table to one or more dimensions."
)
body(
    "The healthcare DW uses the following structure:"
)

p = doc.add_paragraph()
p.paragraph_format.space_after = Pt(4)
run = p.add_run(
    "                      dim_date\n"
    "                          |\n"
    "  dim_member  --  fact_claim  --  dim_provider\n"
    "                          |\n"
    "               fact_claim_line  --  dim_procedure\n"
    "                          |\n"
    "                    fact_payment\n"
    "                          |\n"
    "               dim_diagnosis   dim_location"
)
run.font.name = "Courier New"
run.font.size = Pt(9)
run.font.color.rgb = DARK

heading2("4.2  Fact Tables")
make_table(
    ["Table", "Grain", "Key Measures", "Foreign Keys"],
    [
        ("fact_claim",      "One row per claim",
         "allowed_amount, paid_amount, length_of_stay, is_denied",
         "member_id, provider_id, admit_date_key"),
        ("fact_claim_line", "One row per service line",
         "units, rate, line_amount",
         "claim_number, service_date_key, procedure_code"),
        ("fact_payment",    "One row per payment",
         "paid_amount",
         "claim_number, payment_date_key"),
    ],
    col_widths=[1.6, 1.6, 2.4, 1.6],
)

heading2("4.3  Dimension Tables")
make_table(
    ["Table", "Description", "Key Columns", "Masking Applied"],
    [
        ("dim_member",    "Patient/member details",
         "member_id, gender, city, state, birth_year, is_active",
         "member_id, names, email, phone, dob"),
        ("dim_provider",  "Doctor/hospital details",
         "provider_id, specialty, city, state",
         "provider_id, provider_name, tax_id, license_no"),
        ("dim_date",      "Calendar dimension",
         "date_key, year, quarter, month, month_name, is_weekend",
         "None"),
        ("dim_diagnosis",  "ICD diagnosis codes",
         "diagnosis_code, icd_version",
         "None"),
        ("dim_procedure",  "CPT procedure codes",
         "procedure_code",
         "None"),
        ("dim_location",   "City and state combinations",
         "location_key, city, state",
         "None"),
    ],
    col_widths=[1.5, 1.8, 2.5, 1.4],
)

heading2("4.4  Database Indexes")
body(
    "Indexes were created on all foreign key columns and frequently filtered "
    "columns to speed up JOIN operations and WHERE clauses:"
)
make_table(
    ["Index Name", "Table", "Column", "Purpose"],
    [
        ("idx_fc_member",    "fact_claim",      "member_id",       "Member dimension JOIN"),
        ("idx_fc_provider",  "fact_claim",      "provider_id",     "Provider dimension JOIN"),
        ("idx_fc_admit",     "fact_claim",      "admit_date_key",  "Date range filtering"),
        ("idx_fc_denied",    "fact_claim",      "is_denied",       "Denial rate queries"),
        ("idx_fcl_claim",    "fact_claim_line", "claim_number",    "Claim line aggregation"),
        ("idx_fp_claim",     "fact_payment",    "claim_number",    "Payment aggregation"),
        ("idx_dp_specialty", "dim_provider",    "specialty",       "Specialty filtering"),
    ],
    col_widths=[1.8, 1.8, 1.6, 2.0],
)

page_break()

# ══════════════════════════════════════════════════════════════════════════════
# 5. ETL PIPELINE
# ══════════════════════════════════════════════════════════════════════════════
heading1("5.  ETL Pipeline")

heading2("5.1  Pipeline Architecture")
body(
    "The ETL pipeline transforms raw CSV files into a clean, masked, and structured "
    "warehouse in five sequential steps:"
)

make_table(
    ["Step", "Script", "Action", "Output"],
    [
        ("1 - Profile",     "step1_profile.py",      "Read CSVs, count all quality issues",
         "logs/profiling_summary.json"),
        ("2 - Clean",       "step2_clean.py",         "Remove bad rows, enforce FK integrity",
         "output/rejected_rows/*.csv"),
        ("3 - Mask",        "step3_mask.py",          "Hash PII fields, generalise DOB to year",
         "logs/masking_audit.json"),
        ("4 - Full Load",   "step4_load_full.py",     "Drop schema, recreate, load all data",
         "output/warehouse.db"),
        ("5 - Incremental", "step5_load_incremental.py","Load only new rows using watermark",
         "etl_watermark table updated"),
    ],
    col_widths=[1.5, 1.9, 2.4, 1.4],
)

heading2("5.2  Full ETL Load")
body(
    "The full load (step4_load_full.py) performs a complete refresh of the warehouse:"
)
for s in [
    "Creates all star-schema tables using the DDL file (day1_sql/01_star_schema_ddl.sql)",
    "Calls the cleaning module — validates all entities in dependency order",
    "Calls the masking module — applies HMAC-SHA256 to PII fields",
    "Loads dimension tables first (dim_member, dim_provider, dim_date, dim_location, dim_diagnosis, dim_procedure)",
    "Loads fact tables after dimensions (fact_claim, fact_claim_line, fact_payment)",
    "Records the load timestamp in etl_watermark for incremental use",
]:
    bullet(s)

info_box("Load Time:", "Approximately 3-5 minutes for 1.4 million rows on a standard laptop")

heading2("5.3  Incremental ETL Load")
body(
    "The incremental load (step5_load_incremental.py) processes only new data "
    "since the last run, using a watermark mechanism:"
)
body(
    "A table called etl_watermark stores the last_load_dt (timestamp) for each "
    "entity. On each incremental run, the source CSV is filtered to rows where "
    "the date column is greater than last_load_dt. Only those rows go through "
    "cleaning, masking, and loading. INSERT OR IGNORE ensures no duplicates enter "
    "the warehouse."
)

make_table(
    ["Entity", "Watermark Column", "Refresh Strategy"],
    [
        ("members",     "created_at",    "Filter by created_at > last_load_dt"),
        ("claims",      "admit_date",    "Filter by admit_date > last_load_dt"),
        ("claim_lines", "service_date",  "Filter by service_date > last_load_dt"),
        ("payments",    "payment_date",  "Filter by payment_date > last_load_dt"),
        ("providers",   "(no date col)", "Full upsert — INSERT OR IGNORE all rows"),
        ("diagnoses",   "(no date col)", "Full upsert — INSERT OR IGNORE all rows"),
        ("procedures",  "(no date col)", "Full upsert — INSERT OR IGNORE all rows"),
    ],
    col_widths=[1.5, 1.6, 4.1],
)

heading2("5.4  Audit and Rejected Rows")
body(
    "Every ETL run produces audit outputs saved to output/rejected_rows/. "
    "Each entity's rejected rows are saved as a separate CSV with a rejection_reason "
    "column explaining why each row failed. The masking audit is saved as "
    "logs/masking_audit.json and records which fields were masked and how many values "
    "were processed."
)

page_break()

# ══════════════════════════════════════════════════════════════════════════════
# 6. DATA MASKING
# ══════════════════════════════════════════════════════════════════════════════
heading1("6.  Data Masking")

heading2("6.1  Why Masking Is Required")
body(
    "The source data contains personally identifiable information (PII) and "
    "protected health information (PHI). Analysts querying the warehouse to "
    "understand costs, denial rates, and provider performance do not need to see "
    "real patient names, contact details, or exact dates of birth. Loading "
    "unmasked PII into an analytics database creates unnecessary privacy risk."
)
body(
    "Data masking replaces sensitive fields with safe tokens before any data "
    "reaches the warehouse. The warehouse layer never contains raw PII."
)

heading2("6.2  Masking Technique — Deterministic HMAC-SHA256")
body(
    "The project uses HMAC-SHA256 (Hash-based Message Authentication Code) with a "
    "project-specific salt string. This approach has three critical properties:"
)
for prop in [
    "DETERMINISTIC: The same input always produces the same token. member_id 'M0000001' always becomes 'MBR3F9A2C1D4E' — so JOIN operations between fact_claim, dim_member, and dim_provider all work correctly.",
    "IRREVERSIBLE: The hash cannot be decoded back to the original value without the salt.",
    "SALT-PROTECTED: The salt string makes the tokens unique to this project. Without the exact salt, an attacker cannot reproduce the mapping.",
]:
    bullet(prop)

info_box("Salt Location:", "Stored in config.py as MASKING_SALT. Change this value to re-key all tokens.")

heading2("6.3  Masking Policy Table")
make_table(
    ["Field", "Entity", "Technique", "Before", "After"],
    [
        ("member_id",    "members",   "Deterministic token", "M0000001",          "MBR3F9A2C1D4E"),
        ("first_name",   "members",   "Deterministic token", "Alice",             "FN8B2C3D4E5F"),
        ("last_name",    "members",   "Deterministic token", "Smith",             "LN2C3D4E5F6A"),
        ("email",        "members",   "Deterministic token", "alice@gmail.com",   "EMAIF29A3B1C"),
        ("phone",        "members",   "Deterministic token", "9876543210",        "PH8A2B3C4D"),
        ("dob",          "members",   "Generalise to year",  "1985-03-10",        "1985"),
        ("provider_id",  "providers", "Deterministic token", "P000001",           "PRV7AEE1BC05"),
        ("provider_name","providers", "Deterministic token", "Acme Clinic",       "PNM3B4C5D6E7"),
        ("tax_id",       "providers", "Deterministic token", "908430420",         "TAX2B3C4D5E6"),
        ("license_no",   "providers", "Deterministic token", "LIC790693",         "LIC3C4D5E6F7"),
        ("claim_number", "claims",    "Deterministic token", "CL00155929",        "CLM4E5F6A7B8"),
        ("city / state", "members",   "Kept as-is",          "Mumbai, MH",        "Mumbai, MH"),
    ],
    col_widths=[1.3, 1.2, 1.6, 1.8, 1.3],
)

page_break()

# ══════════════════════════════════════════════════════════════════════════════
# 7. DATA CLEANING RESULTS
# ══════════════════════════════════════════════════════════════════════════════
heading1("7.  Data Cleaning Results")
body(
    "The cleaning module processes entities in dependency order: dimensions first "
    "(members, providers), then facts (claims, then claim_lines, payments, "
    "diagnoses, procedures). This ensures that FK validation is possible at each step."
)

make_table(
    ["Entity", "Total Rows", "Valid (Kept)", "Rejected", "Rejection %", "Top Rejection Reason"],
    [
        ("members",     "2,60,000", "2,32,207", "27,793",   "10.7%", "Missing first_name, future DOB"),
        ("providers",   "40,000",   "38,424",   "1,576",    "3.9%",  "Missing specialty, duplicate ID"),
        ("claims",      "2,40,000", "2,08,640", "31,360",   "13.1%", "Orphan member_id, negative amount"),
        ("claim_lines", "7,60,000", "6,45,733", "1,14,267", "15.0%", "Orphan claim_number, negative units"),
        ("payments",    "1,00,000", "59,615",   "40,385",   "40.4%", "Orphan claim_number, non-numeric amount"),
        ("diagnoses",   "60,000",   "39,741",   "20,259",   "33.8%", "Orphan claim_number, invalid ICD"),
        ("procedures",  "40,000",   "24,844",   "15,156",   "37.9%", "Orphan line_id, missing code"),
        ("TOTAL",       "14,00,000","12,49,204","2,50,796", "17.9%", "All categories combined"),
    ],
    col_widths=[1.3, 1.1, 1.2, 1.0, 1.1, 2.5],
    header_bg="0D1B3E",
)

info_box("Rejected rows:", "All 7 rejection CSV files saved to output/rejected_rows/ with rejection_reason column for audit purposes.")

page_break()

# ══════════════════════════════════════════════════════════════════════════════
# 8. ANALYTICAL SQL QUERIES
# ══════════════════════════════════════════════════════════════════════════════
heading1("8.  Analytical SQL Queries")
body(
    "Ten analytical SQL queries were developed in day1_sql/02_analytical_queries.sql "
    "and also executed via Python in day3_analytics/analytics.py. Each query "
    "answers a specific business question from the warehouse."
)

make_table(
    ["Query", "Business Question", "Key SQL Technique", "Result (Sample)"],
    [
        ("Q1", "Total claims summary",
         "COUNT, SUM, AVG, percentage",
         "208,640 claims | Rs 15.68B paid"),
        ("Q2", "Monthly paid trend",
         "JOIN fact_claim + dim_date, GROUP BY year/month",
         "~Rs 540M/month consistent"),
        ("Q3", "Top procedures by utilisation",
         "GROUP BY procedure_code, ORDER BY COUNT",
         "CPT35289: 417 uses"),
        ("Q4", "Provider performance scorecard",
         "JOIN dim_provider, AVG paid, denial_rate %",
         "Some providers >60% denial"),
        ("Q5", "Denial rate breakdown",
         "CASE WHEN denial_code, GROUP BY denial_reason",
         "D001/D002/D003 each ~16.6%"),
        ("Q6", "Length of stay distribution",
         "GROUP BY length_of_stay, AVG paid",
         "Flat 0-29 days distribution"),
        ("Q7", "Member utilisation",
         "JOIN dim_member, SUM paid, ORDER BY spend",
         "Top member: Rs 785,827"),
        ("Q8", "Geographic analysis",
         "JOIN dim_member on city/state, SUM by state",
         "Punjab leads: Rs 876M"),
        ("Q9", "Payment method breakdown",
         "GROUP BY payment_method, SUM paid",
         "NEFT/UPI/RTGS/Cheque ~equal"),
        ("Q10", "Diagnosis cost analysis",
         "Link diagnoses to claims via claim_number",
         "ICD codes by total cost"),
    ],
    col_widths=[0.5, 1.8, 2.0, 2.0],
)

heading3("Sample SQL — Provider Performance (Q4)")
p = doc.add_paragraph()
run = p.add_run(
    "SELECT dp.provider_id, dp.specialty, dp.state,\n"
    "       COUNT(fc.claim_number)          AS total_claims,\n"
    "       ROUND(SUM(fc.paid_amount), 2)   AS total_paid,\n"
    "       ROUND(AVG(fc.paid_amount), 2)   AS avg_paid,\n"
    "       ROUND(AVG(fc.length_of_stay),1) AS avg_los,\n"
    "       SUM(fc.is_denied)               AS denied_claims,\n"
    "       ROUND(100.0 * SUM(fc.is_denied)\n"
    "             / COUNT(*), 2)            AS denial_rate_pct\n"
    "FROM fact_claim fc\n"
    "JOIN dim_provider dp ON fc.provider_id = dp.provider_id\n"
    "GROUP BY dp.provider_id, dp.specialty, dp.state\n"
    "ORDER BY total_paid DESC\n"
    "LIMIT 20;"
)
run.font.name  = "Courier New"
run.font.size  = Pt(8.5)
run.font.color.rgb = DARK
p.paragraph_format.left_indent  = Inches(0.3)
p.paragraph_format.space_before = Pt(4)
p.paragraph_format.space_after  = Pt(8)

page_break()

# ══════════════════════════════════════════════════════════════════════════════
# 9. PYTHON ANALYTICS AND VISUALIZATIONS
# ══════════════════════════════════════════════════════════════════════════════
heading1("9.  Python Analytics and Visualizations")

heading2("9.1  Analytics Results")
body(
    "All nine analytical queries were executed against the live warehouse via "
    "Python using sqlite3 and Pandas. Results were saved as CSV files in "
    "output/ for further review. Key findings:"
)

make_table(
    ["Metric", "Value", "Insight"],
    [
        ("Total Valid Claims",     "2,08,640",      "After removing 31,360 bad rows from 2,40,000 raw claims"),
        ("Total Allowed Amount",   "Rs 20.92 Billion", "Total billed by providers before adjudication"),
        ("Total Paid Amount",      "Rs 15.68 Billion", "Actual insurer payouts (75% of allowed)"),
        ("Average Paid per Claim", "Rs 75,156",     "Consistent across all claim types"),
        ("Overall Denial Rate",    "49.92%",         "Nearly half of claims denied — very high"),
        ("Top Denial Codes",       "D001, D002, D003", "Each accounts for ~16.6% of all denials"),
        ("Top Spending State",     "Punjab",        "Rs 876 million total paid"),
        ("Top Procedure",          "CPT35289",      "417 uses — most utilised procedure code"),
        ("Payment Methods",        "NEFT, UPI, RTGS, Cheque", "Each contributes approximately 25%"),
        ("Avg Length of Stay",     "~14.5 days",   "Uniform distribution 0-29 days"),
    ],
    col_widths=[2.2, 1.8, 3.2],
)

heading2("9.2  Charts Generated")
body(
    "Eight Matplotlib charts were generated by day3_analytics/visualizations.py "
    "and saved as PNG files in output/charts/. Each chart addresses a specific "
    "business question:"
)

make_table(
    ["File", "Chart Type", "Business Question Answered"],
    [
        ("01_monthly_paid.png",         "Bar chart",           "Which months had highest insurance payouts?"),
        ("02_top_procedures.png",       "Horizontal bar",      "Which procedures are used most often?"),
        ("03_provider_performance.png", "Side-by-side bars",   "Which providers paid most and had highest denial rates?"),
        ("04_denial_rate.png",          "Bar chart",           "How are claims distributed across denial codes?"),
        ("05_los_distribution.png",     "Bar chart",           "How are hospital stays distributed by length?"),
        ("06_member_utilisation.png",   "Scatter plot",        "Which members drive the most claims and spend?"),
        ("07_geographic.png",           "Bar chart",           "Which states have the highest insurance spend?"),
        ("08_payment_methods.png",      "Pie chart",           "What payment methods are used and in what proportion?"),
    ],
    col_widths=[2.2, 1.4, 3.6],
)

info_box("Notebook:", "All 8 charts are also reproducible interactively in analytics_notebook.ipynb (Jupyter).")

page_break()

# ══════════════════════════════════════════════════════════════════════════════
# 10. OPTIMIZATION
# ══════════════════════════════════════════════════════════════════════════════
heading1("10.  Optimization")

heading2("10.1  Database Indexes")
body(
    "Seven indexes were defined in the DDL file covering all foreign key and "
    "commonly filtered columns. Indexes reduce query time for JOIN-heavy analytical "
    "queries from full table scans to direct index lookups."
)

heading2("10.2  Vectorised Pandas Operations")
body(
    "All data transformations use Pandas vectorised operations rather than row-by-row "
    "loops. This is critical for 1.4 million rows — a Python loop over 760,000 "
    "claim_line rows would take minutes; vectorised operations complete in seconds."
)

make_table(
    ["Operation", "Pandas Method", "Why Faster"],
    [
        ("Convert columns to numeric",  "pd.to_numeric(col, errors='coerce')",       "Whole column at once in C"),
        ("Parse date columns",           "pd.to_datetime(col, errors='coerce')",      "Vectorised date parsing"),
        ("Filter bad rows",              "df[boolean_mask]",                          "No Python loop needed"),
        ("Remap FK tokens",              "Series.map(dictionary)",                    "Hash lookup vs apply(lambda)"),
        ("Deduplicate",                  "df.drop_duplicates(subset=[pk])",           "Internal C sort algorithm"),
        ("Clip negative values",         "Series.clip(lower=0)",                      "Element-wise C operation"),
    ],
    col_widths=[2.0, 2.8, 2.4],
)

heading2("10.3  ETL Chunking")
body(
    "The SQLite loader writes data in chunks (SQLite variable limit is 32,766 per "
    "statement). The chunk size is calculated as max_vars // n_cols to ensure no "
    "statement exceeds SQLite's limit regardless of column count."
)

heading2("10.4  Incremental Loading Impact")
body(
    "The watermark-based incremental load means subsequent daily runs process "
    "only the new rows rather than reloading 1.4 million rows. In production "
    "with 1,000 new claims per day, an incremental run takes seconds versus "
    "3-5 minutes for a full reload."
)

page_break()

# ══════════════════════════════════════════════════════════════════════════════
# 11. DELIVERABLES CHECKLIST
# ══════════════════════════════════════════════════════════════════════════════
heading1("11.  Deliverables Checklist")

make_table(
    ["Deliverable", "File / Location", "Status"],
    [
        ("Star Schema DDL",          "day1_sql/01_star_schema_ddl.sql",     "Complete"),
        ("Analytical SQL Queries",   "day1_sql/02_analytical_queries.sql",  "Complete"),
        ("Data Profiling Report",    "logs/profiling_summary.json",         "Complete"),
        ("ETL Full Load Script",     "day2_etl/step4_load_full.py",         "Complete"),
        ("ETL Incremental Script",   "day2_etl/step5_load_incremental.py",  "Complete"),
        ("Data Masking Logic",       "day2_etl/step3_mask.py",              "Complete"),
        ("Masking Audit Log",        "logs/masking_audit.json",             "Complete"),
        ("Rejected Rows (7 files)",  "output/rejected_rows/*.csv",          "Complete"),
        ("SQLite Warehouse",         "output/warehouse.db",                 "Complete"),
        ("Python Analytics",         "day3_analytics/analytics.py",         "Complete"),
        ("8 Matplotlib Charts",      "output/charts/*.png",                 "Complete"),
        ("Jupyter Notebook",         "day3_analytics/analytics_notebook.ipynb","Complete"),
        ("Project Presentation",     "Healthcare_Insurance_DW_Capstone.pptx","Complete"),
        ("Project Report",           "Healthcare_Insurance_DW_Report.docx", "Complete"),
        ("README / Team Guide",      "README.md, TEAM_GUIDE.md",            "Complete"),
        ("GitHub Repository",        "github.com/mhgowda/healthcare-insurance-dw-capstone","Complete"),
    ],
    col_widths=[2.4, 2.8, 1.0],
    header_bg="0D1B3E",
)

page_break()

# ══════════════════════════════════════════════════════════════════════════════
# 12. CONCLUSIONS AND FUTURE SCOPE
# ══════════════════════════════════════════════════════════════════════════════
heading1("12.  Conclusions and Future Scope")

heading2("12.1  Project Conclusions")
body(
    "This project successfully demonstrates a complete end-to-end data engineering "
    "pipeline for healthcare insurance analytics. Starting from raw, messy CSV "
    "files, the project delivers a clean, masked, indexed warehouse with full "
    "analytical capability."
)

for conclusion in [
    "Data quality is the foundation: 17.9% of input rows were rejected — validating data before loading prevents incorrect analytics.",
    "Privacy-first design: Deterministic masking protects 100% of PII/PHI while preserving all analytical join relationships.",
    "The 49.92% denial rate is a critical business finding that warrants further investigation — unusually high and evenly spread across D001/D002/D003.",
    "Geographic analysis (possible only because city/state were unmasked) shows Punjab and Odisha as the highest-spend states.",
    "Incremental ETL makes the solution production-ready — daily refreshes process only new data rather than reloading everything.",
    "The star schema design scales well — adding new dimensions or measures only requires new dimension/fact tables and ETL steps.",
]:
    bullet(conclusion)

heading2("12.2  Limitations")
for lim in [
    "SQLite is suitable for learning and small datasets. Production deployment would use PostgreSQL or Snowflake.",
    "The dataset is synthetically generated — real-world data would have more complex quality patterns.",
    "Masking does not implement role-based access (different masks for different user roles) — a production system would add this.",
    "dim_diagnosis is not linked directly to fact_claim in this schema — a bridge table would improve diagnosis-level analytics.",
]:
    bullet(lim)

heading2("12.3  Future Scope")
make_table(
    ["Enhancement", "Description", "Technology"],
    [
        ("Snowflake Migration",    "Migrate warehouse to Snowflake for cloud scale",     "Snowflake, IDMC"),
        ("IDMC Integration",       "Use IDMC for visual ETL mapping and lineage",         "Informatica IDMC"),
        ("Role-Based Masking",     "Different masking levels for analysts vs auditors",   "Python, RBAC"),
        ("Real-time Streaming",    "Process claims as they arrive rather than batch",     "Kafka, Spark"),
        ("ML Fraud Detection",     "Flag suspicious providers using anomaly detection",   "scikit-learn"),
        ("BI Dashboard",           "Power BI or Tableau dashboard over the warehouse",   "Power BI / Tableau"),
        ("Data Lineage",           "Track every field from source CSV to warehouse",     "Apache Atlas"),
    ],
    col_widths=[1.9, 3.0, 2.3],
)

doc.add_paragraph()
body(
    "This project forms a strong foundation in modern data engineering practices. "
    "The skills demonstrated — schema design, ETL development, data quality, "
    "privacy masking, and analytics — are directly applicable to real-world "
    "data engineering roles in healthcare, insurance, finance, and technology.",
    space_after=12,
)

# ── Final page — references ────────────────────────────────────────────────
page_break()
heading1("References and Resources")
for ref in [
    "Python Documentation: https://docs.python.org/3/",
    "Pandas Documentation: https://pandas.pydata.org/docs/",
    "SQLite Documentation: https://www.sqlite.org/docs.html",
    "python-docx: https://python-docx.readthedocs.io/",
    "python-pptx: https://python-pptx.readthedocs.io/",
    "Matplotlib: https://matplotlib.org/stable/contents.html",
    "HMAC-SHA256: https://docs.python.org/3/library/hmac.html",
    "Star Schema Design (Kimball): https://www.kimballgroup.com/data-warehouse-business-intelligence-resources/",
    "HIPAA PHI / PII Guidelines: https://www.hhs.gov/hipaa/",
    "Project GitHub: https://github.com/mhgowda/healthcare-insurance-dw-capstone",
]:
    bullet(ref)

# ── Save ─────────────────────────────────────────────────────────────────────
doc.save(OUT)
print(f"\nReport saved: {OUT}")
print(f"Pages: approximately 20 (open in Word to see exact count)")
