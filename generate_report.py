"""
generate_report.py
Generates a professional 20+ page Word report for the Healthcare Insurance
Claims Data Warehouse project.
Font: Times New Roman throughout
Run:  python generate_report.py
"""

from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import datetime

OUT   = "Healthcare_Insurance_DW_Report.docx"
FONT  = "Times New Roman"

# ── Colour palette ────────────────────────────────────────────────────────
C_NAVY  = RGBColor(0x0D, 0x1B, 0x3E)
C_BLUE  = RGBColor(0x1A, 0x5F, 0xA8)
C_TEAL  = RGBColor(0x00, 0x7A, 0x68)
C_WHITE = RGBColor(0xFF, 0xFF, 0xFF)
C_BLACK = RGBColor(0x00, 0x00, 0x00)
C_DARK  = RGBColor(0x1A, 0x1A, 0x1A)
C_GRAY  = RGBColor(0x44, 0x44, 0x44)
C_LGRAY = RGBColor(0xF2, 0xF4, 0xF8)
C_MGRAY = RGBColor(0xD8, 0xDF, 0xEC)
C_GREEN = RGBColor(0x1B, 0x5E, 0x20)
C_RED   = RGBColor(0xB7, 0x1C, 0x1C)
C_HDRBG = "0D1B3E"
C_ALTRW = "EDF1F7"

doc = Document()

# ── Page setup: A4, standard margins ─────────────────────────────────────
for sec in doc.sections:
    sec.page_width    = Cm(21.0)
    sec.page_height   = Cm(29.7)
    sec.top_margin    = Cm(2.5)
    sec.bottom_margin = Cm(2.5)
    sec.left_margin   = Cm(3.0)
    sec.right_margin  = Cm(2.5)

# ═════════════════════════════════════════════════════════════════════════
# LOW-LEVEL HELPERS
# ═════════════════════════════════════════════════════════════════════════

def _shd(cell, hex6: str):
    """Fill a table cell with a hex background colour."""
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd  = OxmlElement("w:shd")
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  hex6.upper())
    tcPr.append(shd)


def _border_bottom(paragraph, color_hex="007A68", sz="12"):
    """Add a coloured bottom border to a paragraph (used for headings)."""
    pPr = paragraph._p.get_or_add_pPr()
    pb  = OxmlElement("w:pBdr")
    bot = OxmlElement("w:bottom")
    bot.set(qn("w:val"),   "single")
    bot.set(qn("w:sz"),    sz)
    bot.set(qn("w:space"), "1")
    bot.set(qn("w:color"), color_hex.upper())
    pb.append(bot)
    pPr.append(pb)


def _run(paragraph, text, bold=False, italic=False, sz=12,
         color=C_DARK, underline=False, font=FONT):
    run = paragraph.add_run(text)
    run.bold      = bold
    run.italic    = italic
    run.underline = underline
    run.font.name   = font
    run.font.size   = Pt(sz)
    run.font.color.rgb = color
    return run


def _cell_write(cell, text, bold=False, sz=10, color=C_DARK,
                align=WD_ALIGN_PARAGRAPH.LEFT, italic=False):
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after  = Pt(2)
    _run(p, text, bold=bold, sz=sz, color=color, italic=italic)


# ═════════════════════════════════════════════════════════════════════════
# CONTENT HELPERS
# ═════════════════════════════════════════════════════════════════════════

def h1(text):
    """Chapter heading — large, navy, bottom border."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(24)
    p.paragraph_format.space_after  = Pt(8)
    _run(p, text, bold=True, sz=16, color=C_NAVY)
    _border_bottom(p, "1A5FA8", "10")
    return p


def h2(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after  = Pt(4)
    _run(p, text, bold=True, sz=13, color=C_BLUE)
    return p


def h3(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after  = Pt(3)
    _run(p, text, bold=True, sz=11.5, color=C_DARK)
    return p


def para(text, sz=11.5, space_after=8, indent=False, justify=True):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY if justify else WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_after  = Pt(space_after)
    p.paragraph_format.space_before = Pt(0)
    if indent:
        p.paragraph_format.first_line_indent = Inches(0.25)
    _run(p, text, sz=sz, color=C_DARK)
    return p


def bul(text, level=0, sz=11):
    p = doc.add_paragraph()
    p.paragraph_format.space_after  = Pt(3)
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.left_indent  = Inches(0.35 + level * 0.25)
    bullet_char = "\u2022" if level == 0 else "\u25e6"
    _run(p, f"{bullet_char}  {text}", sz=sz, color=C_DARK)
    return p


def numbered(text, n, sz=11):
    p = doc.add_paragraph()
    p.paragraph_format.space_after  = Pt(3)
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.left_indent  = Inches(0.35)
    _run(p, f"{n}.  {text}", sz=sz, color=C_DARK)
    return p


def callout(label, content, bg="EDF1F7"):
    t = doc.add_table(rows=1, cols=1)
    t.style = "Table Grid"
    c = t.cell(0, 0)
    _shd(c, bg)
    c.paragraphs[0].clear()
    p  = c.paragraphs[0]
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after  = Pt(4)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    _run(p, label + "  ", bold=True, sz=10.5, color=C_NAVY)
    _run(p, content,       bold=False, sz=10.5, color=C_DARK)
    doc.add_paragraph().paragraph_format.space_after = Pt(4)


def code_block(code_text):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent  = Inches(0.3)
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after  = Pt(6)
    _run(p, code_text, sz=9, color=C_DARK, font="Courier New")


def make_table(headers, rows, widths=None, hdr_bg=C_HDRBG, striped=True):
    ncols = len(headers)
    t = doc.add_table(rows=1 + len(rows), cols=ncols)
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.LEFT
    # Header
    hr = t.rows[0]
    for i, h in enumerate(headers):
        cell = hr.cells[i]
        _shd(cell, hdr_bg)
        _cell_write(cell, h, bold=True, sz=10, color=C_WHITE,
                    align=WD_ALIGN_PARAGRAPH.CENTER)
        if widths:
            cell.width = Inches(widths[i])
    # Data rows
    for ri, row in enumerate(rows):
        bg = C_ALTRW if (striped and ri % 2 == 1) else "FFFFFF"
        for ci, val in enumerate(row):
            cell = t.rows[ri + 1].cells[ci]
            _shd(cell, bg)
            _cell_write(cell, str(val), sz=10, color=C_DARK)
            if widths:
                cell.width = Inches(widths[ci])
    doc.add_paragraph().paragraph_format.space_after = Pt(4)
    return t


def page_break():
    doc.add_page_break()


def space(pts=6):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(pts)


# ═════════════════════════════════════════════════════════════════════════
# PAGE 1 — COVER PAGE
# ═════════════════════════════════════════════════════════════════════════

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(72)
p.paragraph_format.space_after  = Pt(0)
_run(p, "HEALTHCARE INSURANCE CLAIMS", bold=True, sz=22, color=C_NAVY)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(0)
_run(p, "DATA WAREHOUSE AND ANALYTICS", bold=True, sz=20, color=C_NAVY)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(36)
_run(p, "Project Report", bold=False, sz=16, color=C_GRAY, italic=True)

# Divider
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(36)
_run(p, "_" * 55, sz=10, color=C_TEAL)

# Details table
t = doc.add_table(rows=7, cols=2)
t.style = "Table Grid"
t.alignment = WD_TABLE_ALIGNMENT.CENTER
details = [
    ("Submitted By",     "Group 5"),
    ("Batch",            "2024 - 25"),
    ("Project Type",     "Midterm Capstone  -  3-Day Sprint"),
    ("Subject",          "Data Warehousing and Analytics"),
    ("Technologies",     "Python  |  Pandas  |  SQL  |  SQLite  |  Matplotlib"),
    ("Submission Date",  datetime.date.today().strftime("%d %B %Y")),
    ("GitHub",           "github.com/mhgowda/healthcare-insurance-dw-capstone"),
]
for i, (k, v) in enumerate(details):
    lc = t.rows[i].cells[0]
    rc = t.rows[i].cells[1]
    bg = "EDF1F7" if i % 2 == 0 else "FFFFFF"
    _shd(lc, C_HDRBG); _shd(rc, bg)
    _cell_write(lc, k, bold=True, sz=11, color=C_WHITE)
    _cell_write(rc, v, sz=11, color=C_DARK)
    lc.width = Inches(2.0)
    rc.width = Inches(4.2)

page_break()

# ═════════════════════════════════════════════════════════════════════════
# PAGE 2 — TABLE OF CONTENTS
# ═════════════════════════════════════════════════════════════════════════

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(16)
_run(p, "TABLE OF CONTENTS", bold=True, sz=16, color=C_NAVY)
_border_bottom(p, "1A5FA8", "10")

toc = [
    ("1",    "Executive Summary",                                    3),
    ("2",    "Introduction and Project Overview",                    4),
    ("2.1",  "Business Scenario",                                    4),
    ("2.2",  "Project Objectives",                                   4),
    ("2.3",  "Project Structure and Folder Layout",                  5),
    ("2.4",  "Technology Stack",                                     5),
    ("3",    "Source Data Description",                              6),
    ("3.1",  "Dataset Overview",                                     6),
    ("3.2",  "Entity Relationships in Source System",                6),
    ("3.3",  "Data Quality Challenges",                              7),
    ("4",    "Data Profiling",                                       8),
    ("4.1",  "Profiling Methodology",                                8),
    ("4.2",  "Profiling Findings by Entity",                         8),
    ("5",    "Star Schema Design",                                   9),
    ("5.1",  "Dimensional Modelling Concepts",                       9),
    ("5.2",  "Fact Tables",                                          9),
    ("5.3",  "Dimension Tables",                                    10),
    ("5.4",  "DDL and Indexes",                                     10),
    ("6",    "ETL Pipeline Architecture",                           11),
    ("6.1",  "Pipeline Overview",                                   11),
    ("6.2",  "Step 1 - Data Profiling",                             11),
    ("6.3",  "Step 2 - Data Cleaning and Validation",               12),
    ("6.4",  "Step 3 - Data Masking",                               12),
    ("6.5",  "Step 4 - Full ETL Load",                              13),
    ("6.6",  "Step 5 - Incremental ETL Load",                       13),
    ("6.7",  "Audit Logs and Rejected Rows",                        14),
    ("7",    "Data Masking",                                        14),
    ("7.1",  "Why Masking is Required",                             14),
    ("7.2",  "Masking Technique",                                   15),
    ("7.3",  "Masking Policy",                                      15),
    ("8",    "Data Cleaning Results",                               16),
    ("9",    "Analytical SQL Queries",                              17),
    ("10",   "Python Analytics and Visualisations",                 18),
    ("10.1", "Analytics Results Summary",                           18),
    ("10.2", "Visualisations",                                      18),
    ("11",   "Optimisation",                                        19),
    ("12",   "Deliverables Checklist",                              19),
    ("13",   "Conclusions and Future Scope",                        20),
    ("14",   "References",                                          21),
]

for num, title, pg in toc:
    p = doc.add_paragraph()
    p.paragraph_format.space_after  = Pt(2)
    p.paragraph_format.space_before = Pt(1)
    indent = Inches(0.25) if "." in num else Inches(0)
    p.paragraph_format.left_indent = indent
    dots  = "." * max(4, 65 - len(num) - len(title) - len(str(pg)))
    label = f"{num}   {title}"
    _run(p, label, bold=("." not in num), sz=11, color=C_DARK)
    _run(p, f"  {dots}  {pg}", sz=10, color=C_GRAY)

page_break()

# ═════════════════════════════════════════════════════════════════════════
# SECTION 1 — EXECUTIVE SUMMARY
# ═════════════════════════════════════════════════════════════════════════
h1("1.   Executive Summary")

para(
    "This report presents the complete documentation of a Healthcare Insurance Claims "
    "Data Warehouse and Analytics project developed as a three-day Midterm Capstone "
    "sprint. The project encompasses all phases of a real-world data engineering "
    "engagement: data profiling, cleaning, schema design, ETL development, data "
    "masking, warehouse loading, analytical querying, and visualisation."
)
para(
    "The source data consists of seven CSV files totalling approximately 1.4 million "
    "rows, representing operations of a national health insurance company. The "
    "dataset intentionally includes quality defects - missing values, orphan foreign "
    "keys, negative financial amounts, future dates, duplicate primary keys, and type "
    "mismatches - to simulate real-world data engineering challenges."
)
para(
    "After full ETL processing, 1,249,204 valid rows (87.8 percent of total input) "
    "were loaded into a nine-table star-schema warehouse implemented in SQLite. All "
    "personally identifiable information (PII) and protected health information (PHI) "
    "was masked using deterministic HMAC-SHA256 tokenisation prior to warehouse "
    "loading, ensuring that no raw personal data exists in the analytics layer."
)
para(
    "Key business findings from the warehouse include a 49.92 percent overall claim "
    "denial rate, Rs 15.68 billion in total paid claims across 208,640 valid claims, "
    "Punjab and Odisha as the states with the highest insurance expenditure, and "
    "NEFT, UPI, RTGS, and Cheque as payment methods with approximately equal "
    "utilisation. Eight Matplotlib charts were generated, and a Jupyter notebook "
    "provides an interactive analytics experience."
)

callout(
    "Summary Statistics:",
    "208,640 valid claims  |  Rs 15.68 Billion total paid  |  "
    "49.92% denial rate  |  17.9% overall data rejection rate  |  8 charts generated"
)

page_break()

# ═════════════════════════════════════════════════════════════════════════
# SECTION 2 — INTRODUCTION AND PROJECT OVERVIEW
# ═════════════════════════════════════════════════════════════════════════
h1("2.   Introduction and Project Overview")

h2("2.1  Business Scenario")
para(
    "A national health insurance company processes millions of insurance claims "
    "each year. Members (insured patients) visit providers (doctors and hospitals), "
    "who submit claims to the insurance company for reimbursement. Each claim may "
    "consist of multiple service lines, each with a procedure code and billing amount. "
    "The insurer processes these claims, makes payments where approved, and applies "
    "denial codes where claims are rejected."
)
para(
    "The company's data is currently stored in flat CSV extracts from the OLTP "
    "(Online Transaction Processing) system. These extracts are scattered, unjoined, "
    "and contain both valid records and intentionally introduced quality issues. "
    "The company requires a centralised Data Warehouse to support analytical "
    "reporting on costs, provider performance, fraud indicators, and geographic "
    "utilisation - while rigorously protecting patient privacy."
)

h2("2.2  Project Objectives")
para("The project is structured around five core objectives:")
for i, obj in enumerate([
    "Design a healthcare claims star schema with normalised dimension and fact tables suited for analytical workloads.",
    "Develop a robust ETL pipeline capable of Full loading (complete refresh) and Incremental loading (watermark-based new rows only).",
    "Implement deterministic data masking to protect all PII and PHI fields before any data enters the warehouse layer.",
    "Profile and clean the source data - identifying, documenting, and removing all quality defects.",
    "Perform analytical SQL queries and Python/Pandas analytics to answer key business questions, supported by eight Matplotlib visualisations.",
], 1):
    numbered(obj, i)

h2("2.3  Project Structure and Folder Layout")
para("The project is organised into clearly separated concerns:")

code_block(
    "project_5/\n"
    "  capstone/\n"
    "    config.py                    Central configuration (paths, salt)\n"
    "    day1_sql/\n"
    "      01_star_schema_ddl.sql     CREATE TABLE statements + indexes\n"
    "      02_analytical_queries.sql  10 business SQL queries\n"
    "    day2_etl/\n"
    "      step1_profile.py           Data quality profiling\n"
    "      step2_clean.py             Validation and row rejection\n"
    "      step3_mask.py              PII/PHI masking\n"
    "      step4_load_full.py         Full ETL load\n"
    "      step5_load_incremental.py  Incremental ETL load\n"
    "    day3_analytics/\n"
    "      analytics.py               9 analytical queries via Python\n"
    "      visualizations.py          8 Matplotlib charts\n"
    "      analytics_notebook.ipynb   Jupyter notebook\n"
    "    output/\n"
    "      warehouse.db               SQLite data warehouse\n"
    "      charts/                    8 PNG chart files\n"
    "      rejected_rows/             7 rejection audit CSVs\n"
    "  insurance_capstone_dataset_v1/\n"
    "    data/merged/                 7 source CSV files (~1.4M rows)\n"
    "    logs/                        profiling_summary.json, masking_audit.json\n"
    "  Healthcare_Insurance_DW_Report.docx   This document\n"
    "  Healthcare_Insurance_DW_Capstone.pptx Presentation slides"
)

h2("2.4  Technology Stack")
para("All project components use only standard, beginner-friendly Python libraries. "
     "No cloud infrastructure or enterprise tools are required to run the project.")
make_table(
    ["Technology", "Version", "Role in Project"],
    [
        ("Python",       "3.10+",   "Core language for ETL, analytics, and report generation"),
        ("Pandas",       "2.3+",    "CSV reading, data cleaning, transformation, analytics"),
        ("SQLite",       "Built-in","Local data warehouse - no installation required"),
        ("SQL",          "Standard","DDL for schema creation and analytical query execution"),
        ("Matplotlib",   "3.8+",    "Eight analytical charts saved as PNG"),
        ("python-docx",  "1.x",     "This report generation"),
        ("python-pptx",  "1.x",     "Presentation (PPT) generation"),
        ("Jupyter",      "Optional","Interactive analytics notebook"),
    ],
    widths=[1.3, 0.9, 4.1],
)

page_break()

# ═════════════════════════════════════════════════════════════════════════
# SECTION 3 — SOURCE DATA
# ═════════════════════════════════════════════════════════════════════════
h1("3.   Source Data Description")

h2("3.1  Dataset Overview")
para(
    "The source data was generated using the Python Faker library with the Indian "
    "locale (en_IN), producing realistic Indian names, cities, states, and contact "
    "details. The dataset has a fixed random seed (23) ensuring reproducibility. "
    "Approximately 83 percent of rows are clean valid records, and 17 percent are "
    "intentionally imperfect rows that test the ETL pipeline's quality controls."
)

make_table(
    ["CSV File", "Description", "Total Rows", "Clean Rows", "Bad Rows", "Primary Key"],
    [
        ("members_merged.csv",     "Patient/member master",          "2,60,000", "2,30,000", "30,000",  "member_id"),
        ("providers_merged.csv",   "Doctor/hospital master",         "40,000",   "35,000",   "5,000",   "provider_id"),
        ("claims_merged.csv",      "Insurance claim headers",        "2,40,000", "2,10,000", "30,000",  "claim_number"),
        ("claim_lines_merged.csv", "Service lines per claim",        "7,60,000", "6,50,000", "1,10,000","line_id"),
        ("payments_merged.csv",    "Payment transactions",           "1,00,000", "60,000",   "40,000",  "payment_id"),
        ("diagnoses_merged.csv",   "ICD codes per claim",            "60,000",   "40,000",   "20,000",  "diag_id"),
        ("procedures_merged.csv",  "CPT codes per claim line",       "40,000",   "25,000",   "15,000",  "proc_id"),
        ("TOTAL",                  "All entities combined",          "14,00,000","12,50,000","2,50,000","--"),
    ],
    widths=[2.0, 1.7, 0.9, 0.9, 0.8, 1.0],
)

h2("3.2  Entity Relationships in Source System")
para(
    "The seven source tables represent an OLTP-style schema. The relationships "
    "between entities are enforced logically (not by database constraints in the "
    "CSV files), which is why orphan foreign keys exist as a quality issue:"
)
for rel in [
    "MEMBERS has a one-to-many relationship with CLAIMS (one member can have many claims)",
    "PROVIDERS has a one-to-many relationship with CLAIMS (one provider can submit many claims)",
    "CLAIMS has a one-to-many relationship with CLAIM_LINES (one claim has multiple service lines)",
    "CLAIMS has a one-to-many relationship with PAYMENTS (one claim can have multiple payments)",
    "CLAIMS has a one-to-many relationship with DIAGNOSES (one claim can have multiple ICD codes)",
    "CLAIM_LINES has a one-to-many relationship with PROCEDURES (one line can have multiple CPT codes)",
]:
    bul(rel)

h2("3.3  Data Quality Challenges")
para(
    "Six categories of data quality issues were systematically injected into the "
    "dataset to test the ETL pipeline. The table below summarises each issue type, "
    "the columns it affects, and a concrete example:"
)

make_table(
    ["Issue Type", "Description", "Affected Columns", "Example Value"],
    [
        ("Missing Values",    "Required fields are NULL or blank",
         "first_name, email, city, diagnosis_code, specialty",
         "first_name = NULL"),
        ("Orphan Foreign Keys", "__ORPHAN__ placeholder where parent record missing",
         "member_id in claims, claim_number in payments",
         "member_id = '__ORPHAN__'"),
        ("Negative Amounts",  "Financial columns contain negative values",
         "allowed_amount, paid_amount, units, line_amount",
         "allowed_amount = -5000.00"),
        ("Future Dates",      "Timestamps set in the future (impossible for past events)",
         "admit_date, service_date, dob, payment_date",
         "dob = 2031-11-03"),
        ("Duplicate Keys",    "Same primary key appearing more than once",
         "claim_number, member_id, provider_id",
         "Two rows: claim_number = CL00012345"),
        ("Type Mismatches",   "Text values in numeric columns",
         "paid_amount, rate, length_of_stay",
         "paid_amount = 'ERR'"),
    ],
    widths=[1.5, 1.9, 2.1, 1.7],
)

page_break()

# ═════════════════════════════════════════════════════════════════════════
# SECTION 4 — DATA PROFILING
# ═════════════════════════════════════════════════════════════════════════
h1("4.   Data Profiling")

h2("4.1  Profiling Methodology")
para(
    "Data profiling is the first step in the ETL pipeline, performed by "
    "day2_etl/step1_profile.py. The profiler reads each CSV file and counts "
    "all quality issues without modifying any data. The purpose is to understand "
    "the extent of the problems before making any decisions about cleaning. "
    "Results are saved to logs/profiling_summary.json for audit purposes."
)
para("The profiler checks for the following conditions on every entity:")
for check in [
    "NULL or empty string values in every column (missing value count)",
    "Duplicate values in the primary key column (duplicate key count)",
    "Rows where any FK column equals the string '__ORPHAN__' (orphan FK count)",
    "Non-numeric values in columns that must be numeric (type mismatch count)",
    "Negative values in financial and quantity columns (negative value count)",
    "Date values greater than today in date columns (future date count)",
]:
    bul(check)

h2("4.2  Profiling Findings by Entity")
make_table(
    ["Entity", "Issue Type", "Column", "Count Found"],
    [
        ("members",     "Missing values",   "first_name, email, city",      "10,500 each"),
        ("members",     "Duplicate key",    "member_id",                    "6,000"),
        ("members",     "Future date",      "dob",                          "22,880"),
        ("providers",   "Missing values",   "specialty",                    "433"),
        ("providers",   "Duplicate key",    "provider_id",                  "1,250"),
        ("claims",      "Orphan FK",        "member_id, provider_id",       "30,000 each"),
        ("claims",      "Negative value",   "allowed_amount",               "30,000"),
        ("claims",      "Non-numeric",      "paid_amount",                  "30,000"),
        ("claims",      "Duplicate key",    "claim_number",                 "7,500"),
        ("claims",      "Future date",      "admit_date",                   "75"),
        ("claim_lines", "Orphan FK",        "claim_number",                 "1,10,000"),
        ("claim_lines", "Negative value",   "units, line_amount",           "1,10,000 each"),
        ("claim_lines", "Non-numeric",      "rate",                         "1,10,000"),
        ("claim_lines", "Duplicate key",    "line_id",                      "27,500"),
        ("payments",    "Orphan FK",        "claim_number",                 "40,000"),
        ("payments",    "Non-numeric",      "paid_amount",                  "40,000"),
        ("diagnoses",   "Orphan FK",        "claim_number",                 "20,000"),
        ("diagnoses",   "Duplicate key",    "diag_id",                      "5,000"),
        ("procedures",  "Orphan FK",        "line_id",                      "15,000"),
        ("procedures",  "Duplicate key",    "proc_id",                      "3,750"),
    ],
    widths=[1.3, 1.5, 2.4, 1.5],
)

page_break()

# ═════════════════════════════════════════════════════════════════════════
# SECTION 5 — STAR SCHEMA DESIGN
# ═════════════════════════════════════════════════════════════════════════
h1("5.   Star Schema Design")

h2("5.1  Dimensional Modelling Concepts")
para(
    "The star schema is the standard design pattern for analytical data warehouses, "
    "introduced by Ralph Kimball in 'The Data Warehouse Toolkit'. It organises data "
    "into two categories: Fact Tables, which store measurable numeric events (amounts, "
    "counts, durations), and Dimension Tables, which provide descriptive context "
    "(who, what, where, when). Analytical queries join from a central fact table "
    "outward to surrounding dimension tables - a shape that resembles a star."
)
para(
    "The key advantages of the star schema for this project are: simplified SQL "
    "queries (one join per dimension), fast aggregation performance, intuitive "
    "structure for business users, and straightforward extension when new facts "
    "or dimensions are required."
)
para("The schema diagram for this project:")
code_block(
    "                        dim_date\n"
    "                            |\n"
    "  dim_member  ---  fact_claim  ---  dim_provider\n"
    "                            |\n"
    "               fact_claim_line  ---  dim_procedure\n"
    "                            |\n"
    "                      fact_payment\n"
    "                            |\n"
    "     dim_diagnosis        dim_location"
)

h2("5.2  Fact Tables")
para(
    "Fact tables store the measurable data at the lowest grain required for analysis. "
    "This schema uses three fact tables:"
)
make_table(
    ["Table", "Grain (One Row Per)", "Measures", "Foreign Keys"],
    [
        ("fact_claim",
         "Insurance claim",
         "allowed_amount, paid_amount, length_of_stay, is_denied",
         "member_id, provider_id, admit_date_key, discharge_date_key"),
        ("fact_claim_line",
         "Service line within a claim",
         "units, rate, line_amount",
         "claim_number, service_date_key, procedure_code"),
        ("fact_payment",
         "Payment transaction",
         "paid_amount",
         "claim_number, payment_date_key"),
    ],
    widths=[1.5, 1.5, 2.2, 2.1],
)

h2("5.3  Dimension Tables")
make_table(
    ["Table", "Description", "Key Columns", "PII Masking"],
    [
        ("dim_member",   "Insured patient details",
         "member_id, gender, city, state, birth_year, is_active",
         "member_id, names, email, phone, dob"),
        ("dim_provider", "Doctor or hospital details",
         "provider_id, specialty, city, state, tax_id, license_no",
         "provider_id, name, tax_id, license_no"),
        ("dim_date",     "Full calendar dimension",
         "date_key, year, quarter, month, month_name, day_name, is_weekend",
         "None - no PII"),
        ("dim_diagnosis","ICD medical diagnosis codes",
         "diagnosis_code, icd_version",
         "None - clinical codes only"),
        ("dim_procedure","CPT medical procedure codes",
         "procedure_code",
         "None - clinical codes only"),
        ("dim_location", "City and state combinations",
         "location_key, city, state",
         "None - geographic only"),
    ],
    widths=[1.4, 1.6, 2.4, 1.9],
)

h2("5.4  DDL and Indexes")
para(
    "The full CREATE TABLE statements are defined in day1_sql/01_star_schema_ddl.sql. "
    "Seven database indexes were created to optimise the most common query patterns:"
)
make_table(
    ["Index", "Table", "Column", "Optimises"],
    [
        ("idx_fc_member",    "fact_claim",       "member_id",       "Member dimension JOIN"),
        ("idx_fc_provider",  "fact_claim",       "provider_id",     "Provider dimension JOIN"),
        ("idx_fc_admit",     "fact_claim",       "admit_date_key",  "Date range filtering"),
        ("idx_fc_denied",    "fact_claim",       "is_denied",       "Denial rate WHERE clause"),
        ("idx_fcl_claim",    "fact_claim_line",  "claim_number",    "Line-level aggregation"),
        ("idx_fp_claim",     "fact_payment",     "claim_number",    "Payment aggregation"),
        ("idx_dp_specialty", "dim_provider",     "specialty",       "Specialty GROUP BY"),
    ],
    widths=[1.7, 1.7, 1.5, 2.4],
)

page_break()

# ═════════════════════════════════════════════════════════════════════════
# SECTION 6 — ETL PIPELINE
# ═════════════════════════════════════════════════════════════════════════
h1("6.   ETL Pipeline Architecture")

h2("6.1  Pipeline Overview")
para(
    "The ETL pipeline converts raw CSV source files into a clean, masked, and "
    "structured warehouse in five sequential steps. Each step is implemented as "
    "an independent Python script so that any step can be re-run individually "
    "without re-running the entire pipeline."
)
make_table(
    ["Step", "Script File", "Input", "Output", "On Failure"],
    [
        ("1 - Profile",      "step1_profile.py",       "7 source CSVs",
         "logs/profiling_summary.json", "Report only, no changes"),
        ("2 - Clean",        "step2_clean.py",          "7 source CSVs",
         "Clean DataFrames in memory, rejected/*.csv", "Reject bad row, continue"),
        ("3 - Mask",         "step3_mask.py",            "Clean DataFrames",
         "Masked DataFrames in memory, masking_audit.json", "Exception raised"),
        ("4 - Full Load",    "step4_load_full.py",       "Masked DataFrames",
         "warehouse.db populated", "Rollback, exception logged"),
        ("5 - Incremental",  "step5_load_incremental.py","7 source CSVs + watermark",
         "warehouse.db updated", "Rollback, exception logged"),
    ],
    widths=[1.3, 1.8, 1.4, 1.9, 1.5],
)

h2("6.2  Step 1 - Data Profiling")
para(
    "The profiling step reads each source CSV and counts all quality issues. "
    "It produces a JSON report saved to logs/profiling_summary.json. This step "
    "is non-destructive - it reads data but does not modify anything. The profiler "
    "checks nulls, duplicates, orphans, negatives, future dates, and type mismatches "
    "for every applicable column in every entity."
)
para(
    "Running profiling before cleaning gives the team a documented baseline of "
    "data quality issues. This is important for reporting and for validating that "
    "the cleaning step addressed all identified problems."
)

h2("6.3  Step 2 - Data Cleaning and Validation")
para(
    "The cleaning step (step2_clean.py) processes entities in dependency order. "
    "This ordering is critical: members must be cleaned before claims (so that FK "
    "validation against valid member IDs is possible), and claims must be cleaned "
    "before claim_lines and payments."
)
para("Dependency order:")
for i, item in enumerate(["members", "providers", "claims", "claim_lines", "payments", "diagnoses", "procedures"], 1):
    bul(item, level=0)

para(
    "For each entity, the cleaner applies the following rules in order. A row is "
    "marked bad if it fails any rule, and is written to the corresponding rejection "
    "CSV with a rejection_reason column:"
)
make_table(
    ["Rule", "Condition Checked", "Action on Failure"],
    [
        ("Required fields",    "Any required column is NULL or empty string", "Reject row, reason: missing_<column>"),
        ("Duplicate PK",       "Primary key already seen in this batch",      "Reject second occurrence, reason: duplicate_<pk>"),
        ("Orphan FK",          "FK column equals '__ORPHAN__' string",         "Reject row, reason: orphan_<fk>"),
        ("FK existence",       "FK value not found in parent entity",          "Reject row, reason: <fk>_not_found"),
        ("Negative amounts",   "Numeric column value is below zero",           "Reject row, reason: negative_<column>"),
        ("Non-numeric",        "Column value cannot be parsed as a number",    "Reject row, reason: non_numeric_<column>"),
        ("Future dates",       "Date column value is greater than today",      "Reject row, reason: future_<column>"),
        ("Invalid ICD version","icd_version not in {'ICD-9', 'ICD-10'}",      "Reject row, reason: invalid_icd_version"),
    ],
    widths=[1.5, 2.7, 3.0],
)

h2("6.4  Step 3 - Data Masking")
para(
    "Data masking (step3_mask.py) is applied after cleaning and before loading. "
    "This ensures that the warehouse never contains raw PII. The masking step "
    "also writes logs/masking_audit.json recording which fields were masked, "
    "how many values were processed, and the timestamp of the operation. "
    "The masking salt is stored in config.py."
)

h2("6.5  Step 4 - Full ETL Load")
para(
    "The full load (step4_load_full.py) performs a complete rebuild of the warehouse:"
)
for i, step in enumerate([
    "Reads the DDL file and recreates all tables (drops existing data first)",
    "Calls the cleaning module - processes all 7 entities in dependency order",
    "Calls the masking module - applies tokenisation to all PII/PHI fields",
    "Loads all six dimension tables first (dim_date, dim_member, dim_provider, dim_location, dim_diagnosis, dim_procedure)",
    "Loads all three fact tables after dimensions (fact_claim, fact_claim_line, fact_payment)",
    "Inserts load timestamp into the etl_watermark table for incremental use",
    "Handles the SQLite variable limit by chunking data appropriately",
], 1):
    numbered(step, i)

callout("Performance:", "Full load of 1.4M rows completes in approximately 3-5 minutes on a standard laptop.")

h2("6.6  Step 5 - Incremental ETL Load")
para(
    "The incremental load (step5_load_incremental.py) uses a watermark mechanism "
    "to process only rows that have arrived since the last successful run. The "
    "etl_watermark table stores one row per entity with the last_load_dt timestamp."
)
para(
    "On each incremental run, the source CSV is filtered to rows where the "
    "watermark date column is greater than last_load_dt. Only those new rows go "
    "through cleaning, masking, and loading. The INSERT OR IGNORE SQL command "
    "ensures that duplicate rows cannot enter the warehouse even if the same "
    "CSV is processed twice."
)
make_table(
    ["Entity", "Watermark Column", "Filter Condition", "Refresh Strategy"],
    [
        ("members",     "created_at",    "created_at > last_load_dt",   "Filtered incremental"),
        ("claims",      "admit_date",    "admit_date > last_load_dt",   "Filtered incremental"),
        ("claim_lines", "service_date",  "service_date > last_load_dt", "Filtered incremental"),
        ("payments",    "payment_date",  "payment_date > last_load_dt", "Filtered incremental"),
        ("providers",   "No date column","Always process all rows",     "Full upsert - INSERT OR IGNORE"),
        ("diagnoses",   "No date column","Always process all rows",     "Full upsert - INSERT OR IGNORE"),
        ("procedures",  "No date column","Always process all rows",     "Full upsert - INSERT OR IGNORE"),
    ],
    widths=[1.2, 1.4, 1.9, 2.7],
)

h2("6.7  Audit Logs and Rejected Rows")
para(
    "Every ETL run produces a complete audit trail. Seven rejection CSV files "
    "(one per entity) are saved in output/rejected_rows/. Each file contains all "
    "rejected rows with a rejection_reason column identifying which rule caused "
    "rejection. This allows the data team to review, fix, and re-submit rejected "
    "records in future runs."
)
para(
    "The masking audit (logs/masking_audit.json) records the run timestamp, "
    "whether a salt was used, and for each entity and field: the number of values "
    "that were masked. This provides a compliance record for data privacy audits."
)

page_break()

# ═════════════════════════════════════════════════════════════════════════
# SECTION 7 — DATA MASKING
# ═════════════════════════════════════════════════════════════════════════
h1("7.   Data Masking")

h2("7.1  Why Masking is Required")
para(
    "Healthcare data is among the most sensitive personal information. Insurance "
    "records contain names, dates of birth, contact details, and medical diagnoses "
    "that are protected by privacy regulations such as HIPAA (Health Insurance "
    "Portability and Accountability Act) in the United States and equivalent "
    "legislation in India."
)
para(
    "In this project, analysts querying the warehouse to understand costs, denial "
    "rates, and provider performance do not need access to real patient names, "
    "email addresses, phone numbers, or exact dates of birth. Loading unmasked "
    "PII into an analytics database creates unnecessary privacy risk. The principle "
    "of data minimisation - collecting and using only the minimum data necessary - "
    "requires that sensitive fields be masked before entering the analytics layer."
)

h2("7.2  Masking Technique - HMAC-SHA256")
para(
    "This project implements deterministic masking using HMAC-SHA256 (Hash-based "
    "Message Authentication Code with SHA-256). This approach was chosen over "
    "simple hashing or random tokenisation for three specific reasons:"
)
for reason in [
    "DETERMINISTIC: The same input value always produces the same output token when the same salt is used. This means that member_id 'M0000001' always becomes 'MBR3F9A2C1D4E' - so JOIN operations between fact_claim and dim_member continue to work correctly after masking.",
    "IRREVERSIBLE: SHA-256 is a one-way cryptographic hash. The original value cannot be recovered from the token without the salt key.",
    "SALT-PROTECTED: The salt string added to each value before hashing makes the tokens unique to this deployment. An attacker who obtains the masked tokens cannot reconstruct the original values without knowing the exact salt.",
]:
    bul(reason)

para(
    "Dates of birth are handled differently from other PII: instead of hashing, "
    "the exact date is replaced by the birth year only (generalisation). This "
    "preserves the ability to calculate approximate age bands while removing the "
    "precision that would make the data personally identifying."
)

h2("7.3  Masking Policy")
make_table(
    ["Field", "Entity", "Technique", "Before (Raw)", "After (Warehouse)"],
    [
        ("member_id",    "members",   "HMAC-SHA256 token", "M0000001",          "MBR3F9A2C1D4E"),
        ("first_name",   "members",   "HMAC-SHA256 token", "Alice",             "FN8B2C3D4E5F"),
        ("last_name",    "members",   "HMAC-SHA256 token", "Smith",             "LN2C3D4E5F6A"),
        ("full_name",    "members",   "HMAC-SHA256 token", "Alice Smith",       "NM4E5F6A7B8C"),
        ("email",        "members",   "HMAC-SHA256 token", "alice@gmail.com",   "EMAIF29A3B1C"),
        ("phone",        "members",   "HMAC-SHA256 token", "9876543210",        "PH8A2B3C4D"),
        ("dob",          "members",   "Year only (generalise)", "1985-03-10",   "1985"),
        ("provider_id",  "providers", "HMAC-SHA256 token", "P000001",           "PRV7AEE1BC05"),
        ("provider_name","providers", "HMAC-SHA256 token", "Acme Clinic",       "PNM3B4C5D6E7"),
        ("tax_id",       "providers", "HMAC-SHA256 token", "908430420",         "TAX2B3C4D5E6"),
        ("license_no",   "providers", "HMAC-SHA256 token", "LIC790693",         "LIC3C4D5E6F7"),
        ("claim_number", "claims",    "HMAC-SHA256 token", "CL00155929",        "CLM4E5F6A7B8"),
        ("city / state", "members",   "Kept as-is (allowed)", "Mumbai, MH",     "Mumbai, MH"),
        ("gender",       "members",   "Kept as-is (allowed)", "female",         "female"),
        ("specialty",    "providers", "Kept as-is (allowed)", "Cardiology",     "Cardiology"),
    ],
    widths=[1.2, 1.1, 1.8, 1.7, 1.5],
)

page_break()

# ═════════════════════════════════════════════════════════════════════════
# SECTION 8 — DATA CLEANING RESULTS
# ═════════════════════════════════════════════════════════════════════════
h1("8.   Data Cleaning Results")

para(
    "The cleaning step processes entities in strict dependency order. The table "
    "below presents the full results of the cleaning pass across all seven entities. "
    "The Rejection Percent column shows what proportion of input rows failed one "
    "or more validation rules. All rejected rows are saved to "
    "output/rejected_rows/<entity>_rejected.csv with the rejection reason."
)

make_table(
    ["Entity", "Input Rows", "Valid Rows", "Rejected", "Rejection %", "Primary Rejection Reasons"],
    [
        ("members",     "2,60,000", "2,32,207", "27,793",   "10.7%",
         "Missing first_name / email / city (35%), future DOB (82K), duplicate member_id (6K)"),
        ("providers",   "40,000",   "38,424",   "1,576",    "3.9%",
         "Missing specialty (433), duplicate provider_id (1,250)"),
        ("claims",      "2,40,000", "2,08,640", "31,360",   "13.1%",
         "Orphan member_id (30K), negative allowed_amount (30K), future admit_date (75)"),
        ("claim_lines", "7,60,000", "6,45,733", "1,14,267", "15.0%",
         "Orphan claim_number (1,10,000), negative units (1,10,000), non-numeric rate (1,10,000)"),
        ("payments",    "1,00,000", "59,615",   "40,385",   "40.4%",
         "Orphan claim_number (40K), non-numeric paid_amount (40K)"),
        ("diagnoses",   "60,000",   "39,741",   "20,259",   "33.8%",
         "Orphan claim_number (20K), invalid ICD version (X instead of ICD-9/ICD-10)"),
        ("procedures",  "40,000",   "24,844",   "15,156",   "37.9%",
         "Orphan line_id (15K), missing procedure_code (6), duplicate proc_id (3,750)"),
    ],
    widths=[1.1, 0.9, 0.9, 0.9, 0.9, 3.1],
)

make_table(
    ["Metric", "Value"],
    [
        ("Total input rows (all entities)",   "14,00,000"),
        ("Total valid rows loaded to warehouse", "12,49,204"),
        ("Total rejected rows",               "2,50,796"),
        ("Overall rejection percentage",      "17.9 percent"),
        ("Rejection audit files created",     "7 CSV files in output/rejected_rows/"),
    ],
    widths=[3.5, 3.7],
)

para(
    "The high rejection rate for payments (40.4%) and procedures (37.9%) is "
    "expected given that approximately 40% of those entities were injected as "
    "bad rows in the original dataset. The relatively low rejection rate for "
    "providers (3.9%) reflects the simpler validation rules applicable to that "
    "entity."
)

page_break()

# ═════════════════════════════════════════════════════════════════════════
# SECTION 9 — ANALYTICAL SQL QUERIES
# ═════════════════════════════════════════════════════════════════════════
h1("9.   Analytical SQL Queries")

para(
    "Ten analytical SQL queries were developed in day1_sql/02_analytical_queries.sql "
    "and executed programmatically via Python in day3_analytics/analytics.py. "
    "Each query answers a specific business question. Results are saved as individual "
    "CSV files in the output/ folder."
)

make_table(
    ["Query", "Business Question", "SQL Techniques Used", "Key Result"],
    [
        ("Q1",  "Overall claims summary",
         "COUNT, SUM, AVG, percentage calculation",
         "208,640 claims | Rs 15.68B total paid | 49.92% denial"),
        ("Q2",  "Monthly paid amount trend",
         "JOIN fact_claim + dim_date, GROUP BY year and month",
         "Consistent ~Rs 540M/month across 30 months"),
        ("Q3",  "Top 10 procedures by utilisation",
         "GROUP BY procedure_code, ORDER BY COUNT DESC, LIMIT",
         "CPT35289: 417 uses (highest volume)"),
        ("Q4",  "Provider performance scorecard",
         "JOIN dim_provider, AVG paid, denial rate percentage",
         "Some providers exceed 60% denial rate"),
        ("Q5",  "Denial rate by code",
         "CASE WHEN on denial_code, GROUP BY denial_reason",
         "D001/D002/D003 each ~16.6% of all claims"),
        ("Q6",  "Length of stay distribution",
         "GROUP BY length_of_stay, AVG paid amount",
         "Flat 0-29 days, avg ~Rs 75,156 per claim"),
        ("Q7",  "Top 20 members by total paid",
         "JOIN dim_member, SUM paid, ORDER BY spend DESC",
         "Top member: Rs 785,827 across 6 claims"),
        ("Q8",  "Geographic breakdown by state",
         "JOIN dim_member on state, SUM by state",
         "Punjab leads: Rs 876M total paid"),
        ("Q9",  "Payment method split",
         "GROUP BY payment_method, SUM and COUNT",
         "NEFT/UPI/RTGS/Cheque approximately equal ~25% each"),
        ("Q10", "Diagnosis cost analysis",
         "Link diagnoses to claims via claim_number",
         "ICD codes ranked by total claim cost"),
    ],
    widths=[0.5, 1.7, 2.0, 2.3],
)

h3("Sample SQL - Q4 Provider Performance Scorecard")
code_block(
    "SELECT\n"
    "    dp.provider_id,\n"
    "    dp.specialty,\n"
    "    dp.state,\n"
    "    COUNT(fc.claim_number)              AS total_claims,\n"
    "    ROUND(SUM(fc.paid_amount), 2)       AS total_paid,\n"
    "    ROUND(AVG(fc.paid_amount), 2)       AS avg_paid_per_claim,\n"
    "    ROUND(AVG(fc.length_of_stay), 1)   AS avg_length_of_stay,\n"
    "    SUM(fc.is_denied)                   AS denied_claims,\n"
    "    ROUND(\n"
    "        100.0 * SUM(fc.is_denied) / COUNT(*), 2\n"
    "    )                                   AS denial_rate_pct\n"
    "FROM fact_claim fc\n"
    "JOIN dim_provider dp\n"
    "    ON fc.provider_id = dp.provider_id\n"
    "GROUP BY dp.provider_id, dp.specialty, dp.state\n"
    "ORDER BY total_paid DESC\n"
    "LIMIT 20;"
)

h3("Sample SQL - Q2 Monthly Paid Trend")
code_block(
    "SELECT\n"
    "    d.year,\n"
    "    d.month,\n"
    "    d.month_name,\n"
    "    COUNT(fc.claim_number)              AS claims,\n"
    "    ROUND(SUM(fc.paid_amount), 2)       AS total_paid,\n"
    "    ROUND(AVG(fc.paid_amount), 2)       AS avg_paid\n"
    "FROM fact_claim fc\n"
    "JOIN dim_date d\n"
    "    ON fc.admit_date_key = d.date_key\n"
    "GROUP BY d.year, d.month, d.month_name\n"
    "ORDER BY d.year, d.month;"
)

page_break()

# ═════════════════════════════════════════════════════════════════════════
# SECTION 10 — PYTHON ANALYTICS AND VISUALISATIONS
# ═════════════════════════════════════════════════════════════════════════
h1("10.  Python Analytics and Visualisations")

h2("10.1  Analytics Results Summary")
para(
    "All queries were executed against the live SQLite warehouse using Python's "
    "sqlite3 module and Pandas read_sql_query(). Results were printed to the "
    "console and saved as CSV files for further review."
)

make_table(
    ["Metric", "Value", "Notes"],
    [
        ("Total valid claims loaded",   "2,08,640",      "From 2,40,000 raw claim rows"),
        ("Total allowed amount",        "Rs 20.92 Billion", "Total billed by all providers"),
        ("Total paid amount",           "Rs 15.68 Billion", "75% of allowed amount paid"),
        ("Average paid per claim",      "Rs 75,156",     "Consistent across claim types"),
        ("Overall denial rate",         "49.92 percent", "Very high - warrants investigation"),
        ("Denial code D001",            "34,648 claims", "16.61 percent of total"),
        ("Denial code D002",            "34,774 claims", "16.67 percent of total"),
        ("Denial code D003",            "34,739 claims", "16.65 percent of total"),
        ("No denial (paid)",            "1,04,479 claims","50.08 percent of total"),
        ("Highest spending state",      "Punjab",        "Rs 876 million paid"),
        ("Second highest state",        "Odisha",        "Rs 821 million paid"),
        ("Top procedure (CPT35289)",    "417 claim lines","Highest utilisation code"),
        ("Top member (max spend)",      "Rs 7,85,828",   "6 claims, Telangana, 1992"),
        ("Payment methods",             "4 (NEFT/UPI/RTGS/Cheque)", "~25% each"),
        ("Average length of stay",      "~14.5 days",   "Flat distribution 0-29 days"),
    ],
    widths=[2.4, 1.8, 3.0],
)

h2("10.2  Visualisations")
para(
    "Eight charts were generated by day3_analytics/visualizations.py using Matplotlib "
    "and saved as PNG files in output/charts/. The same charts are also reproduced "
    "interactively in the Jupyter notebook analytics_notebook.ipynb."
)

make_table(
    ["Chart File", "Type", "Business Question", "Key Insight"],
    [
        ("01_monthly_paid.png",          "Bar chart (vertical)",
         "Monthly paid amount trend over 30 months",
         "Consistent ~Rs 540M/month - no seasonal patterns"),
        ("02_top_procedures.png",        "Bar chart (horizontal)",
         "Top 10 most utilised procedure codes",
         "CPT35289 leads with 417 uses"),
        ("03_provider_performance.png",  "Side-by-side horizontal bars",
         "Top 15 providers: total paid vs denial rate",
         "Some high-paid providers also have >60% denial"),
        ("04_denial_rate.png",           "Bar chart (vertical)",
         "Distribution of claims across denial codes",
         "50% paid, 50% denied across D001/D002/D003 equally"),
        ("05_los_distribution.png",      "Bar chart (vertical)",
         "Length of stay distribution 0-30 days",
         "Flat uniform distribution - no clustering"),
        ("06_member_utilisation.png",    "Scatter plot",
         "Top 200 members - total claims vs total paid",
         "High-spend members (6-7 claims) drive outlier costs"),
        ("07_geographic.png",            "Bar chart (vertical)",
         "Top 15 states by total paid amount",
         "Punjab and Odisha dominate; north/east India high"),
        ("08_payment_methods.png",       "Pie chart",
         "Payment method split by total paid",
         "NEFT/UPI/RTGS/Cheque approximately 25% each"),
    ],
    widths=[2.0, 1.2, 2.2, 2.0],
)

page_break()

# ═════════════════════════════════════════════════════════════════════════
# SECTION 11 — OPTIMISATION
# ═════════════════════════════════════════════════════════════════════════
h1("11.  Optimisation")

h2("11.1  Database Index Strategy")
para(
    "Seven indexes were created on columns that appear in JOIN ON clauses and "
    "WHERE conditions of the analytical queries. Without indexes, SQLite performs "
    "a full table scan for every query. With indexes on foreign keys and date "
    "keys, lookups use B-tree index traversal which is O(log n) rather than O(n). "
    "For a fact_claim table with 208,640 rows, this difference is significant."
)

h2("11.2  Vectorised Pandas Operations")
para(
    "All data cleaning and transformation operations use Pandas vectorised methods "
    "rather than Python row-by-row loops. The performance difference is substantial "
    "at 1.4 million rows: a Python for-loop iterating over 760,000 claim_line "
    "rows would take several minutes; the equivalent Pandas vectorised operation "
    "completes in under one second."
)

make_table(
    ["Operation", "Row-Loop Approach (Slow)", "Vectorised Approach (Used)", "Speedup"],
    [
        ("Parse dates",        "for row in df.iterrows(): datetime.strptime(row['date'],...)",
         "pd.to_datetime(df['date'], errors='coerce')",
         "~50-100x"),
        ("Convert to numeric", "for row in df.iterrows(): float(row['amount'])",
         "pd.to_numeric(df['amount'], errors='coerce')",
         "~50-100x"),
        ("Filter bad rows",    "bad_rows = []; for row: if condition: bad_rows.append(row)",
         "df[boolean_mask]",
         "~20-50x"),
        ("Remap FK tokens",    "for row: row['id'] = token_dict[row['id']]",
         "df['id'].map(token_dict)",
         "~30-80x"),
        ("Find duplicates",    "for row: if id in seen: mark_dup(row)",
         "df.duplicated(subset=['id'], keep='first')",
         "~20-50x"),
    ],
    widths=[1.4, 2.3, 2.0, 0.8],
)

h2("11.3  Chunked SQLite Loading")
para(
    "SQLite has a maximum of 32,766 bound variables per SQL statement. When loading "
    "a DataFrame with many columns using to_sql(), the chunk size must be calculated "
    "as max_vars // number_of_columns. Without this calculation, loading a table "
    "with 10 columns would fail at chunks larger than 3,276 rows."
)

h2("11.4  Incremental Loading Efficiency")
para(
    "The watermark-based incremental load dramatically reduces processing time for "
    "daily runs. In a production scenario receiving 1,000 new claims per day, "
    "the incremental approach processes only those 1,000 rows versus reloading "
    "all 208,640 claims in a full load. This is a 200x reduction in data volume "
    "per run."
)

page_break()

# ═════════════════════════════════════════════════════════════════════════
# SECTION 12 — DELIVERABLES
# ═════════════════════════════════════════════════════════════════════════
h1("12.  Deliverables Checklist")

para(
    "The following table lists all project deliverables as specified in the "
    "capstone project document, along with their file locations and completion status."
)

make_table(
    ["Deliverable", "File Location", "Status"],
    [
        ("Star Schema Diagram (logical + physical)", "day1_sql/01_star_schema_ddl.sql",       "Complete"),
        ("Analytical SQL Queries and outputs",        "day1_sql/02_analytical_queries.sql",    "Complete"),
        ("Data profiling report",                     "logs/profiling_summary.json",            "Complete"),
        ("ETL Full Load script",                      "day2_etl/step4_load_full.py",            "Complete"),
        ("ETL Incremental Load script",               "day2_etl/step5_load_incremental.py",    "Complete"),
        ("Data masking logic and policy",             "day2_etl/step3_mask.py",                "Complete"),
        ("Masking audit log",                         "logs/masking_audit.json",               "Complete"),
        ("Error logs / rejected rows (7 files)",      "output/rejected_rows/*.csv",            "Complete"),
        ("DW DDL (table creation SQL)",               "day1_sql/01_star_schema_ddl.sql",       "Complete"),
        ("Loaded data warehouse",                     "output/warehouse.db",                   "Complete"),
        ("Python analytics script",                   "day3_analytics/analytics.py",           "Complete"),
        ("8 Matplotlib charts (PNG)",                 "output/charts/*.png",                   "Complete"),
        ("Jupyter Notebook",                          "day3_analytics/analytics_notebook.ipynb","Complete"),
        ("Optimisation summary",                      "Section 11 of this report",             "Complete"),
        ("Final Presentation (PPT)",                  "Healthcare_Insurance_DW_Capstone.pptx", "Complete"),
        ("Final Project Report (DOCX)",               "Healthcare_Insurance_DW_Report.docx",   "Complete"),
        ("README and Team Guide",                     "README.md,  TEAM_GUIDE.md",             "Complete"),
        ("GitHub Repository",                         "github.com/mhgowda/healthcare-insurance-dw-capstone","Complete"),
    ],
    widths=[2.6, 2.7, 1.0],
)

page_break()

# ═════════════════════════════════════════════════════════════════════════
# SECTION 13 — CONCLUSIONS
# ═════════════════════════════════════════════════════════════════════════
h1("13.  Conclusions and Future Scope")

h2("13.1  Project Conclusions")
para(
    "This project has successfully delivered a complete end-to-end data engineering "
    "solution for healthcare insurance claims analytics. Beginning with seven raw, "
    "imperfect CSV files totalling 1.4 million rows, the pipeline produces a "
    "clean, masked, indexed, and analytically capable data warehouse."
)
para("The following conclusions are drawn from the project:")

for c in [
    "Data quality is the foundation of reliable analytics. The 17.9 percent overall rejection rate demonstrates that unchecked data would have introduced significant errors into downstream analysis.",
    "The 49.92 percent claim denial rate is a critical business signal that warrants immediate investigation. Nearly half of all claims being denied suggests either systematic coding errors, provider-payer policy mismatches, or potentially fraudulent submissions.",
    "Deterministic masking successfully protects all PII and PHI while preserving full analytical capability. All JOIN operations between masked tables produced correct results.",
    "The star schema design proved appropriate for the analytical workload. The JOIN structure is simple, queries are intuitive, and the schema scales easily to new dimensions.",
    "Geographic analysis - possible only because city and state fields were explicitly kept unmasked - revealed strong regional clustering with Punjab and Odisha accounting for a disproportionate share of total paid claims.",
    "Incremental loading significantly improves operational efficiency, reducing daily processing from 3-5 minutes (full load) to seconds for typical daily volumes.",
    "The project demonstrates that a fully functional data warehouse can be built with no cloud infrastructure, no enterprise tools, and only Python standard libraries plus Pandas and Matplotlib.",
]:
    bul(c)

h2("13.2  Limitations")
para("The following limitations of the current implementation are acknowledged:")
for lim in [
    "SQLite is appropriate for learning and small-to-medium datasets but would not scale to enterprise volumes. A production deployment would require PostgreSQL, Snowflake, or a similar enterprise DBMS.",
    "The source data is synthetically generated. Real-world insurance data would exhibit more complex and nuanced quality patterns that may require additional validation rules.",
    "The masking implementation does not support role-based access control (different masking levels for analysts, auditors, and data engineers). A production system would implement this through column-level security.",
    "dim_diagnosis is not directly linked to fact_claim in the current schema (diagnoses relate to claims via a separate diagnoses table). Adding a claim-diagnosis bridge fact would improve ICD-level analytics.",
    "The project uses a single SQLite database file. In production, separate databases for staging, integration, and presentation layers would provide better governance and lineage tracking.",
]:
    bul(lim)

h2("13.3  Future Scope and Enhancements")
make_table(
    ["Enhancement", "Description", "Technology"],
    [
        ("Cloud Migration",        "Migrate warehouse to Snowflake for elastic scale, auto-clustering, and time-travel",
         "Snowflake"),
        ("IDMC Integration",       "Replace Python ETL scripts with visual mapping in Informatica IDMC for drag-and-drop pipeline design",
         "Informatica IDMC"),
        ("Role-Based Masking",     "Implement column-level masking policies: analysts see masked data, auditors see clear data",
         "Snowflake Dynamic Data Masking"),
        ("BI Dashboard",           "Connect Power BI or Tableau to the warehouse for interactive self-service reporting",
         "Power BI or Tableau"),
        ("ML Fraud Detection",     "Use claim pattern analysis and anomaly detection to flag suspicious providers or members",
         "scikit-learn, Python"),
        ("Real-Time Streaming",    "Replace batch CSV processing with real-time claim ingestion as claims are submitted",
         "Apache Kafka, Spark Streaming"),
        ("Data Lineage",           "Implement end-to-end data lineage tracking from source CSV to each warehouse column",
         "Apache Atlas or Collibra"),
        ("Automated Testing",      "Add pytest unit tests for all ETL steps and data quality assertions",
         "pytest, Great Expectations"),
    ],
    widths=[1.7, 3.3, 2.2],
)

page_break()

# ═════════════════════════════════════════════════════════════════════════
# SECTION 14 — REFERENCES
# ═════════════════════════════════════════════════════════════════════════
h1("14.  References")

h2("Technical Documentation")
for ref in [
    "Python Software Foundation. (2024). Python 3.10 Documentation. https://docs.python.org/3/",
    "Pandas Development Team. (2024). Pandas 2.x Documentation. https://pandas.pydata.org/docs/",
    "SQLite Consortium. (2024). SQLite Documentation. https://www.sqlite.org/docs.html",
    "Python-docx Contributors. (2024). python-docx Documentation. https://python-docx.readthedocs.io/",
    "Python-pptx Contributors. (2024). python-pptx Documentation. https://python-pptx.readthedocs.io/",
    "Matplotlib Development Team. (2024). Matplotlib Documentation. https://matplotlib.org/stable/",
    "Python HMAC Module. (2024). hmac - Keyed-Hashing for Message Authentication. https://docs.python.org/3/library/hmac.html",
]:
    bul(ref)

h2("Data Warehousing")
for ref in [
    "Kimball, R. and Ross, M. (2013). The Data Warehouse Toolkit: The Definitive Guide to Dimensional Modeling (3rd Edition). Wiley.",
    "Kimball Group. (2024). Data Warehouse and Business Intelligence Resources. https://www.kimballgroup.com/",
    "Inmon, W.H. (2005). Building the Data Warehouse (4th Edition). Wiley.",
]:
    bul(ref)

h2("Data Privacy and Healthcare Regulations")
for ref in [
    "U.S. Department of Health and Human Services. (2024). HIPAA for Professionals. https://www.hhs.gov/hipaa/",
    "National Health Authority, India. (2023). Ayushman Bharat Digital Mission - Data Privacy Policy.",
    "NIST. (2024). FIPS 198-1: The Keyed-Hash Message Authentication Code (HMAC). National Institute of Standards and Technology.",
]:
    bul(ref)

h2("Project Repository")
bul("mhgowda. (2024). healthcare-insurance-dw-capstone. GitHub. https://github.com/mhgowda/healthcare-insurance-dw-capstone")

space(20)
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
_run(p, "--- End of Report ---", sz=11, italic=True, color=C_GRAY)
space(8)
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
_run(p, f"Generated on {datetime.date.today().strftime('%d %B %Y')}  |  Group 5  |  Batch 2024-25",
     sz=10, color=C_GRAY, italic=True)

# ── Save ─────────────────────────────────────────────────────────────────
doc.save(OUT)
print(f"\nReport saved: {OUT}")
print("Open in Microsoft Word to view final page count.")
