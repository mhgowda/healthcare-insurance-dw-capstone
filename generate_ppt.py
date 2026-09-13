"""
generate_ppt.py  –  Healthcare Insurance DW Capstone PPT
Run:    python generate_ppt.py
Output: Healthcare_Insurance_DW_Capstone.pptx
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pathlib import Path

ROOT       = Path(__file__).resolve().parent
CHARTS_DIR = ROOT / "capstone" / "output" / "charts"
OUT_FILE   = ROOT / "Healthcare_Insurance_DW_Capstone.pptx"

# ── Palette ────────────────────────────────────────────────────────────────
NAVY    = RGBColor(0x0D, 0x1B, 0x3E)
BLUE    = RGBColor(0x1A, 0x6F, 0xC4)
TEAL    = RGBColor(0x00, 0xBF, 0xA5)
GOLD    = RGBColor(0xFF, 0xB3, 0x00)
RED     = RGBColor(0xE5, 0x39, 0x35)
GREEN   = RGBColor(0x2E, 0x7D, 0x32)
PURPLE  = RGBColor(0x6A, 0x1B, 0x9A)
ORANGE  = RGBColor(0xF5, 0x7C, 0x00)
WHITE   = RGBColor(0xFF, 0xFF, 0xFF)
OFFWH   = RGBColor(0xF5, 0xF7, 0xFB)
LTGRAY  = RGBColor(0xE8, 0xED, 0xF5)
GRAY    = RGBColor(0x55, 0x5F, 0x77)
DARK    = RGBColor(0x12, 0x1A, 0x2E)
LTBLUE  = RGBColor(0xBB, 0xD6, 0xF5)
LTGREEN = RGBColor(0xC8, 0xE6, 0xC9)
LTRED   = RGBColor(0xFF, 0xCC, 0xBC)
LTYEL   = RGBColor(0xFF, 0xF9, 0xC4)

prs = Presentation()
prs.slide_width  = Inches(13.33)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]


# ═══════════════════════════════════════════════════════════════════════════
# PRIMITIVE HELPERS
# ═══════════════════════════════════════════════════════════════════════════

def R(s, l, t, w, h, fill):
    sh = s.shapes.add_shape(1, Inches(l), Inches(t), Inches(w), Inches(h))
    sh.fill.solid(); sh.fill.fore_color.rgb = fill
    sh.line.fill.background()
    return sh

def T(s, text, l, t, w, h, sz=14, bold=False, color=DARK,
      align=PP_ALIGN.LEFT, italic=False):
    tb = s.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tb.word_wrap = True
    tf = tb.text_frame; tf.word_wrap = True
    p  = tf.paragraphs[0]; p.alignment = align
    run = p.add_run(); run.text = text
    run.font.size = Pt(sz); run.font.bold = bold
    run.font.color.rgb = color; run.font.italic = italic
    return tb

def IMG(s, path, l, t, w, h=None):
    p = Path(path)
    if not p.exists(): return
    if h: s.shapes.add_picture(str(p), Inches(l), Inches(t), Inches(w), Inches(h))
    else: s.shapes.add_picture(str(p), Inches(l), Inches(t), Inches(w))

def hline(s, x, y, w, col=GRAY, thick=0.025):
    R(s, x, y, w, thick, col)

def vline(s, x, y, h, col=GRAY, thick=0.025):
    R(s, x, y, thick, h, col)

# ── Background helpers ──────────────────────────────────────────────────────

def dark_bg(s):
    """Full dark navy background with teal top stripe and gold bottom stripe."""
    R(s, 0, 0, 13.33, 7.5, NAVY)
    R(s, 0, 0, 13.33, 0.10, TEAL)
    R(s, 0, 7.38, 13.33, 0.12, GOLD)
    # decorative right panel
    R(s, 11.0, 0, 2.33, 7.5, RGBColor(0x10, 0x23, 0x4E))

def light_bg(s):
    """Light slide background with coloured top bar."""
    R(s, 0, 0, 13.33, 7.5, OFFWH)
    R(s, 0, 0, 13.33, 0.55, NAVY)
    R(s, 0, 7.3, 13.33, 0.20, NAVY)

def slide_title(s, title, subtitle=""):
    """Standard slide title block."""
    T(s, title, 0.35, 0.08, 12.6, 0.42,
      sz=24, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
    if subtitle:
        T(s, subtitle, 0.35, 0.47, 10, 0.28,
          sz=12, color=LTBLUE, align=PP_ALIGN.LEFT)
    # teal underline
    R(s, 0.35, 0.78, 3.5, 0.055, TEAL)

def section_divider(s, number, title, sub=""):
    """Full-bleed section divider slide."""
    dark_bg(s)
    # big number watermark
    T(s, number, 0.3, 0.8, 3, 3.5, sz=160, bold=True,
      color=RGBColor(0x1E, 0x33, 0x60), align=PP_ALIGN.LEFT)
    R(s, 0.35, 3.2, 7.0, 0.07, TEAL)
    T(s, title, 0.35, 3.35, 10.3, 1.2,
      sz=42, bold=True, color=WHITE)
    if sub:
        T(s, sub, 0.35, 4.65, 10.3, 0.8,
          sz=18, color=LTBLUE)

# ── Table helper ────────────────────────────────────────────────────────────

def trow(s, cells, widths, x, y, rh, bg, tc, sz=11, bold=False):
    cx = x
    for cell, cw in zip(cells, widths):
        R(s, cx, y, cw-0.03, rh, bg)
        T(s, str(cell), cx+0.08, y+0.04, cw-0.16, rh-0.05,
          sz=sz, bold=bold, color=tc)
        cx += cw

# ── KPI card ────────────────────────────────────────────────────────────────

def kpi(s, x, y, w, h, label, value, sub="", col=BLUE):
    R(s, x, y, w, h, col)
    R(s, x, y+h-0.07, w, 0.07, TEAL)          # bottom accent
    T(s, label, x+0.15, y+0.12, w-0.2, 0.32,
      sz=11, color=RGBColor(0xDD, 0xEE, 0xFF))
    T(s, value, x+0.12, y+0.44, w-0.2, 0.75,
      sz=30, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
    if sub:
        T(s, sub, x+0.15, y+1.15, w-0.2, 0.3,
          sz=10, color=RGBColor(0xCC, 0xE5, 0xFF))

# ── Info box ────────────────────────────────────────────────────────────────

def info_box(s, x, y, w, h, title, body, col=BLUE, text_col=DARK):
    R(s, x, y, w, h, LTGRAY)
    R(s, x, y, 0.08, h, col)
    T(s, title, x+0.18, y+0.08, w-0.25, 0.35,
      sz=13, bold=True, color=col)
    T(s, body,  x+0.18, y+0.48, w-0.25, h-0.55,
      sz=11, color=text_col)


# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 1 — TITLE
# ═══════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
dark_bg(s)

# Left content
T(s, "MIDTERM CAPSTONE PROJECT", 0.5, 1.0, 10, 0.45,
  sz=12, bold=True, color=TEAL)
T(s, "Healthcare Insurance", 0.5, 1.52, 10.3, 1.1,
  sz=52, bold=True, color=WHITE)
T(s, "Claims Data Warehouse", 0.5, 2.55, 10.3, 1.0,
  sz=46, bold=False, color=LTBLUE)
T(s, "& Analytics Platform", 0.5, 3.48, 10.3, 0.8,
  sz=36, bold=False, color=LTBLUE)

R(s, 0.5, 4.55, 7.5, 0.055, TEAL)
T(s, "SQL  |  Python  |  ETL Pipeline  |  Star Schema  |  Data Masking  |  Visualization",
  0.5, 4.7, 10.3, 0.45, sz=14, color=GRAY)

T(s, "Group 5  |  Batch 2024-25", 0.5, 5.4, 5, 0.4, sz=13, color=WHITE)
T(s, "Midterm Capstone", 0.5, 5.85, 5, 0.35, sz=12, color=GRAY)

# Right decorative panel stats
for ry, val, lbl in [(1.3, "1.4M", "Source Rows"),
                     (2.8, "9",    "DW Tables"),
                     (4.3, "8",    "Charts")]:
    T(s, val, 11.2, ry,    2.0, 0.8, sz=32, bold=True, color=GOLD,   align=PP_ALIGN.CENTER)
    T(s, lbl, 11.2, ry+0.7, 2.0, 0.35, sz=11, color=LTBLUE, align=PP_ALIGN.CENTER)
    R(s, 11.4, ry+1.1, 1.5, 0.03, TEAL)


# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 2 — PROJECT OVERVIEW
# ═══════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
light_bg(s)
slide_title(s, "Project Overview", "End-to-end healthcare insurance claims data warehouse solution")

# Business scenario box
R(s, 0.3, 0.95, 12.73, 1.55, RGBColor(0xE8, 0xF0, 0xFE))
R(s, 0.3, 0.95, 0.1,   1.55, BLUE)
T(s, "Business Scenario", 0.55, 1.0, 6, 0.38, sz=13, bold=True, color=BLUE)
T(s, "A national health insurer needs a centralized Data Warehouse to analyze claims costs, "
     "utilisation, provider performance and fraud indicators — while fully protecting patient "
     "privacy (PII/PHI). Source data arrives as 7 CSV files with ~1.4 million rows containing "
     "intentional quality issues including missing values, orphan foreign keys, and type mismatches.",
  0.5, 1.38, 12.3, 1.0, sz=12, color=DARK)

# 4 objective boxes
objs = [
    ("Star Schema\nDesign",   "Model a DW with\ndims + facts",          BLUE),
    ("ETL Pipeline",          "Full + Incremental\nload with audit logs",TEAL),
    ("Data Masking",          "Protect PII/PHI\npreserve join keys",     PURPLE),
    ("Analytics\n& Charts",   "SQL + Python +\n8 Matplotlib charts",    ORANGE),
]
for i, (t, b, c) in enumerate(objs):
    x = 0.3 + i * 3.25
    R(s, x, 2.7, 3.1, 2.1, c)
    R(s, x, 2.7, 3.1, 0.08, RGBColor(0xFF,0xFF,0xFF))
    T(s, str(i+1), x+0.12, 2.73, 0.5, 0.55, sz=20, bold=True, color=WHITE)
    T(s, t,   x+0.6, 2.73, 2.4, 0.7,  sz=13, bold=True, color=WHITE)
    T(s, b,   x+0.12, 3.4,  2.9, 0.9,  sz=11, color=RGBColor(0xEE,0xF6,0xFF))

# Tech stack badges
T(s, "Technologies:", 0.3, 5.05, 2.2, 0.38, sz=12, bold=True, color=NAVY)
techs = ["Python 3.10+", "Pandas", "SQLite / SQL", "Matplotlib", "Jupyter Notebook"]
cols  = [NAVY, BLUE, TEAL, ORANGE, PURPLE]
for i, (tech, tc) in enumerate(zip(techs, cols)):
    R(s, 0.3 + i*2.58, 5.48, 2.45, 0.55, tc)
    T(s, tech, 0.3+i*2.58+0.12, 5.55, 2.2, 0.38,
      sz=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

T(s, "Dataset: 7 CSV files  |  ~1.4 Million rows  |  1.2M clean + 0.2M intentionally imperfect",
  0.3, 6.2, 12.7, 0.4, sz=11, italic=True, color=GRAY, align=PP_ALIGN.CENTER)


# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 3 — SOURCE DATA
# ═══════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
light_bg(s)
slide_title(s, "Source Data  –  7 CSV Files", "OLTP-style extracts from the insurance system")

hdrs = ["CSV File", "Description", "Rows", "Primary Key", "Key Columns"]
ws   = [2.9, 3.2, 1.1, 1.8, 3.5]
trow(s, hdrs, ws, 0.2, 0.9, 0.42, NAVY, WHITE, sz=11, bold=True)

rows = [
    ("members_merged.csv",     "Patient/member master",          "2,60,000", "member_id",   "name, DOB, email, phone, city, state"),
    ("providers_merged.csv",   "Doctors and hospitals",          "40,000",   "provider_id", "specialty, city, state, tax_id, license_no"),
    ("claims_merged.csv",      "Insurance claim headers",        "2,40,000", "claim_number","member_id, provider_id, admit_date, paid_amount"),
    ("claim_lines_merged.csv", "Services inside each claim",     "7,60,000", "line_id",     "claim_number, procedure_code, units, rate"),
    ("payments_merged.csv",    "Payments on claims",             "1,00,000", "payment_id",  "claim_number, payment_date, paid_amount"),
    ("diagnoses_merged.csv",   "ICD diagnosis codes per claim",  "60,000",   "diag_id",     "claim_number, diagnosis_code, icd_version"),
    ("procedures_merged.csv",  "CPT codes per claim line",       "40,000",   "proc_id",     "line_id, procedure_code"),
]
bgs = [WHITE, LTGRAY]
for i, row in enumerate(rows):
    trow(s, row, ws, 0.2, 1.35+i*0.54, 0.5, bgs[i%2], DARK, 10)

trow(s, ["TOTAL", "7 source entities", "~14,00,000", "—", "1.2M clean + 0.2M imperfect"],
     ws, 0.2, 5.15, 0.48, NAVY, WHITE, sz=11, bold=True)

# Issue badges at bottom
issues = [("Missing Values","blank/NULL fields"), ("Orphan FKs","__ORPHAN__ marker"),
          ("Negative Amounts","invalid numbers"), ("Future Dates","impossible timestamps"),
          ("Duplicate Keys","repeated IDs"),      ("Type Mismatches","text in numbers")]
for i, (t, b) in enumerate(issues):
    x = 0.2 + (i%3)*4.35
    y = 5.82 + (i//3)*0.6
    R(s, x, y, 4.2, 0.52, [RED,ORANGE,PURPLE,BLUE,TEAL,GREEN][i])
    T(s, t+": "+b, x+0.12, y+0.08, 4.0, 0.35, sz=10, bold=True, color=WHITE)


# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 4 — ER DIAGRAM
# ═══════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
light_bg(s)
slide_title(s, "Entity Relationship Diagram",
            "Source system — table structures and foreign key relationships")

# ── ER table drawing function ───────────────────────────────────────────────
# Each table: header bar + alternating field rows
# field tuples: (field_name, data_type, flag)  flag = "PK" | "FK" | ""

ROW_H   = 0.265   # height of each field row
HDR_H   = 0.36    # height of header bar
PK_COL  = RGBColor(0xFF, 0xA0, 0x00)   # amber for PK badge
FK_COL  = RGBColor(0x6A, 0x1B, 0x9A)   # purple for FK badge
F1_COL  = WHITE
F2_COL  = RGBColor(0xF4, 0xF6, 0xFB)   # very light stripe

def er_table(sl, x, y, w, name, fields, hdr_color):
    """Draw one ER entity box — header + field rows, perfectly aligned."""
    total_h = HDR_H + len(fields) * ROW_H + 0.05
    # outer border
    border = sl.shapes.add_shape(1, Inches(x-0.02), Inches(y-0.02),
                                  Inches(w+0.04), Inches(total_h+0.04))
    border.fill.solid(); border.fill.fore_color.rgb = RGBColor(0xCC,0xD6,0xE8)
    border.line.fill.background()

    # header
    R(sl, x, y, w, HDR_H, hdr_color)
    T(sl, name, x+0.12, y+0.06, w-0.18, HDR_H-0.08,
      sz=11, bold=True, color=WHITE)

    # field rows
    for i, (fname, ftype, flag) in enumerate(fields):
        fy  = y + HDR_H + i * ROW_H
        fbg = F1_COL if i % 2 == 0 else F2_COL
        R(sl, x, fy, w, ROW_H, fbg)
        # thin separator
        R(sl, x, fy, w, 0.012, RGBColor(0xDD,0xE4,0xEE))

        if flag:
            badge_col = PK_COL if flag == "PK" else FK_COL
            R(sl, x+0.06, fy+0.05, 0.26, 0.165, badge_col)
            T(sl, flag, x+0.06, fy+0.045, 0.26, 0.175,
              sz=6, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
            T(sl, fname, x+0.38, fy+0.05, w*0.52, 0.2,
              sz=9, color=DARK, bold=(flag == "PK"))
        else:
            T(sl, fname, x+0.12, fy+0.05, w*0.58, 0.2, sz=9, color=DARK)

        T(sl, ftype, x+w-1.05, fy+0.06, 0.98, 0.185,
          sz=8, color=GRAY, italic=True, align=PP_ALIGN.RIGHT)

    return total_h

# ── Layout: 3-column grid ────────────────────────────────────────────────
# Column X positions
C1 = 0.18    # left
C2 = 4.55    # centre
C3 = 8.92    # right
TW = 4.18    # table width

# Row Y positions
R1 = 0.92    # top row
R2 = 4.05    # bottom row  (adjusted after measuring actual table heights)

# Top row tables
h_mem  = er_table(s, C1, R1, TW, "MEMBERS",
    [("member_id",    "VARCHAR", "PK"),
     ("first_name",   "VARCHAR", ""),
     ("last_name",    "VARCHAR", ""),
     ("email",        "VARCHAR", ""),
     ("phone",        "VARCHAR", ""),
     ("dob",          "DATE",    ""),
     ("city",         "VARCHAR", ""),
     ("state",        "VARCHAR", ""),
     ("is_active",    "BOOLEAN", "")], BLUE)

h_cla  = er_table(s, C2, R1, TW, "CLAIMS",
    [("claim_number",  "VARCHAR", "PK"),
     ("member_id",     "VARCHAR", "FK"),
     ("provider_id",   "VARCHAR", "FK"),
     ("admit_date",    "DATE",    ""),
     ("discharge_date","DATE",    ""),
     ("length_of_stay","INT",     ""),
     ("allowed_amount","NUMERIC", ""),
     ("paid_amount",   "NUMERIC", ""),
     ("denial_code",   "VARCHAR", "")], NAVY)

h_pay  = er_table(s, C3, R1, TW, "PAYMENTS",
    [("payment_id",     "VARCHAR", "PK"),
     ("claim_number",   "VARCHAR", "FK"),
     ("payment_date",   "DATE",    ""),
     ("payment_method", "VARCHAR", ""),
     ("paid_amount",    "NUMERIC", ""),
     ("adjustment_code","VARCHAR", "")], ORANGE)

# Bottom row tables — positioned exactly below top row + gap
GAP = 0.22
R2_actual = R1 + max(h_mem, h_cla, h_pay) + GAP

h_prv  = er_table(s, C1, R2_actual, TW, "PROVIDERS",
    [("provider_id",   "VARCHAR", "PK"),
     ("provider_name", "VARCHAR", ""),
     ("specialty",     "VARCHAR", ""),
     ("city",          "VARCHAR", ""),
     ("state",         "VARCHAR", ""),
     ("tax_id",        "VARCHAR", ""),
     ("license_no",    "VARCHAR", "")], TEAL)

h_cl   = er_table(s, C2, R2_actual, TW, "CLAIM_LINES",
    [("line_id",        "VARCHAR", "PK"),
     ("claim_number",   "VARCHAR", "FK"),
     ("service_date",   "DATE",    ""),
     ("procedure_code", "VARCHAR", ""),
     ("units",          "INT",     ""),
     ("rate",           "NUMERIC", ""),
     ("line_amount",    "NUMERIC", "")], NAVY)

h_diag = er_table(s, C3, R2_actual, TW, "DIAGNOSES",
    [("diag_id",        "VARCHAR", "PK"),
     ("claim_number",   "VARCHAR", "FK"),
     ("diagnosis_code", "VARCHAR", ""),
     ("icd_version",    "VARCHAR", "")], RED)

# ── Relationship connector lines ─────────────────────────────────────────
# All connectors use the exact midpoints of table edges

def mid_right(tx, ty, tw, th):
    """Right-edge midpoint of a table."""
    return tx + tw, ty + th / 2

def mid_left(tx, ty, th):
    """Left-edge midpoint of a table."""
    return tx, ty + th / 2

def mid_bottom(tx, ty, tw, th):
    """Bottom-edge midpoint of a table."""
    return tx + tw / 2, ty + th

def mid_top(tx, ty, tw):
    """Top-edge midpoint of a table."""
    return tx + tw / 2, ty

LTHICK = 0.03

# 1. MEMBERS right -> CLAIMS left  (member_id FK)
mx, my = mid_right(C1, R1, TW, h_mem)
cx, cy = mid_left(C2, R1, h_cla)
# horizontal bridge
hline(s, mx, my - LTHICK/2, cx - mx, BLUE, LTHICK)

# 2. PROVIDERS right -> CLAIMS left  (provider_id FK)
# providers is on row2, claims on row1 — need L-shaped connector
px, py  = mid_right(C1, R2_actual, TW, h_prv)
cx2, cy2 = mid_left(C2, R1, h_cla)
mid_x = C2 - 0.18
hline(s, px, py, mid_x - px, TEAL, LTHICK)
vline(s, mid_x, cy2, py - cy2, TEAL, LTHICK)
hline(s, mid_x, cy2, C2 - mid_x, TEAL, LTHICK)

# 3. CLAIMS right -> PAYMENTS left  (claim_number FK)
clx, cly = mid_right(C2, R1, TW, h_cla)
payx, payy = mid_left(C3, R1, h_pay)
hline(s, clx, cly - LTHICK/2, payx - clx, NAVY, LTHICK)

# 4. CLAIMS bottom -> CLAIM_LINES top  (claim_number FK)
clbx, clby = mid_bottom(C2, R1, TW, h_cla)
cltx, clty = mid_top(C2, R2_actual, TW)
# vertical connector between same column
vline(s, clbx, clby, clty - clby, NAVY, LTHICK)

# 5. CLAIMS right -> DIAGNOSES left  (claim_number FK)
# L-shaped: right of claims (row1) -> right, down, to right of diagnoses (row2)
clrx, clry = mid_right(C2, R1, TW, h_cla)
dglx, dgly = mid_left(C3, R2_actual, h_diag)
bridge_x = C3 - 0.15
hline(s, clrx, clry, bridge_x - clrx, RED, LTHICK)
vline(s, bridge_x, clry, dgly - clry, RED, LTHICK)
hline(s, bridge_x, dgly, C3 - bridge_x, RED, LTHICK)

# 6. CLAIM_LINES right -> PAYMENTS (no direct relation — skip)
# 7. PROVIDERS bottom -> CLAIM_LINES left  (via claim -> claim_lines, no direct)

# ── Legend box ──────────────────────────────────────────────────────────────
# Place legend in bottom-right corner if space allows
leg_x = C3 + TW + 0.15
leg_y = R2_actual
leg_w = 13.13 - leg_x
if leg_w < 0.5:
    leg_x = 0.18
    leg_y = R2_actual + h_diag + 0.15
    leg_w = TW

R(s, leg_x, leg_y, leg_w, 1.05, RGBColor(0xEE,0xF2,0xFA))
R(s, leg_x, leg_y, leg_w, 0.30, NAVY)
T(s, "LEGEND", leg_x+0.1, leg_y+0.05, leg_w-0.15, 0.22,
  sz=10, bold=True, color=WHITE)
for li, (lc, lt) in enumerate([
    (PK_COL,  "PK  Primary Key"),
    (FK_COL,  "FK  Foreign Key"),
    (NAVY,    "Line  Relationship"),
]):
    ly2 = leg_y + 0.36 + li * 0.24
    R(s, leg_x+0.1, ly2, 0.22, 0.165, lc)
    T(s, lt, leg_x+0.38, ly2, leg_w-0.45, 0.185, sz=9, color=DARK)


# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 5 — STAR SCHEMA
# ═══════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
light_bg(s)
slide_title(s, "Star Schema  –  Data Warehouse Design",
            "Dimensional model: fact tables (numbers) surrounded by dimension tables (descriptions)")

# ── DW box drawing function ──────────────────────────────────────────────
DW_HDR  = 0.35
DW_ROW  = 0.235
DW_F1   = WHITE
DW_F2   = RGBColor(0xF0,0xF4,0xFB)

def dw_table(sl, x, y, w, name, badge, badge_col, hdr_col, fields):
    """Draw one star-schema box with badge, header, and field rows."""
    total_h = DW_HDR + len(fields) * DW_ROW + 0.04
    # shadow / border
    bdr = sl.shapes.add_shape(1, Inches(x+0.04), Inches(y+0.04),
                               Inches(w), Inches(total_h))
    bdr.fill.solid(); bdr.fill.fore_color.rgb = RGBColor(0xC8,0xD2,0xE4)
    bdr.line.fill.background()
    # main box
    R(sl, x, y, w, total_h, RGBColor(0xFA,0xFC,0xFF))
    # header
    R(sl, x, y, w, DW_HDR, hdr_col)
    # badge pill
    R(sl, x+0.08, y+0.07, 0.52, 0.21, badge_col)
    T(sl, badge, x+0.08, y+0.065, 0.52, 0.22,
      sz=7, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    T(sl, name, x+0.66, y+0.055, w-0.72, 0.26,
      sz=10, bold=True, color=WHITE)
    # fields
    for i, fld in enumerate(fields):
        fy  = y + DW_HDR + i * DW_ROW
        fbg = DW_F1 if i % 2 == 0 else DW_F2
        R(sl, x, fy, w, DW_ROW, fbg)
        R(sl, x, fy, w, 0.01, RGBColor(0xE0,0xE8,0xF4))
        T(sl, fld, x+0.12, fy+0.03, w-0.18, DW_ROW-0.04, sz=8, color=DARK)
    return total_h

# ── Grid layout — fixed coordinates ────────────────────────────────────
# Slide usable area: x 0.18–13.15,  y 0.88–7.28
# Layout:
#   Row 1 (y=0.88):  dim_member  | [gap] dim_date [gap] |  dim_provider
#   Row 2 (y=3.05):  fact_claim_line | fact_claim | fact_payment
#   Row 3 (y=5.45):  dim_diagnosis | dim_location | dim_procedure

DW = 3.62    # box width
G  = 0.245   # gap between boxes

# Column left edges (3 columns, evenly spaced)
X1 = 0.18
X2 = X1 + DW + G
X3 = X2 + DW + G

Y1 = 0.88
Y2 = 3.20
Y3 = 5.50

# Row 1
h_dm  = dw_table(s, X1, Y1, DW, "dim_member",   "DIM", TEAL, BLUE,
    ["member_id  (PK) — masked",
     "gender",
     "city,  state",
     "birth_year",
     "is_active"])

h_dd  = dw_table(s, X2, Y1, DW, "dim_date",     "DIM", TEAL, BLUE,
    ["date_key  (PK)",
     "year,  quarter,  month",
     "month_name",
     "day_name,  is_weekend"])

h_dp  = dw_table(s, X3, Y1, DW, "dim_provider", "DIM", TEAL, BLUE,
    ["provider_id  (PK) — masked",
     "specialty",
     "city,  state",
     "tax_id — masked",
     "license_no — masked"])

# Row 2 — fact tables
h_fcl = dw_table(s, X1, Y2, DW, "fact_claim_line", "FACT", GOLD, NAVY,
    ["line_id  (PK)",
     "claim_number  (FK)",
     "service_date_key  (FK)",
     "procedure_code  (FK)",
     "units,  rate,  line_amount"])

h_fc  = dw_table(s, X2, Y2, DW, "fact_claim",      "FACT", GOLD, NAVY,
    ["claim_number  (PK)",
     "member_id  (FK)",
     "provider_id  (FK)",
     "admit_date_key  (FK)",
     "length_of_stay",
     "allowed_amount,  paid_amount",
     "denial_code,  is_denied"])

h_fp  = dw_table(s, X3, Y2, DW, "fact_payment",    "FACT", GOLD, NAVY,
    ["payment_id  (PK)",
     "claim_number  (FK)",
     "payment_date_key  (FK)",
     "payment_method",
     "paid_amount"])

# Row 3 — remaining dims
h_ddiag = dw_table(s, X1, Y3, DW, "dim_diagnosis", "DIM", TEAL, BLUE,
    ["diagnosis_code  (PK)",
     "icd_version"])

h_dloc  = dw_table(s, X2, Y3, DW, "dim_location",  "DIM", TEAL, BLUE,
    ["location_key  (PK)",
     "city,  state"])

h_dproc = dw_table(s, X3, Y3, DW, "dim_procedure", "DIM", TEAL, BLUE,
    ["procedure_code  (PK)"])

# ── Connectors — precise midpoints ──────────────────────────────────────
# Helper: midpoint of each edge
def top_mid(x, y, w):    return x + w/2, y
def bot_mid(x, y, w, h): return x + w/2, y + h
def lft_mid(x, y, h):    return x,       y + h/2
def rgt_mid(x, y, w, h): return x + w,   y + h/2

C_THICK = 0.028
C_DIM   = TEAL
C_FACT  = NAVY

# dim_member  (row1) -> fact_claim (row2): vertical from bottom of dm to top of fc
dmb_x, dmb_y = bot_mid(X1, Y1, DW, h_dm)
fct_x, fct_y = top_mid(X2, Y2, DW)
# L-shape: down from dm, right to fc centre-x
vline(s, dmb_x, dmb_y, Y2 - 0.18 - dmb_y, C_DIM, C_THICK)
hline(s, dmb_x, Y2 - 0.18, fct_x - dmb_x, C_DIM, C_THICK)
vline(s, fct_x, Y2 - 0.18, 0.18, C_DIM, C_THICK)

# dim_date (row1) -> fact_claim (row2): straight vertical (same column X2)
ddb_x, ddb_y = bot_mid(X2, Y1, DW, h_dd)
vline(s, ddb_x, ddb_y, fct_y - ddb_y, C_DIM, C_THICK)

# dim_provider (row1) -> fact_claim (row2): mirror of dim_member
dpb_x, dpb_y = bot_mid(X3, Y1, DW, h_dp)
vline(s, dpb_x, dpb_y, Y2 - 0.18 - dpb_y, C_DIM, C_THICK)
hline(s, fct_x, Y2 - 0.18, dpb_x - fct_x, C_DIM, C_THICK)
vline(s, dpb_x, Y2 - 0.18, 0.18, C_DIM, C_THICK)

# fact_claim (row2) -> fact_claim_line (row2): horizontal between same row
fcrx, fcry = lft_mid(X2, Y2, h_fc)
fclrx, fclry = rgt_mid(X1, Y2, DW, h_fcl)
hline(s, fclrx, fclry, fcrx - fclrx, C_FACT, C_THICK)

# fact_claim (row2) -> fact_payment (row2): horizontal
fcrx2, fcry2 = rgt_mid(X2, Y2, DW, h_fc)
fplx,  fply  = lft_mid(X3, Y2, h_fp)
hline(s, fcrx2, fcry2, fplx - fcrx2, C_FACT, C_THICK)

# fact_claim_line (row2) -> dim_diagnosis (row3): vertical
fclt_x, fclt_y = bot_mid(X1, Y2, DW, h_fcl)
ddt_x,  ddt_y  = top_mid(X1, Y3, DW)
vline(s, fclt_x, fclt_y, ddt_y - fclt_y, C_DIM, C_THICK)

# fact_claim (row2) -> dim_location (row3): vertical same column
fcbt_x, fcbt_y = bot_mid(X2, Y2, DW, h_fc)
dlct_x, dlct_y = top_mid(X2, Y3, DW)
vline(s, fcbt_x, fcbt_y, dlct_y - fcbt_y, C_DIM, C_THICK)

# fact_claim_line (row2) -> dim_procedure (row3): L-shape
fclb_x, fclb_y = bot_mid(X1, Y2, DW, h_fcl)
dpt_x,  dpt_y  = top_mid(X3, Y3, DW)
mid_y3 = Y3 - 0.15
hline(s, fclb_x, mid_y3, dpt_x - fclb_x, C_DIM, C_THICK)
vline(s, dpt_x,  mid_y3, dpt_y - mid_y3,  C_DIM, C_THICK)

# fact_payment (row2) -> dim_procedure (row3): vertical same column
fpbt_x, fpbt_y = bot_mid(X3, Y2, DW, h_fp)
vline(s, fpbt_x, fpbt_y, dpt_y - fpbt_y, C_DIM, C_THICK)

# ── Legend ───────────────────────────────────────────────────────────────
# Small legend at bottom right — doesn't exist as separate box, inline text
R(s, 0.18, 7.08, 12.97, 0.22, RGBColor(0xE8,0xED,0xF8))
legend_items = [
    (GOLD,  "FACT = Fact Table (stores measurable data)"),
    (TEAL,  "DIM = Dimension Table (stores descriptive data)"),
    (NAVY,  "Lines = Relationships between tables"),
]
for li, (lc, lt) in enumerate(legend_items):
    lx = 0.28 + li * 4.35
    R(s, lx, 7.1, 0.22, 0.165, lc)
    T(s, lt, lx+0.28, 7.1, 4.0, 0.165, sz=9, color=DARK)


# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 6 — ETL PIPELINE
# ═══════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
light_bg(s)
slide_title(s, "ETL Pipeline Architecture",
            "Extract  ->  Profile  ->  Clean  ->  Mask  ->  Load  ->  Analyse")

# Flow boxes
steps = [
    ("01", "Source\nCSVs",         "7 files\n1.4M rows",           BLUE,   ""),
    ("02", "Profile",              "Detect all\nquality issues",    TEAL,   "step1_profile.py"),
    ("03", "Clean &\nValidate",    "Remove bad rows\nSave rejects", ORANGE, "step2_clean.py"),
    ("04", "Mask\nPII / PHI",      "Hash sensitive\nfields",       PURPLE, "step3_mask.py"),
    ("05", "Full ETL\nLoad",       "Build schema\nLoad all data",   NAVY,   "step4_load_full.py"),
    ("06", "Incremental\nLoad",    "New rows only\nwatermark-based",GREEN,  "step5_incremental.py"),
]
for i, (num, title, body, col, fname) in enumerate(steps):
    x = 0.28 + i * 2.17
    R(s, x, 1.05, 2.02, 2.5, col)
    T(s, num,   x+0.1, 1.1,  0.55, 0.5, sz=22, bold=True, color=WHITE)
    T(s, title, x+0.1, 1.62, 1.85, 0.75, sz=13, bold=True, color=WHITE)
    T(s, body,  x+0.1, 2.38, 1.85, 0.95, sz=11,
      color=RGBColor(0xDD,0xEE,0xFF))
    if fname:
        T(s, fname, x, 3.62, 2.05, 0.32,
          sz=8, italic=True, color=GRAY, align=PP_ALIGN.CENTER)
    if i < 5:
        T(s, ">", x+1.95, 2.0, 0.3, 0.5,
          sz=22, bold=True, color=GRAY, align=PP_ALIGN.CENTER)

# Two mode comparison
R(s, 0.25, 4.05, 6.15, 2.85, RGBColor(0xFF,0xEB,0xEE))
R(s, 0.25, 4.05, 0.1, 2.85, RED)
T(s, "FULL LOAD",   0.5, 4.12, 5.5, 0.4, sz=14, bold=True, color=RED)
T(s, "-> Drop all existing warehouse data\n"
     "-> Run DDL to recreate all tables\n"
     "-> Clean + mask + load all 1.4M rows\n"
     "-> Use when: first run or full refresh needed",
  0.5, 4.55, 5.7, 1.95, sz=12, color=DARK)

R(s, 6.65, 4.05, 6.43, 2.85, RGBColor(0xE8,0xF5,0xE9))
R(s, 6.65, 4.05, 0.1,  2.85, GREEN)
T(s, "INCREMENTAL LOAD", 6.9, 4.12, 6.0, 0.4, sz=14, bold=True, color=GREEN)
T(s, "-> Read last_load_dt from etl_watermark table\n"
     "-> Filter CSV: rows where date > last_load_dt\n"
     "-> Process and load ONLY those new rows\n"
     "-> Use when: daily/regular refresh needed",
  6.9, 4.55, 6.0, 1.95, sz=12, color=DARK)

T(s, "Outputs: warehouse.db  |  rejected_rows/*.csv  |  "
     "logs/profiling_summary.json  |  logs/masking_audit.json",
  0.3, 7.0, 12.7, 0.35, sz=10, italic=True, color=GRAY, align=PP_ALIGN.CENTER)


# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 7 — DATA QUALITY CHALLENGES
# ═══════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
light_bg(s)
slide_title(s, "Data Quality Challenges", "Issues identified during profiling phase")

issues = [
    ("Missing Values",        "Blank or NULL fields in required columns",
     "first_name, email, city, diagnosis_code, specialty",   BLUE),
    ("Orphan Foreign Keys",   "__ORPHAN__ marker where a valid parent ID is expected",
     "member_id in claims, claim_number in payments",         RED),
    ("Negative Amounts",      "Invalid negative values in financial columns",
     "allowed_amount, paid_amount, units, line_amount",       ORANGE),
    ("Future Dates",          "Dates that should not exist (past events in the future)",
     "admit_date, service_date, DOB, payment_date",           PURPLE),
    ("Duplicate Primary Keys","Same ID appearing more than once in the same file",
     "claim_number, member_id, provider_id, line_id",        TEAL),
    ("Type Mismatches",       "Text values found in columns that must be numeric",
     "paid_amount = 'ERR', rate = 'MISSING', LOS = 'TEXT'",  GREEN),
]
for i, (title, desc, cols, col) in enumerate(issues):
    cx = 0.25 + (i % 3) * 4.35
    cy = 1.05 + (i // 3) * 2.8
    R(s, cx, cy, 4.2, 2.55, col)
    R(s, cx, cy, 4.2, 0.5, NAVY)
    T(s, title, cx+0.15, cy+0.06, 3.9, 0.38, sz=13, bold=True, color=WHITE)
    T(s, desc,  cx+0.15, cy+0.6,  3.9, 0.7,  sz=11, color=WHITE)
    T(s, "Columns: "+cols, cx+0.15, cy+1.35, 3.9, 0.9,
      sz=9, italic=True, color=RGBColor(0xDD,0xEE,0xFF))


# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 8 — DATA CLEANING RESULTS
# ═══════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
light_bg(s)
slide_title(s, "Data Cleaning Results", "Validation applied entity-by-entity in dependency order")

hdrs = ["Entity", "Total In", "Valid (Kept)", "Rejected", "Rate", "Main Rejection Reasons"]
ws   = [2.0, 1.3, 1.5, 1.3, 0.9, 6.3]
trow(s, hdrs, ws, 0.2, 0.92, 0.4, NAVY, WHITE, sz=11, bold=True)

rows = [
    ("members",     "2,60,000", "2,32,207", "27,793",   "10.7%","Missing name, future DOB, duplicate ID"),
    ("providers",   "40,000",   "38,424",   "1,576",    "3.9%", "Missing specialty, duplicate ID"),
    ("claims",      "2,40,000", "2,08,640", "31,360",   "13.1%","Orphan member_id, negative amount, future date"),
    ("claim_lines", "7,60,000", "6,45,733", "1,14,267", "15.0%","Orphan claim_number, negative units, bad rate"),
    ("payments",    "1,00,000", "59,615",   "40,385",   "40.4%","Orphan claim_number, non-numeric amount"),
    ("diagnoses",   "60,000",   "39,741",   "20,259",   "33.8%","Orphan claim_number, invalid ICD version"),
    ("procedures",  "40,000",   "24,844",   "15,156",   "37.9%","Orphan line_id, missing procedure code"),
]
for i, row in enumerate(rows):
    bg = WHITE if i%2==0 else LTGRAY
    trow(s, row, ws, 0.2, 1.35+i*0.5, 0.46, bg, DARK, sz=10)

trow(s, ["TOTAL","14,00,000","12,49,204","2,50,796","17.9%","All quality issues combined"],
     ws, 0.2, 4.88, 0.46, NAVY, WHITE, sz=11, bold=True)

# KPI strip
for i, (val, lbl, col) in enumerate([
    ("12,49,204", "Valid Rows Loaded",    GREEN),
    ("2,50,796",  "Rows Rejected",        RED),
    ("17.9%",     "Overall Reject Rate",  ORANGE),
    ("7 files",   "Audit CSVs Saved",     BLUE),
]):
    x = 0.2 + i * 3.27
    R(s, x, 5.5, 3.1, 1.65, col)
    T(s, val, x+0.15, 5.6,  2.8, 0.72, sz=28, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    T(s, lbl, x+0.1,  6.32, 2.9, 0.35, sz=11, color=WHITE,            align=PP_ALIGN.CENTER)


# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 9 — DATA MASKING
# ═══════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
light_bg(s)
slide_title(s, "Data Masking  –  PII / PHI Protection",
            "Deterministic HMAC-SHA256 tokenisation applied before warehouse loading")

# Two explanation boxes
info_box(s, 0.25, 0.92, 6.1, 1.35,
    "Why Masking?",
    "Analysts querying the warehouse must never see real patient names, emails, "
    "phone numbers or exact dates of birth. Masking protects privacy while keeping "
    "all analytical relationships intact.",
    BLUE)

info_box(s, 6.55, 0.92, 6.5, 1.35,
    "How It Works  —  HMAC-SHA256",
    "DETERMINISTIC: same input always produces the same token (JOINs work)\n"
    "IRREVERSIBLE: cannot be decoded back to the original value\n"
    "SALT: a secret key makes tokens unique to this project only",
    GREEN)

# Masking table
hdrs = ["Field", "Entity", "Technique", "Before (Raw)", "After (Masked)"]
ws   = [2.0, 1.7, 2.2, 3.3, 3.1]
trow(s, hdrs, ws, 0.2, 2.45, 0.4, NAVY, WHITE, sz=11, bold=True)

mrows = [
    ("member_id",    "members",   "Deterministic token",   "M0000001",          "MBR3F9A2C1D4E"),
    ("first_name",   "members",   "Deterministic token",   "Alice",             "FN8B2C3D4E5F6A"),
    ("email",        "members",   "Deterministic token",   "alice@gmail.com",   "EMAIF29A3B1C2D"),
    ("phone",        "members",   "Deterministic token",   "9876543210",        "PH8A2B3C4D5E"),
    ("dob",          "members",   "Year only (generalise)","1985-03-10",        "1985"),
    ("tax_id",       "providers", "Deterministic token",   "908430420",         "TAX2B3C4D5E6F"),
    ("claim_number", "claims",    "Deterministic token",   "CL00155929",        "CLM4E5F6A7B8C"),
    ("city / state", "members",   "Kept as-is",            "Mumbai, Maharashtra","Mumbai, Maharashtra"),
]
for i, row in enumerate(mrows):
    bg = WHITE if i%2==0 else LTGRAY
    # highlight the "Kept as-is" row
    if "Kept" in row[2]: bg = RGBColor(0xE8,0xF5,0xE9)
    trow(s, row, ws, 0.2, 2.88+i*0.44, 0.41, bg, DARK, sz=10)

T(s, "Fields kept clear: city, state, gender, specialty, procedure_code, "
     "diagnosis_code, amounts, payment_method, dates (non-DOB)",
  0.2, 6.65, 12.9, 0.38, sz=10, italic=True, color=GRAY)


# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 10 — ANALYTICAL SQL
# ═══════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
light_bg(s)
slide_title(s, "Analytical SQL Queries", "10 business questions answered from the warehouse")

queries = [
    ("Q1", "Total Claims Summary",     "COUNT, SUM, AVG on fact_claim",                 BLUE),
    ("Q2", "Monthly Paid Trend",       "JOIN fact_claim + dim_date, GROUP BY year/month",TEAL),
    ("Q3", "Top 10 Procedures",        "GROUP BY procedure_code ORDER BY COUNT(*)",      GREEN),
    ("Q4", "Provider Scorecard",       "JOIN dim_provider, AVG paid, denial_rate %",     ORANGE),
    ("Q5", "Denial Rate Breakdown",    "CASE WHEN denial_code, GROUP BY denial_reason",  RED),
    ("Q6", "Length of Stay Dist.",     "GROUP BY length_of_stay, AVG paid_amount",       PURPLE),
    ("Q7", "Member Utilisation",       "JOIN dim_member, SUM paid, ORDER BY spend",      NAVY),
    ("Q8", "Geographic Analysis",      "JOIN dim_member on city/state, SUM by state",    BLUE),
    ("Q9", "Payment Method Split",     "GROUP BY payment_method, SUM paid_amount",       TEAL),
    ("Q10","Diagnosis Cost Analysis",  "Link diagnoses to claims, cost by ICD code",     ORANGE),
]
for i, (num, title, detail, col) in enumerate(queries):
    x = 0.2 + (i%2) * 6.55
    y = 1.0  + (i//2) * 1.2
    R(s, x, y, 6.35, 1.05, LTGRAY)
    R(s, x, y, 0.65, 1.05, col)
    T(s, num,   x+0.08, y+0.1,  0.5, 0.5,  sz=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    T(s, title, x+0.82, y+0.07, 5.3, 0.38, sz=13, bold=True, color=NAVY)
    T(s, detail,x+0.82, y+0.48, 5.3, 0.45, sz=10, italic=True, color=GRAY)

T(s, "All queries run against the live SQLite warehouse via Python / Jupyter Notebook",
  0.2, 7.05, 12.9, 0.35, sz=11, italic=True, color=GRAY, align=PP_ALIGN.CENTER)


# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 11 — KPI RESULTS
# ═══════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
light_bg(s)
slide_title(s, "Key Business Metrics", "Live results from the warehouse after full ETL load")

kpis_data = [
    ("Total Claims",       "2,08,640",   "",              NAVY),
    ("Total Allowed",      "Rs 20.9B",   "billed",        BLUE),
    ("Total Paid",         "Rs 15.68B",  "paid out",      GREEN),
    ("Avg per Claim",      "Rs 75,156",  "average",       TEAL),
    ("Denial Rate",        "49.92%",     "claims denied", RED),
    ("Top State",          "Punjab",     "highest spend", ORANGE),
    ("Top Procedure",      "CPT35289",   "417 uses",      PURPLE),
    ("Payment Methods",    "4 types",    "equal split",   BLUE),
]
for i, (lbl, val, sub, col) in enumerate(kpis_data):
    x = 0.2 + (i%4) * 3.27
    y = 1.0 + (i//4) * 2.5
    kpi(s, x, y, 3.1, 2.1, lbl, val, sub, col)

T(s, "Source: warehouse.db  |  208,640 valid claims  |  ~12.5 million total rows across all tables",
  0.2, 6.45, 12.9, 0.4, sz=11, italic=True, color=GRAY, align=PP_ALIGN.CENTER)


# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 12 — VISUALIZATIONS (4 charts: monthly + procedures + denial + LOS)
# ═══════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
light_bg(s)
slide_title(s, "Analytics Visualizations  –  Part 1 of 2",
            "Monthly trend  |  Top procedures  |  Denial rate  |  Length of stay")

chart_grid = [
    ("01_monthly_paid.png",    "Monthly Paid Amount",       "Consistent ~Rs 540M/month",     0.2, 0.92),
    ("02_top_procedures.png",  "Top 10 Procedures",         "CPT35289 highest at 417 uses",  6.77, 0.92),
    ("04_denial_rate.png",     "Denial Rate by Code",       "49.9% denied — D001/D002/D003", 0.2, 4.1),
    ("05_los_distribution.png","Length of Stay (0-30 days)","Flat distribution across LOS",  6.77, 4.1),
]
for (fname, title, insight, x, y) in chart_grid:
    R(s, x, y, 6.38, 2.95, LTGRAY)
    R(s, x, y, 6.38, 0.32, NAVY)
    T(s, title,   x+0.12, y+0.04, 5.9, 0.25, sz=11, bold=True, color=WHITE)
    T(s, insight, x+0.12, y+0.34, 5.9, 0.25, sz=9,  italic=True, color=GRAY)
    IMG(s, CHARTS_DIR / fname, x+0.08, y+0.62, 6.22, 2.22)


# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 13 — VISUALIZATIONS (4 charts: provider + geo + member + payment)
# ═══════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
light_bg(s)
slide_title(s, "Analytics Visualizations  –  Part 2 of 2",
            "Provider performance  |  Geographic  |  Member utilisation  |  Payment methods")

chart_grid2 = [
    ("03_provider_performance.png","Provider Performance",  "Some providers >60% denial rate", 0.2, 0.92),
    ("07_geographic.png",          "Paid by State",         "Punjab & Odisha lead",            6.77, 0.92),
    ("06_member_utilisation.png",  "Member Utilisation",    "High-spend members by gender",    0.2, 4.1),
    ("08_payment_methods.png",     "Payment Methods",       "NEFT/UPI/RTGS/Cheque equal split",6.77, 4.1),
]
for (fname, title, insight, x, y) in chart_grid2:
    R(s, x, y, 6.38, 2.95, LTGRAY)
    R(s, x, y, 6.38, 0.32, NAVY)
    T(s, title,   x+0.12, y+0.04, 5.9, 0.25, sz=11, bold=True, color=WHITE)
    T(s, insight, x+0.12, y+0.34, 5.9, 0.25, sz=9,  italic=True, color=GRAY)
    IMG(s, CHARTS_DIR / fname, x+0.08, y+0.62, 6.22, 2.22)


# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 14 — OPTIMISATION
# ═══════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
light_bg(s)
slide_title(s, "Optimisation", "Database indexes + vectorised Pandas + masking impact")

# Indexes
R(s, 0.2, 0.92, 6.15, 3.45, RGBColor(0xE3,0xF2,0xFD))
R(s, 0.2, 0.92, 0.1,  3.45, BLUE)
T(s, "Database Indexes  (in DDL)", 0.45, 0.98, 5.8, 0.38, sz=13, bold=True, color=BLUE)

idx_data = [
    ("idx_fc_member",    "fact_claim",       "member_id",       "Member JOIN"),
    ("idx_fc_provider",  "fact_claim",       "provider_id",     "Provider JOIN"),
    ("idx_fc_admit",     "fact_claim",       "admit_date_key",  "Date filter"),
    ("idx_fc_denied",    "fact_claim",       "is_denied",       "Denial queries"),
    ("idx_fcl_claim",    "fact_claim_line",  "claim_number",    "Line JOIN"),
    ("idx_fp_claim",     "fact_payment",     "claim_number",    "Payment JOIN"),
    ("idx_dp_specialty", "dim_provider",     "specialty",       "Specialty filter"),
]
iws = [2.1, 2.0, 1.55]
trow(s, ["Index", "Table", "Column"], iws, 0.4, 1.4, 0.32, BLUE, WHITE, sz=9, bold=True)
for i, (idx, tbl, col, _) in enumerate(idx_data):
    bg = WHITE if i%2==0 else LTGRAY
    trow(s, [idx, tbl, col], iws, 0.4, 1.74+i*0.3, 0.28, bg, DARK, sz=9)

# Pandas vectorisation
R(s, 6.55, 0.92, 6.55, 3.45, RGBColor(0xE8,0xF5,0xE9))
R(s, 6.55, 0.92, 0.1,  3.45, GREEN)
T(s, "Vectorised Pandas Operations", 6.8, 0.98, 6.1, 0.38, sz=13, bold=True, color=GREEN)

vec = [
    ("pd.to_numeric(col, errors='coerce')",  "Converts entire column at once — no row loops"),
    ("pd.to_datetime(col, errors='coerce')", "Parses all dates in single operation"),
    ("df[boolean_mask]",                     "Filters millions of rows instantly"),
    ("Series.map(dictionary)",               "Remaps FK columns — 100x faster than apply"),
    ("pd.concat + drop_duplicates",          "Deduplication across full entity datasets"),
]
y = 1.42
for code, desc in vec:
    R(s, 6.7, y, 6.2, 0.23, RGBColor(0xD0,0xF0,0xD8))
    T(s, code, 6.78, y+0.03, 6.0, 0.18, sz=9, bold=True, color=DARK)
    T(s, desc, 6.78, y+0.24, 6.0, 0.18, sz=9, italic=True, color=GRAY)
    y += 0.52

# Masking impact
R(s, 0.2, 4.55, 12.9, 1.55, RGBColor(0xFF,0xF8,0xE1))
R(s, 0.2, 4.55, 0.1,  1.55, GOLD)
T(s, "Masking Impact on Analytics", 0.45, 4.62, 9, 0.38, sz=13, bold=True, color=ORANGE)
T(s, "Deterministic masking means member_id 'M0000001' ALWAYS becomes 'MBR3F9A2C1D4E'  ->  "
     "JOINs between fact_claim, dim_member and dim_provider work correctly.\n"
     "City/state kept clear  ->  geographic analysis fully functional.   "
     "DOB -> birth_year  ->  age band analysis still possible.",
  0.45, 5.05, 12.4, 0.9, sz=12, color=DARK)

T(s, "Full ETL run: ~3-5 min on laptop  |  Incremental: seconds  |  Warehouse: 9 tables, ~12.5M rows",
  0.2, 6.28, 12.9, 0.38, sz=11, italic=True, color=GRAY, align=PP_ALIGN.CENTER)


# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 15 — DELIVERABLES
# ═══════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
light_bg(s)
slide_title(s, "Project Deliverables", "All required outputs per the capstone document")

all_deliverables = [
    ("Star Schema Design",           "01_star_schema_ddl.sql — 9 tables with FK indexes",    BLUE,   True),
    ("ER Diagram",                   "Entity relationships across all 7 source tables",        TEAL,   True),
    ("Analytical SQL Queries",       "02_analytical_queries.sql — 10 business queries",        BLUE,   True),
    ("Data Profiling Report",        "logs/profiling_summary.json — quality findings",         ORANGE, True),
    ("ETL Full Load Script",         "step4_load_full.py — drops + reloads everything",        NAVY,   True),
    ("ETL Incremental Load",         "step5_load_incremental.py — watermark-based new rows",   NAVY,   True),
    ("Data Masking Logic",           "step3_mask.py — HMAC-SHA256 deterministic hashing",      PURPLE, True),
    ("Masking Audit Log",            "logs/masking_audit.json — fields masked + counts",       PURPLE, True),
    ("Error / Rejection Logs",       "output/rejected_rows/*.csv — 7 audit files",             RED,    True),
    ("Python Analytics",             "analytics.py — 9 queries + CSV results",                 GREEN,  True),
    ("8 Matplotlib Charts",          "output/charts/*.png — all visualisations saved",         GREEN,  True),
    ("Jupyter Notebook",             "analytics_notebook.ipynb — interactive analysis",        TEAL,   True),
    ("Optimisation Summary",         "Indexes + vectorised Pandas + masking impact",           ORANGE, True),
    ("GitHub Repository",            "https://github.com/mhgowda/healthcare-insurance-dw-capstone", BLUE, True),
]
for i, (title, detail, col, done) in enumerate(all_deliverables):
    x = 0.2 + (i%2) * 6.55
    y = 0.95 + (i//2) * 0.55
    R(s, x, y, 6.35, 0.48, WHITE if i%2==0 else LTGRAY)
    R(s, x, y, 0.08, 0.48, col)
    R(s, x+0.15, y+0.12, 0.24, 0.24, GREEN)
    T(s, "OK", x+0.16, y+0.12, 0.22, 0.22, sz=8, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    T(s, title,  x+0.48, y+0.03, 2.7, 0.23, sz=11, bold=True, color=NAVY)
    T(s, detail, x+0.48, y+0.26, 5.7, 0.2,  sz=9,  italic=True, color=GRAY)

T(s, "GitHub: https://github.com/mhgowda/healthcare-insurance-dw-capstone",
  0.2, 7.05, 12.9, 0.35, sz=11, bold=True, color=BLUE, align=PP_ALIGN.CENTER)


# ═══════════════════════════════════════════════════════════════════════════
# SLIDE 16 — THANK YOU
# ═══════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
dark_bg(s)

T(s, "Thank You", 0.5, 1.5, 10.3, 1.6,
  sz=72, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
T(s, "Questions & Discussion", 0.5, 3.15, 10.3, 0.8,
  sz=26, color=TEAL, align=PP_ALIGN.CENTER)

R(s, 2.0, 4.2, 9.0, 0.06, GOLD)

T(s, "GitHub Repository", 0.5, 4.45, 10.3, 0.4,
  sz=13, color=GRAY, align=PP_ALIGN.CENTER)
T(s, "https://github.com/mhgowda/healthcare-insurance-dw-capstone",
  0.5, 4.9, 10.3, 0.45, sz=17, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

# Stats row
for i, (val, lbl) in enumerate([
    ("1.4M",  "Source Rows"),
    ("9",     "DW Tables"),
    ("17.9%", "Rejected"),
    ("49.9%", "Denial Rate"),
    ("8",     "Charts"),
]):
    x = 0.5 + i * 2.45
    R(s, x, 5.7, 2.3, 1.35, RGBColor(0x10,0x23,0x4E))
    T(s, val, x+0.1, 5.78, 2.1, 0.65, sz=26, bold=True, color=GOLD, align=PP_ALIGN.CENTER)
    T(s, lbl, x+0.1, 6.42, 2.1, 0.35, sz=10, color=LTBLUE, align=PP_ALIGN.CENTER)

T(s, "Midterm Capstone  |  Group 5  |  Batch 2024-25",
  0.5, 7.12, 10.3, 0.35, sz=11, color=GRAY, align=PP_ALIGN.CENTER)


# ═══════════════════════════════════════════════════════════════════════════
# SAVE
# ═══════════════════════════════════════════════════════════════════════════
prs.save(str(OUT_FILE))
print(f"\nPPT saved: {OUT_FILE}")
print(f"Total slides: {len(prs.slides)}")
