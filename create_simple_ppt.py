"""
Q2 2026 Professional Work Summary — Simple & Professional Theme
"""

import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

# ── Theme Colors ──────────────────────────────────────────────
WHITE      = RGBColor(255, 255, 255)
CHARCOAL   = RGBColor(50, 50, 50)
CORP_BLUE  = RGBColor(31, 73, 125)
LIGHT_BLUE = RGBColor(240, 246, 250)
LIGHT_GRAY = RGBColor(235, 235, 235)
MED_GRAY   = RGBColor(120, 120, 120)

prs = Presentation()
prs.slide_width  = Inches(13.333)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]

W = prs.slide_width
H = prs.slide_height

def add_bg(slide, color=WHITE):
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color

def add_shape(slide, left, top, width, height, color, line_color=None):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    if line_color:
        shape.line.color.rgb = line_color
        shape.line.width = Pt(1)
    else:
        shape.line.fill.background()
    shape.shadow.inherit = False
    return shape

def add_rounded_rect(slide, left, top, width, height, color, line_color=None):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    if line_color:
        shape.line.color.rgb = line_color
        shape.line.width = Pt(1)
    else:
        shape.line.fill.background()
    shape.shadow.inherit = False
    return shape

def add_textbox(slide, left, top, width, height, text, size=18,
                 color=CHARCOAL, bold=False, align=PP_ALIGN.LEFT, font_name="Calibri"):
    txBox = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(size)
    p.font.color.rgb = color
    p.font.bold = bold
    p.font.name = font_name
    p.alignment = align
    return txBox

def section_header(slide, title, subtitle="", top=0.4):
    add_textbox(slide, 0.8, top, 11, 0.6,
                 title, size=28, color=CORP_BLUE, bold=True, font_name="Calibri")
    add_shape(slide, 0.8, top + 0.65, 3.0, 0.05, CORP_BLUE)
    if subtitle:
        add_textbox(slide, 0.8, top + 0.75, 11, 0.5,
                     subtitle, size=14, color=MED_GRAY, font_name="Calibri")

def add_bullet_list(slide, items, left, top, width, height, size=12, color=CHARCOAL):
    tb = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = tb.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.text = item
        p.font.size = Pt(size)
        p.font.color.rgb = color
        p.font.name = "Calibri"
        p.space_after = Pt(4)
        p.space_before = Pt(4)

# SLIDE 1: Title
s1 = prs.slides.add_slide(BLANK)
add_bg(s1, WHITE)
add_shape(s1, 0, 0, 13.333, 0.2, CORP_BLUE)
add_shape(s1, 0, 7.3, 13.333, 0.2, CORP_BLUE)
add_textbox(s1, 1.0, 2.0, 11, 0.5, "DETAILED WORK SUMMARY", size=18, color=MED_GRAY, bold=True)
add_textbox(s1, 1.0, 2.5, 11, 1.2, "Q2 2026: April | May | June", size=48, color=CORP_BLUE, bold=True)
add_shape(s1, 1.0, 4.0, 4.0, 0.05, CORP_BLUE)
add_textbox(s1, 1.0, 4.3, 10, 0.8, "Comprehensive breakdown of engineering, database optimization, frontend development, and AI integration across all workspaces.", size=16, color=CHARCOAL)
add_textbox(s1, 1.0, 5.8, 10, 0.4, "Jasil N Balussery | Technology / Software Development | myG", size=14, color=CHARCOAL, bold=True)

# SLIDE 2: Exec Summary
s2 = prs.slides.add_slide(BLANK)
add_bg(s2, WHITE)
add_shape(s2, 0, 0, 0.2, 7.5, CORP_BLUE)
section_header(s2, "Executive Summary & Key Metrics")
add_textbox(s2, 0.8, 1.5, 11.5, 1.0, "During Q2 2026, extensive full-stack development, database engineering, AI/ML integration, and multiple greenfield product launches were delivered across 8 distinct projects. Work spanned Python (Flask/Django), PostgreSQL, JavaScript, Machine Learning, LLM integration, Google APIs, and cloud deployment.", size=15)

metrics = [("8", "Projects"), ("80+", "Dev Sessions"), ("12.6M+", "DB Rows"), ("26", "Mat Views"), ("< 1s", "AI Response"), ("4", "ML Models")]
for i, (val, lbl) in enumerate(metrics):
    row = i // 3
    col = i % 3
    lx = 0.8 + col * 3.9
    ty = 2.8 + row * 1.8
    add_rounded_rect(s2, lx, ty, 3.5, 1.5, LIGHT_BLUE, CORP_BLUE)
    add_textbox(s2, lx, ty + 0.2, 3.5, 0.6, val, size=32, color=CORP_BLUE, bold=True, align=PP_ALIGN.CENTER)
    add_textbox(s2, lx, ty + 0.9, 3.5, 0.5, lbl, size=15, color=CHARCOAL, align=PP_ALIGN.CENTER)

# SLIDE 3: OSG Portal
s3 = prs.slides.add_slide(BLANK)
add_bg(s3, WHITE)
add_shape(s3, 0, 0, 0.2, 7.5, CORP_BLUE)
section_header(s3, "1. OSG-myG-PORTAL", "Claims Management & WhatsApp Automation")

cards3 = [
    ("WhatsApp API Integration (Telinfy)", [
        "• Engineered automated messaging pipelines for portal events",
        "• Configured 4 WhatsApp templates: Registered, Replacement, Repair Completed, Spare Parts",
        "• Debugged payload routing for reliable outbound notifications",
        "• Built trigger buttons in portal for WhatsApp message dispatch",
    ]),
    ("Claims Processing & Data Forensics", [
        "• Developed investigative scripts to trace and resolve claim anomalies",
        "• Built repair utilities to retroactively fix corrupted claim states",
        "• SR Number auto-generation on status change (Submit → Registered)",
        "• Google Sheets bi-directional sync for claim tracking",
    ])
]
for i, (title, points) in enumerate(cards3):
    lx = 0.8 + i * 5.9
    add_rounded_rect(s3, lx, 1.8, 5.6, 4.5, WHITE, LIGHT_GRAY)
    add_shape(s3, lx, 1.8, 5.6, 0.5, CORP_BLUE)
    add_textbox(s3, lx + 0.2, 1.9, 5.2, 0.4, title, size=14, color=WHITE, bold=True)
    add_bullet_list(s3, points, lx + 0.2, 2.5, 5.2, 3.5, size=13)

# SLIDE 4: AI Agent
s4 = prs.slides.add_slide(BLANK)
add_bg(s4, WHITE)
add_shape(s4, 0, 0, 0.2, 7.5, CORP_BLUE)
section_header(s4, "2. Enterprise AI Agent", "Loyalty Portal — Natural Language Business Intelligence")

cards4 = [
    ("Multi-Layered Agent Architecture", [
        "• SQL Agent: Translates natural language into complex PostgreSQL queries",
        "• AI Analyst: Interprets raw DB output into readable business insights",
        "• FastPath Engine: Common queries answered in < 0.2 seconds",
        "• 16-phase implementation covering Foundation to Export System",
        "• Schema context injection prevents AI hallucinations",
    ]),
    ("LLM Integration & Optimization", [
        "• NVIDIA NIM API (Nemotron Ultra) for SQL generation",
        "• OpenRouter API as fallback LLM provider",
        "• Response time reduced from 5-10 minutes to < 1 second",
        "• Aggressive prompt engineering & query scope reduction",
        "• Validated with 20+ test business questions",
    ])
]
for i, (title, points) in enumerate(cards4):
    lx = 0.8 + i * 5.9
    add_rounded_rect(s4, lx, 1.8, 5.6, 4.5, WHITE, LIGHT_GRAY)
    add_shape(s4, lx, 1.8, 5.6, 0.5, CORP_BLUE)
    add_textbox(s4, lx + 0.2, 1.9, 5.2, 0.4, title, size=14, color=WHITE, bold=True)
    add_bullet_list(s4, points, lx + 0.2, 2.5, 5.2, 3.5, size=13)

# SLIDE 5: Database
s5 = prs.slides.add_slide(BLANK)
add_bg(s5, WHITE)
add_shape(s5, 0, 0, 0.2, 7.5, CORP_BLUE)
section_header(s5, "3. Database Administration & Optimization", "12.6M+ Rows · DigitalOcean PostgreSQL")

cards5 = [
    ("Materialized View Overhaul", [
        "• Diagnosed critical cross-join bug in mv_yearly_cohort",
        "• Built robust refresh system for 26 materialized views",
        "• Concurrent regeneration without database locks",
    ]),
    ("Caching Infrastructure", [
        "• Integrated Redis for production caching",
        "• LocMemCache for local development",
        "• 12.6M+ row calculations reduced to sub-second responses",
    ]),
    ("Data Operations", [
        "• Custom scripts bypass web-server timeouts for large uploads",
        "• Automated scrubbing of anomalous data (SMC/EI, HEAD OFFICE)",
        "• Data integrity verification: 5,033,297+ total records",
    ])
]
for i, (title, points) in enumerate(cards5):
    lx = 0.8 + i * 3.9
    add_rounded_rect(s5, lx, 1.8, 3.7, 4.5, WHITE, LIGHT_GRAY)
    add_shape(s5, lx, 1.8, 3.7, 0.5, CORP_BLUE)
    add_textbox(s5, lx + 0.2, 1.9, 3.3, 0.4, title, size=14, color=WHITE, bold=True)
    add_bullet_list(s5, points, lx + 0.2, 2.5, 3.3, 3.5, size=13)

# SLIDE 6: ML
s6 = prs.slides.add_slide(BLANK)
add_bg(s6, WHITE)
add_shape(s6, 0, 0, 0.2, 7.5, CORP_BLUE)
section_header(s6, "4. Machine Learning & Predictive Forecasting", "Live Forecasting Engines")

cards6 = [
    ("Model Deployment (Scikit-Learn)", [
        "• Random Forest — Ensemble decision tree model",
        "• MLPRegressor — Neural Network proxy for patterns",
        "• GradientBoostingRegressor — Sequential boosting",
        "• LSTM deep learning model on historical data",
    ]),
    ("Feature Engineering", [
        "• MalayalamCalendarFeaturizer — Kerala-specific feature engine",
        "• Converts dates to arrays with festival proximity (Onam, Vishu)",
        "• Real-time sales prediction from mid-day actuals",
    ]),
    ("Dormant Customer Reactivation", [
        "• SQL pre-aggregation engines for dormant customer bucketing",
        "• Cohort-based analysis across 2020-2025 cohorts",
        "• UI Visualizers: Probability Gauges, Semi-Donut Charts",
    ])
]
for i, (title, points) in enumerate(cards6):
    lx = 0.8 + i * 3.9
    add_rounded_rect(s6, lx, 1.8, 3.7, 4.5, WHITE, LIGHT_GRAY)
    add_shape(s6, lx, 1.8, 3.7, 0.5, CORP_BLUE)
    add_textbox(s6, lx + 0.2, 1.9, 3.3, 0.4, title, size=14, color=WHITE, bold=True)
    add_bullet_list(s6, points, lx + 0.2, 2.5, 3.3, 3.5, size=13)

# SLIDE 7: SHE START
s7 = prs.slides.add_slide(BLANK)
add_bg(s7, WHITE)
add_shape(s7, 0, 0, 0.2, 7.5, CORP_BLUE)
section_header(s7, "5. SHE START — Evaluation Dashboard", "Startup Program Evaluation System")

cards7 = [
    ("Live Google Sheets Sync", [
        "• gspread library for bidirectional sync (25-second refresh)",
        "• Auto-detection of new applicants",
        "• District & business name auto-population",
    ]),
    ("6-Panelist Scoring Algorithm", [
        "• 6 panelists score each applicant",
        "• Auto-drop highest & lowest scores to prevent bias",
        "• Weighted Final Score calculation (Interview, Growth, etc.)",
    ]),
    ("Interactive Dashboard", [
        "• Inline cell editing for 5 score columns with silent auto-save",
        "• Auto badging: Strong Selection, Waitlist",
        "• Dedicated user roles & Excel report download",
    ])
]
for i, (title, points) in enumerate(cards7):
    lx = 0.8 + i * 3.9
    add_rounded_rect(s7, lx, 1.8, 3.7, 4.5, WHITE, LIGHT_GRAY)
    add_shape(s7, lx, 1.8, 3.7, 0.5, CORP_BLUE)
    add_textbox(s7, lx + 0.2, 1.9, 3.3, 0.4, title, size=14, color=WHITE, bold=True)
    add_bullet_list(s7, points, lx + 0.2, 2.5, 3.3, 3.5, size=13)

# SLIDE 8: Retail Dashboard & Bigg Boss
s8 = prs.slides.add_slide(BLANK)
add_bg(s8, WHITE)
add_shape(s8, 0, 0, 0.2, 7.5, CORP_BLUE)
section_header(s8, "6 & 7. Retail Analytics & Event Portals")

cards8 = [
    ("Enterprise Retail Dashboard", [
        "• High-Speed Reporting: Rust-based calamine engine (5-10x faster)",
        "• DataTables integration for dynamic, searchable grids",
        "• New Sections: Monthly Retention, Campaign Analysis",
        "• Advanced Metrics: Dynamic ASP mapping, Cross-referencing",
    ]),
    ("BIGG BOSS & FoneFlix Registration", [
        "• Custom OAuth 2.0 Client ID for Google Drive video uploads",
        "• Responsive hero section with OTT reality-show aesthetics",
        "• Project Pivot: Successfully transitioned from Bigg Boss to FoneFlix",
        "• Refactored form inputs, consent clauses, and branding",
    ])
]
for i, (title, points) in enumerate(cards8):
    lx = 0.8 + i * 5.9
    add_rounded_rect(s8, lx, 1.8, 5.6, 4.5, WHITE, LIGHT_GRAY)
    add_shape(s8, lx, 1.8, 5.6, 0.5, CORP_BLUE)
    add_textbox(s8, lx + 0.2, 1.9, 5.2, 0.4, title, size=14, color=WHITE, bold=True)
    add_bullet_list(s8, points, lx + 0.2, 2.5, 5.2, 3.5, size=13)

# SLIDE 9: Projects Summary
s9 = prs.slides.add_slide(BLANK)
add_bg(s9, WHITE)
add_shape(s9, 0, 0, 0.2, 7.5, CORP_BLUE)
section_header(s9, "All Projects Delivered — Q2 2026")

projects = [
    ("OSG-myG-PORTAL", "Claims Management + WhatsApp Automation", "Render"),
    ("myG Loyalty Dashboard", "Enterprise Analytics + AI Agent", "Render"),
    ("SHE START Dashboard", "Startup Evaluation System", "Render"),
    ("BIGBOSS / FoneFlix", "Registration + Video Upload Portal", "Render"),
    ("HR E-Signature Portal", "Employee Signature System", "Render"),
    ("FIFA World Cup Contest", "Prediction & Leaderboard Platform", "Vercel"),
    ("Enterprise AI Agent", "Natural Language -> SQL Business Intel", "Integrated"),
    ("n8n Chat Agent", "Automated AI Chat Widget", "n8n Cloud")
]

add_shape(s9, 0.8, 1.6, 11.5, 0.5, CORP_BLUE)
add_textbox(s9, 1.0, 1.65, 3.0, 0.4, "Project", size=14, color=WHITE, bold=True)
add_textbox(s9, 4.5, 1.65, 5.0, 0.4, "Description", size=14, color=WHITE, bold=True)
add_textbox(s9, 10.0, 1.65, 2.0, 0.4, "Platform", size=14, color=WHITE, bold=True)

for idx, (p_name, p_desc, p_plat) in enumerate(projects):
    ty = 2.2 + idx * 0.55
    bg_color = LIGHT_BLUE if idx % 2 == 0 else WHITE
    add_shape(s9, 0.8, ty, 11.5, 0.5, bg_color, LIGHT_GRAY)
    add_textbox(s9, 1.0, ty + 0.1, 3.3, 0.4, p_name, size=13, color=CHARCOAL, bold=True)
    add_textbox(s9, 4.5, ty + 0.1, 5.3, 0.4, p_desc, size=13, color=CHARCOAL)
    add_textbox(s9, 10.0, ty + 0.1, 2.0, 0.4, p_plat, size=13, color=CHARCOAL)

# SLIDE 10: Thank you
s10 = prs.slides.add_slide(BLANK)
add_bg(s10, WHITE)
add_shape(s10, 0, 0, 13.333, 0.2, CORP_BLUE)
add_shape(s10, 0, 7.3, 13.333, 0.2, CORP_BLUE)

add_textbox(s10, 0, 3.0, 13.333, 1.0, "THANK YOU", size=54, color=CORP_BLUE, bold=True, align=PP_ALIGN.CENTER)
add_textbox(s10, 0, 4.0, 13.333, 0.5, "Jasil N Balussery | Technology / Software Development", size=18, color=MED_GRAY, align=PP_ALIGN.CENTER)
add_textbox(s10, 0, 4.5, 13.333, 0.5, "April | May | June 2026", size=16, color=MED_GRAY, align=PP_ALIGN.CENTER)

OUTPUT = r"c:\Users\jasil_myg\Desktop\OSG-myG-PORTAL-mainnnnn - Copy\Q2_2026_Work_Summary_Simple_Professional.pptx"
prs.save(OUTPUT)
print(f"Saved: {OUTPUT}")
