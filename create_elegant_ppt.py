"""
Q2 2026 Professional Work Summary — Elegant Corporate Theme
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
OFF_WHITE  = RGBColor(248, 249, 250)
CHARCOAL   = RGBColor(40, 45, 50)
CORP_BLUE  = RGBColor(18, 53, 91)
ACCENT_BLUE= RGBColor(46, 117, 182)
LIGHT_GRAY = RGBColor(220, 224, 229)
MED_GRAY   = RGBColor(100, 110, 120)

prs = Presentation()
prs.slide_width  = Inches(13.333)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]

def add_shape(slide, mso_shape, left, top, width, height, fill_color, line_color=None):
    shape = slide.shapes.add_shape(mso_shape, Inches(left), Inches(top), Inches(width), Inches(height))
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    if line_color:
        shape.line.color.rgb = line_color
        shape.line.width = Pt(1)
    else:
        shape.line.fill.background()
    shape.shadow.inherit = False
    return shape

def add_textbox(slide, left, top, width, height, text, size=18, color=CHARCOAL, bold=False, align=PP_ALIGN.LEFT):
    txBox = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(size)
    p.font.color.rgb = color
    p.font.bold = bold
    p.font.name = "Calibri"
    p.alignment = align
    return txBox

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
        p.space_after = Pt(6)
        p.space_before = Pt(2)

def slide_template(slide, title, subtitle):
    # Background
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, 13.333, 7.5, OFF_WHITE)
    # Header Band
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, 13.333, 1.4, CORP_BLUE)
    # Accent Line
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 1.35, 13.333, 0.05, ACCENT_BLUE)
    
    # Title & Subtitle
    add_textbox(slide, 0.5, 0.2, 12, 0.6, title, size=32, color=WHITE, bold=True)
    if subtitle:
        add_textbox(slide, 0.5, 0.75, 12, 0.4, subtitle, size=14, color=LIGHT_GRAY)
        
    # Footer
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 7.2, 13.333, 0.3, CORP_BLUE)
    add_textbox(slide, 0.5, 7.22, 12, 0.3, "myG Technology | Jasil N Balussery | Q2 2026", size=10, color=WHITE, align=PP_ALIGN.CENTER)

# --- SLIDE 1: TITLE ---
s1 = prs.slides.add_slide(BLANK)
# Split screen: Left Blue, Right Off-White
add_shape(s1, MSO_SHAPE.RECTANGLE, 0, 0, 5.0, 7.5, CORP_BLUE)
add_shape(s1, MSO_SHAPE.RECTANGLE, 5.0, 0, 8.333, 7.5, OFF_WHITE)
# Graphics
add_shape(s1, MSO_SHAPE.RIGHT_TRIANGLE, 5.0, 0, 1.5, 7.5, CORP_BLUE)
add_shape(s1, MSO_SHAPE.OVAL, -1.0, -1.0, 3.0, 3.0, ACCENT_BLUE)

add_textbox(s1, 0.5, 0.5, 4.0, 0.5, "myG Technology", size=16, color=LIGHT_GRAY, bold=True)
add_textbox(s1, 5.8, 2.0, 7.0, 1.5, "DETAILED WORK\nSUMMARY", size=54, color=CORP_BLUE, bold=True)
add_shape(s1, MSO_SHAPE.RECTANGLE, 5.8, 4.0, 1.5, 0.08, ACCENT_BLUE)
add_textbox(s1, 5.8, 4.3, 7.0, 0.8, "Q2 2026: April | May | June", size=24, color=CHARCOAL, bold=True)
add_textbox(s1, 5.8, 5.5, 7.0, 1.0, "Jasil N Balussery\nTechnology / Software Development", size=16, color=MED_GRAY)

# --- SLIDE 2: EXEC SUMMARY ---
s2 = prs.slides.add_slide(BLANK)
slide_template(s2, "Executive Summary & Key Metrics", "Overview of accomplishments and high-level KPIs")

add_textbox(s2, 0.5, 1.6, 12.3, 1.0, "During Q2 2026, extensive full-stack development, database engineering, AI/ML integration, and multiple greenfield product launches were delivered across 8 distinct projects. Work spanned Python (Flask/Django), PostgreSQL, JavaScript, Machine Learning, LLM integration, Google APIs, and cloud deployment.", size=16, color=CHARCOAL)

metrics = [("8", "Projects"), ("80+", "Dev Sessions"), ("12.6M+", "DB Rows"), ("26", "Mat Views"), ("< 1s", "AI Response"), ("4", "ML Models")]
for i, (val, lbl) in enumerate(metrics):
    row = i // 3
    col = i % 3
    lx = 0.5 + col * 4.2
    ty = 3.0 + row * 1.8
    add_shape(s2, MSO_SHAPE.ROUNDED_RECTANGLE, lx, ty, 3.8, 1.5, WHITE, LIGHT_GRAY)
    add_shape(s2, MSO_SHAPE.OVAL, lx + 0.2, ty + 0.3, 0.9, 0.9, ACCENT_BLUE)
    add_textbox(s2, lx + 0.2, ty + 0.45, 0.9, 0.9, val, size=18, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    add_textbox(s2, lx + 1.2, ty + 0.4, 2.5, 0.6, val, size=28, color=CORP_BLUE, bold=True)
    add_textbox(s2, lx + 1.2, ty + 0.9, 2.5, 0.5, lbl, size=14, color=MED_GRAY)

# --- HELPER FOR CONTENT SLIDES ---
def add_card(slide, left, top, width, height, title, points):
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height, WHITE, LIGHT_GRAY)
    add_shape(slide, MSO_SHAPE.RECTANGLE, left, top, 0.15, height, ACCENT_BLUE) # left accent bar
    add_textbox(slide, left + 0.3, top + 0.1, width - 0.4, 0.4, title, size=16, color=CORP_BLUE, bold=True)
    add_bullet_list(slide, points, left + 0.3, top + 0.6, width - 0.4, height - 0.7, size=13)

# --- SLIDE 3: OSG Portal ---
s3 = prs.slides.add_slide(BLANK)
slide_template(s3, "1. OSG-myG-PORTAL", "Claims Management & WhatsApp Automation")
add_card(s3, 0.5, 1.8, 6.0, 5.0, "WhatsApp API Integration", [
    "Engineered automated messaging pipelines for portal events",
    "Configured 4 WhatsApp templates: Registered, Replacement, Repair Completed, Spare Parts",
    "Debugged payload routing for reliable outbound notifications",
    "Built trigger buttons in portal for WhatsApp message dispatch"
])
add_card(s3, 6.8, 1.8, 6.0, 5.0, "Claims Processing & Data Forensics", [
    "Developed investigative scripts to trace and resolve claim anomalies",
    "Built repair utilities to retroactively fix corrupted claim states",
    "SR Number auto-generation on status change (Submit → Registered)",
    "Google Sheets bi-directional sync for claim tracking",
    "Massive data parsers for legacy Excel imports (17MB+ files)"
])

# --- SLIDE 4: AI Agent ---
s4 = prs.slides.add_slide(BLANK)
slide_template(s4, "2. Enterprise AI Agent", "Loyalty Portal — Natural Language Business Intelligence")
add_card(s4, 0.5, 1.8, 6.0, 5.0, "Multi-Layered Architecture", [
    "SQL Agent: Translates natural language into complex PostgreSQL queries",
    "AI Analyst: Interprets raw DB output into readable business insights",
    "FastPath Engine: Common queries answered in < 0.2 seconds",
    "16-phase implementation covering Foundation to Export System",
    "Schema context injection prevents AI hallucinations"
])
add_card(s4, 6.8, 1.8, 6.0, 5.0, "LLM Integration & Optimization", [
    "NVIDIA NIM API (Nemotron Ultra) for SQL generation",
    "OpenRouter API as fallback LLM provider",
    "Response time reduced from 5-10 minutes to < 1 second",
    "Aggressive prompt engineering & query scope reduction",
    "Validated with 20+ test business questions"
])

# --- SLIDE 5: DB Admin ---
s5 = prs.slides.add_slide(BLANK)
slide_template(s5, "3. Database Administration", "12.6M+ Rows · DigitalOcean PostgreSQL")
add_card(s5, 0.5, 1.8, 3.9, 5.0, "Materialized View Overhaul", [
    "Diagnosed critical cross-join bug in mv_yearly_cohort",
    "Built robust refresh system for 26 materialized views",
    "Concurrent regeneration without database locks"
])
add_card(s5, 4.7, 1.8, 3.9, 5.0, "Caching Infrastructure", [
    "Integrated Redis for production caching",
    "LocMemCache for local development",
    "12.6M+ row calculations reduced to sub-second responses"
])
add_card(s5, 8.9, 1.8, 3.9, 5.0, "Data Operations", [
    "Custom scripts bypass web-server timeouts for large uploads",
    "Automated scrubbing of anomalous data",
    "Data integrity verification: 5M+ total records"
])

# --- SLIDE 6: ML & Predictive Forecasting ---
s6 = prs.slides.add_slide(BLANK)
slide_template(s6, "4. Machine Learning", "Live Forecasting Engines & Data Processing")
add_card(s6, 0.5, 1.8, 3.9, 5.0, "Model Deployment", [
    "Random Forest — Ensemble decision tree model",
    "MLPRegressor — Neural Network proxy for patterns",
    "GradientBoostingRegressor — Sequential boosting",
    "LSTM deep learning model on historical data"
])
add_card(s6, 4.7, 1.8, 3.9, 5.0, "Feature Engineering", [
    "MalayalamCalendarFeaturizer — Kerala-specific feature engine",
    "Converts dates to arrays with festival proximity (Onam, Vishu)",
    "Real-time sales prediction from mid-day actuals"
])
add_card(s6, 8.9, 1.8, 3.9, 5.0, "Dormant Reactivation", [
    "SQL pre-aggregation engines for dormant customer bucketing",
    "Cohort-based analysis across 2020-2025 cohorts",
    "UI Visualizers: Probability Gauges, Semi-Donut Charts"
])

# --- SLIDE 7: Portals (She Start & FoneFlix) ---
s7 = prs.slides.add_slide(BLANK)
slide_template(s7, "5. Evaluation & Event Portals", "SHE START & BIGG BOSS/FoneFlix")
add_card(s7, 0.5, 1.8, 6.0, 5.0, "SHE START Dashboard", [
    "Live Google Sheets Sync: gspread library for bidirectional sync",
    "6-Panelist Scoring Algorithm: Auto-drop highest/lowest scores",
    "Interactive Dashboard: Inline cell editing with silent auto-save",
    "Auto badging: Strong Selection, Waitlist"
])
add_card(s7, 6.8, 1.8, 6.0, 5.0, "BIGG BOSS / FoneFlix Registration", [
    "Video Upload Architecture: Custom OAuth 2.0 Google Drive integration",
    "Dynamic UI/UX: OTT reality-show aesthetics, neon-glow typography",
    "Project Pivot: Successfully transitioned codebase from Bigg Boss to FoneFlix",
    "Refactored form inputs, consent clauses, and branding"
])

# --- SLIDE 8: Summary Table ---
s8 = prs.slides.add_slide(BLANK)
slide_template(s8, "All Projects Delivered", "Q2 2026 Portfolio Overview")
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
add_shape(s8, MSO_SHAPE.RECTANGLE, 0.5, 1.8, 12.33, 0.5, ACCENT_BLUE)
add_textbox(s8, 0.6, 1.85, 3.0, 0.4, "Project", size=14, color=WHITE, bold=True)
add_textbox(s8, 4.0, 1.85, 5.0, 0.4, "Description", size=14, color=WHITE, bold=True)
add_textbox(s8, 9.5, 1.85, 2.0, 0.4, "Platform", size=14, color=WHITE, bold=True)

for idx, (p_name, p_desc, p_plat) in enumerate(projects):
    ty = 2.4 + idx * 0.55
    bg_color = WHITE if idx % 2 == 0 else OFF_WHITE
    add_shape(s8, MSO_SHAPE.RECTANGLE, 0.5, ty, 12.33, 0.5, bg_color, LIGHT_GRAY)
    add_textbox(s8, 0.6, ty + 0.1, 3.3, 0.4, p_name, size=13, color=CHARCOAL, bold=True)
    add_textbox(s8, 4.0, ty + 0.1, 5.3, 0.4, p_desc, size=13, color=CHARCOAL)
    add_textbox(s8, 9.5, ty + 0.1, 2.0, 0.4, p_plat, size=13, color=CHARCOAL)

# --- SLIDE 9: THANK YOU ---
s9 = prs.slides.add_slide(BLANK)
add_shape(s9, MSO_SHAPE.RECTANGLE, 0, 0, 13.333, 7.5, CORP_BLUE)
add_shape(s9, MSO_SHAPE.OVAL, 9.0, -2.0, 6.0, 6.0, ACCENT_BLUE)
add_shape(s9, MSO_SHAPE.OVAL, -2.0, 4.0, 5.0, 5.0, OFF_WHITE)

add_textbox(s9, 0, 3.0, 13.333, 1.0, "THANK YOU", size=54, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
add_shape(s9, MSO_SHAPE.RECTANGLE, 6.166, 4.2, 1.0, 0.05, ACCENT_BLUE)
add_textbox(s9, 0, 4.5, 13.333, 0.5, "Jasil N Balussery | Technology / Software Development", size=18, color=LIGHT_GRAY, align=PP_ALIGN.CENTER)
add_textbox(s9, 0, 5.0, 13.333, 0.5, "April | May | June 2026", size=16, color=LIGHT_GRAY, align=PP_ALIGN.CENTER)

OUTPUT = r"c:\Users\jasil_myg\Desktop\OSG-myG-PORTAL-mainnnnn - Copy\Q2_2026_Work_Summary_Elegant_Corporate.pptx"
prs.save(OUTPUT)
print(f"Saved: {OUTPUT}")
