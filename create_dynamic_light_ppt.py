"""
Q2 2026 Professional Work Summary — Dynamic Light Theme (Varied Layouts)
"""

import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

# ── Colors ──────────────────────────────────────────────
BG_LIGHT   = RGBColor(248, 250, 252) # Slate 50
WHITE      = RGBColor(255, 255, 255)
TECH_BLUE  = RGBColor(0, 102, 255)   # #0066FF
DARK_BLUE  = RGBColor(15, 23, 42)    # #0F172A
AMBER      = RGBColor(245, 158, 11)  # #F59E0B
TEXT_MAIN  = RGBColor(30, 41, 59)    # Slate 800
TEXT_SUB   = RGBColor(100, 116, 139) # Slate 500
BORDER_GRY = RGBColor(226, 232, 240) # Slate 200
LIGHT_BLUE = RGBColor(219, 234, 254) # Blue 100

prs = Presentation()
prs.slide_width  = Inches(13.333)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]

def add_shape(slide, mso_shape, left, top, width, height, fill_color, line_color=None, line_width=1):
    shape = slide.shapes.add_shape(mso_shape, Inches(left), Inches(top), Inches(width), Inches(height))
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    if line_color:
        shape.line.color.rgb = line_color
        shape.line.width = Pt(line_width)
    else:
        shape.line.fill.background()
    shape.shadow.inherit = False
    return shape

def add_textbox(slide, left, top, width, height, text, size=18, color=TEXT_MAIN, bold=False, align=PP_ALIGN.LEFT, font_name="Segoe UI"):
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

def add_bullet_list(slide, items, left, top, width, height, size=12, color=TEXT_SUB, space_after=6, space_before=2):
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
        p.font.name = "Segoe UI"
        p.space_after = Pt(space_after)
        p.space_before = Pt(space_before)

# --- SLIDE 1: TITLE (Dynamic Split) ---
s1 = prs.slides.add_slide(BLANK)
add_shape(s1, MSO_SHAPE.RECTANGLE, 0, 0, 13.333, 7.5, BG_LIGHT)
# Angled background
add_shape(s1, MSO_SHAPE.RIGHT_TRIANGLE, 0, 0, 13.333, 7.5, TECH_BLUE)
add_shape(s1, MSO_SHAPE.RIGHT_TRIANGLE, 0, 0, 13.0, 7.2, DARK_BLUE)
s1.shapes[-1].rotation = 180 # Flip to put it on left/top
s1.shapes[-2].rotation = 180 

add_textbox(s1, 1.0, 1.5, 10.0, 0.5, "QUARTERLY BUSINESS REVIEW", size=16, color=AMBER, bold=True)
add_textbox(s1, 1.0, 2.0, 10.0, 1.5, "Q2 2026 WORK\nSUMMARY", size=65, color=DARK_BLUE, bold=True)
add_shape(s1, MSO_SHAPE.RECTANGLE, 1.0, 4.2, 2.0, 0.08, TECH_BLUE)
add_textbox(s1, 1.0, 4.6, 10.0, 0.8, "Jasil N Balussery", size=26, color=TEXT_MAIN, bold=True)
add_textbox(s1, 1.0, 5.2, 10.0, 0.8, "Technology / Software Development | myG", size=16, color=TEXT_SUB)

# --- SLIDE 2: EXEC SUMMARY (Top Header + Floating KPIs) ---
s2 = prs.slides.add_slide(BLANK)
add_shape(s2, MSO_SHAPE.RECTANGLE, 0, 0, 13.333, 7.5, BG_LIGHT)
# Huge top banner
add_shape(s2, MSO_SHAPE.RECTANGLE, 0, 0, 13.333, 3.5, DARK_BLUE)
add_textbox(s2, 0.5, 0.5, 12, 0.6, "EXECUTIVE SUMMARY", size=32, color=WHITE, bold=True)
add_shape(s2, MSO_SHAPE.RECTANGLE, 0.5, 1.2, 1.0, 0.05, TECH_BLUE)
add_textbox(s2, 0.5, 1.5, 12, 1.5, "During Q2 2026, extensive full-stack development, database engineering, AI/ML integration, and multiple greenfield product launches were delivered across 8 distinct projects. Work spanned Python (Flask/Django), PostgreSQL, JavaScript, Machine Learning, LLM integration, Google APIs, and cloud deployment.", size=16, color=LIGHT_BLUE)

# KPIs overlapping the banner
metrics = [("8", "Projects"), ("80+", "Dev Sessions"), ("12.6M+", "DB Rows"), ("26", "Mat Views"), ("10x", "Faster Reports"), ("4", "ML Models")]
for i, (val, lbl) in enumerate(metrics):
    row = i // 3
    col = i % 3
    lx = 0.8 + col * 4.0
    ty = 2.8 + row * 2.0
    add_shape(s2, MSO_SHAPE.ROUNDED_RECTANGLE, lx, ty, 3.6, 1.6, WHITE, BORDER_GRY)
    add_shape(s2, MSO_SHAPE.OVAL, lx - 0.2, ty + 0.3, 1.0, 1.0, TECH_BLUE)
    add_textbox(s2, lx - 0.2, ty + 0.45, 1.0, 1.0, f"0{i+1}", size=18, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    add_textbox(s2, lx + 1.0, ty + 0.3, 2.5, 0.6, val, size=32, color=DARK_BLUE, bold=True)
    add_textbox(s2, lx + 1.0, ty + 0.9, 2.5, 0.5, lbl, size=14, color=TEXT_SUB)

# --- SLIDE 3: OSG Portal (Vertical Timeline Layout) ---
s3 = prs.slides.add_slide(BLANK)
add_shape(s3, MSO_SHAPE.RECTANGLE, 0, 0, 13.333, 7.5, BG_LIGHT)
add_textbox(s3, 0.5, 0.3, 12, 0.6, "01. OSG-myG-PORTAL", size=28, color=DARK_BLUE, bold=True)
add_textbox(s3, 0.5, 0.8, 12, 0.4, "CLAIMS MANAGEMENT & WHATSAPP AUTOMATION", size=14, color=TECH_BLUE, bold=True)

# Vertical timeline line
add_shape(s3, MSO_SHAPE.RECTANGLE, 1.5, 1.8, 0.05, 5.0, BORDER_GRY)

# Node 1
add_shape(s3, MSO_SHAPE.OVAL, 1.375, 2.2, 0.3, 0.3, TECH_BLUE)
add_shape(s3, MSO_SHAPE.ROUNDED_RECTANGLE, 2.0, 1.8, 10.0, 2.0, WHITE, BORDER_GRY)
add_textbox(s3, 2.2, 2.0, 9.5, 0.4, "WhatsApp API Integration", size=18, color=DARK_BLUE, bold=True)
add_bullet_list(s3, [
    "Engineered automated messaging pipelines for portal events",
    "Configured 4 WhatsApp templates: Registered, Replacement, Repair Completed, Spare Parts",
    "Debugged payload routing for reliable outbound notifications",
    "Built trigger buttons in portal for WhatsApp message dispatch"
], 2.2, 2.4, 9.5, 1.2, size=13)

# Node 2
add_shape(s3, MSO_SHAPE.OVAL, 1.375, 4.6, 0.3, 0.3, AMBER)
add_shape(s3, MSO_SHAPE.ROUNDED_RECTANGLE, 2.0, 4.2, 10.0, 2.2, WHITE, BORDER_GRY)
add_textbox(s3, 2.2, 4.4, 9.5, 0.4, "Claims Processing & Data Forensics", size=18, color=DARK_BLUE, bold=True)
add_bullet_list(s3, [
    "Developed investigative scripts to trace and resolve claim anomalies",
    "Built repair utilities to retroactively fix corrupted claim states",
    "SR Number auto-generation on status change (Submit → Registered)",
    "Google Sheets bi-directional sync for claim tracking",
    "Massive data parsers for legacy Excel imports (17MB+ files)"
], 2.2, 4.8, 9.5, 1.4, size=13)


# --- SLIDE 4: AI Agent (Layered Architecture Layout) ---
s4 = prs.slides.add_slide(BLANK)
add_shape(s4, MSO_SHAPE.RECTANGLE, 0, 0, 13.333, 7.5, WHITE)
add_textbox(s4, 0.5, 0.3, 12, 0.6, "02. ENTERPRISE AI AGENT", size=28, color=DARK_BLUE, bold=True)
add_textbox(s4, 0.5, 0.8, 12, 0.4, "NATURAL LANGUAGE BUSINESS INTELLIGENCE", size=14, color=TECH_BLUE, bold=True)

# Layer 1
add_shape(s4, MSO_SHAPE.RECTANGLE, 1.0, 1.8, 11.333, 2.2, BG_LIGHT, BORDER_GRY)
add_shape(s4, MSO_SHAPE.RECTANGLE, 1.0, 1.8, 0.2, 2.2, TECH_BLUE)
add_textbox(s4, 1.4, 2.0, 10.5, 0.4, "Multi-Layered Architecture", size=18, color=DARK_BLUE, bold=True)
add_bullet_list(s4, [
    "SQL Agent: Translates natural language into complex PostgreSQL queries",
    "AI Analyst: Interprets raw DB output into readable business insights",
    "FastPath Engine: Common queries answered in < 0.2 seconds",
    "16-phase implementation covering Foundation to Export System",
    "Schema context injection prevents AI hallucinations"
], 1.4, 2.4, 10.5, 1.5, size=14, space_after=4)

# Layer 2
add_shape(s4, MSO_SHAPE.RECTANGLE, 1.0, 4.4, 11.333, 2.0, BG_LIGHT, BORDER_GRY)
add_shape(s4, MSO_SHAPE.RECTANGLE, 1.0, 4.4, 0.2, 2.0, AMBER)
add_textbox(s4, 1.4, 4.6, 10.5, 0.4, "LLM Integration & Optimization", size=18, color=DARK_BLUE, bold=True)
add_bullet_list(s4, [
    "NVIDIA NIM API (Nemotron Ultra) for SQL generation",
    "OpenRouter API as fallback LLM provider",
    "Response time reduced from 5-10 minutes to < 1 second",
    "Aggressive prompt engineering & query scope reduction",
    "Validated with 20+ test business questions"
], 1.4, 5.0, 10.5, 1.2, size=14, space_after=4)

# --- SLIDE 5: DB Admin (3 Pillars Layout) ---
s5 = prs.slides.add_slide(BLANK)
add_shape(s5, MSO_SHAPE.RECTANGLE, 0, 0, 13.333, 7.5, DARK_BLUE)
add_textbox(s5, 0.5, 0.3, 12, 0.6, "03. DATABASE ADMINISTRATION", size=28, color=WHITE, bold=True)
add_textbox(s5, 0.5, 0.8, 12, 0.4, "12.6M+ ROWS · DIGITALOCEAN POSTGRESQL", size=14, color=TECH_BLUE, bold=True)

# Pillar 1
add_shape(s5, MSO_SHAPE.ROUNDED_RECTANGLE, 0.5, 1.6, 3.8, 5.2, WHITE)
add_textbox(s5, 0.5, 1.7, 3.8, 1.0, "01", size=60, color=BG_LIGHT, bold=True, align=PP_ALIGN.CENTER)
add_textbox(s5, 0.7, 2.6, 3.4, 0.5, "Materialized View Overhaul", size=16, color=DARK_BLUE, bold=True)
add_bullet_list(s5, [
    "Diagnosed critical cross-join bug in mv_yearly_cohort",
    "Built robust refresh system for 26 materialized views",
    "Concurrent regeneration without database locks"
], 0.7, 3.2, 3.4, 3.0, size=13)

# Pillar 2
add_shape(s5, MSO_SHAPE.ROUNDED_RECTANGLE, 4.75, 1.6, 3.8, 5.2, WHITE)
add_textbox(s5, 4.75, 1.7, 3.8, 1.0, "02", size=60, color=BG_LIGHT, bold=True, align=PP_ALIGN.CENTER)
add_textbox(s5, 4.95, 2.6, 3.4, 0.5, "Caching Infrastructure", size=16, color=DARK_BLUE, bold=True)
add_bullet_list(s5, [
    "Integrated Redis for production caching",
    "LocMemCache for local development",
    "12.6M+ row calculations reduced to sub-second responses"
], 4.95, 3.2, 3.4, 3.0, size=13)

# Pillar 3
add_shape(s5, MSO_SHAPE.ROUNDED_RECTANGLE, 9.0, 1.6, 3.8, 5.2, WHITE)
add_textbox(s5, 9.0, 1.7, 3.8, 1.0, "03", size=60, color=BG_LIGHT, bold=True, align=PP_ALIGN.CENTER)
add_textbox(s5, 9.2, 2.6, 3.4, 0.5, "Data Operations", size=16, color=DARK_BLUE, bold=True)
add_bullet_list(s5, [
    "Custom scripts bypass web-server timeouts for large uploads",
    "Automated scrubbing of anomalous data",
    "Data integrity verification: 5M+ total records"
], 9.2, 3.2, 3.4, 3.0, size=13)


# --- SLIDE 6: Machine Learning (Dashboard Layout) ---
s6 = prs.slides.add_slide(BLANK)
add_shape(s6, MSO_SHAPE.RECTANGLE, 0, 0, 13.333, 7.5, BG_LIGHT)
add_textbox(s6, 0.5, 0.3, 12, 0.6, "04. MACHINE LEARNING", size=28, color=DARK_BLUE, bold=True)
add_textbox(s6, 0.5, 0.8, 12, 0.4, "LIVE FORECASTING ENGINES & DATA PROCESSING", size=14, color=TECH_BLUE, bold=True)

# Main Large Card (Left)
add_shape(s6, MSO_SHAPE.ROUNDED_RECTANGLE, 0.5, 1.6, 6.0, 5.3, TECH_BLUE)
add_shape(s6, MSO_SHAPE.OVAL, -1.0, 4.0, 4.0, 4.0, WHITE) # decorative
s6.shapes[-1].fill.transparency = 0.9
add_textbox(s6, 1.0, 2.5, 5.0, 1.0, "Model Deployment", size=24, color=WHITE, bold=True)
add_bullet_list(s6, [
    "Random Forest — Ensemble decision tree model",
    "MLPRegressor — Neural Network proxy for patterns",
    "GradientBoostingRegressor — Sequential boosting",
    "LSTM deep learning model on historical data"
], 1.0, 3.2, 5.0, 3.0, size=15, color=LIGHT_BLUE, space_after=10)

# Top Right Card
add_shape(s6, MSO_SHAPE.ROUNDED_RECTANGLE, 6.8, 1.6, 6.0, 2.5, WHITE, BORDER_GRY)
add_textbox(s6, 7.1, 1.9, 5.4, 0.4, "Feature Engineering", size=18, color=DARK_BLUE, bold=True)
add_bullet_list(s6, [
    "MalayalamCalendarFeaturizer — Kerala-specific feature engine",
    "Converts dates to arrays with festival proximity (Onam, Vishu)",
    "Real-time sales prediction from mid-day actuals"
], 7.1, 2.4, 5.4, 1.5, size=13)

# Bottom Right Card
add_shape(s6, MSO_SHAPE.ROUNDED_RECTANGLE, 6.8, 4.4, 6.0, 2.5, WHITE, BORDER_GRY)
add_textbox(s6, 7.1, 4.7, 5.4, 0.4, "Dormant Reactivation", size=18, color=DARK_BLUE, bold=True)
add_bullet_list(s6, [
    "SQL pre-aggregation engines for dormant customer bucketing",
    "Cohort-based analysis across 2020-2025 cohorts",
    "UI Visualizers: Probability Gauges, Semi-Donut Charts"
], 7.1, 5.2, 5.4, 1.5, size=13)

# --- SLIDE 7: Portals (Horizontal Split) ---
s7 = prs.slides.add_slide(BLANK)
# Top Half Dark Blue
add_shape(s7, MSO_SHAPE.RECTANGLE, 0, 0, 13.333, 3.75, DARK_BLUE)
# Bottom Half White
add_shape(s7, MSO_SHAPE.RECTANGLE, 0, 3.75, 13.333, 3.75, WHITE)

add_textbox(s7, 0.5, 0.3, 12, 0.6, "05. EVALUATION & EVENT PORTALS", size=28, color=WHITE, bold=True)
add_textbox(s7, 0.5, 0.8, 12, 0.4, "SHE START & BIGG BOSS/FONEFLIX", size=14, color=TECH_BLUE, bold=True)

# Top Content
add_textbox(s7, 0.5, 1.5, 12, 0.4, "SHE START Dashboard", size=20, color=AMBER, bold=True)
add_bullet_list(s7, [
    "Live Google Sheets Sync: gspread library for bidirectional sync (25-second refresh)",
    "6-Panelist Scoring Algorithm: Auto-drop highest/lowest scores to prevent bias",
    "Interactive Dashboard: Inline cell editing with silent auto-save",
    "Auto badging: Strong Selection, Waitlist"
], 0.5, 2.0, 12, 1.5, size=14, color=LIGHT_BLUE)

# Bottom Content
add_textbox(s7, 0.5, 4.2, 12, 0.4, "BIGG BOSS / FoneFlix Registration", size=20, color=TECH_BLUE, bold=True)
add_bullet_list(s7, [
    "Video Upload Architecture: Custom OAuth 2.0 Google Drive integration",
    "Dynamic UI/UX: OTT reality-show aesthetics, neon-glow typography",
    "Project Pivot: Successfully transitioned codebase from Bigg Boss to FoneFlix",
    "Refactored form inputs, consent clauses, and branding"
], 0.5, 4.7, 12, 1.5, size=14, color=TEXT_MAIN)


# --- SLIDE 8: Summary Table (Modern Grid) ---
s8 = prs.slides.add_slide(BLANK)
add_shape(s8, MSO_SHAPE.RECTANGLE, 0, 0, 13.333, 7.5, BG_LIGHT)
add_textbox(s8, 0.5, 0.3, 12, 0.6, "06. ALL PROJECTS DELIVERED", size=28, color=DARK_BLUE, bold=True)
add_textbox(s8, 0.5, 0.8, 12, 0.4, "Q2 2026 PORTFOLIO OVERVIEW", size=14, color=TECH_BLUE, bold=True)

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

# Grid Header
add_shape(s8, MSO_SHAPE.ROUNDED_RECTANGLE, 0.5, 1.6, 12.33, 0.5, DARK_BLUE)
add_textbox(s8, 0.8, 1.65, 3.0, 0.4, "PROJECT", size=13, color=WHITE, bold=True)
add_textbox(s8, 4.5, 1.65, 5.0, 0.4, "DESCRIPTION", size=13, color=WHITE, bold=True)
add_textbox(s8, 10.5, 1.65, 2.0, 0.4, "PLATFORM", size=13, color=WHITE, bold=True)

for idx, (p_name, p_desc, p_plat) in enumerate(projects):
    ty = 2.3 + idx * 0.55
    add_shape(s8, MSO_SHAPE.RECTANGLE, 0.5, ty, 12.33, 0.5, WHITE, BORDER_GRY)
    # Left accent
    if idx % 2 == 0:
        add_shape(s8, MSO_SHAPE.RECTANGLE, 0.5, ty, 0.1, 0.5, TECH_BLUE)
    else:
        add_shape(s8, MSO_SHAPE.RECTANGLE, 0.5, ty, 0.1, 0.5, AMBER)
        
    add_textbox(s8, 0.8, ty + 0.1, 3.5, 0.4, p_name, size=13, color=DARK_BLUE, bold=True)
    add_textbox(s8, 4.5, ty + 0.1, 5.8, 0.4, p_desc, size=13, color=TEXT_SUB)
    
    # Platform badge
    add_shape(s8, MSO_SHAPE.ROUNDED_RECTANGLE, 10.4, ty + 0.08, 1.5, 0.34, LIGHT_BLUE)
    add_textbox(s8, 10.4, ty + 0.1, 1.5, 0.34, p_plat, size=11, color=TECH_BLUE, bold=True, align=PP_ALIGN.CENTER)

# --- SLIDE 9: THANK YOU (Centered Concentric) ---
s9 = prs.slides.add_slide(BLANK)
add_shape(s9, MSO_SHAPE.RECTANGLE, 0, 0, 13.333, 7.5, TECH_BLUE)
# Concentric circles
add_shape(s9, MSO_SHAPE.OVAL, 1.666, -2.5, 10.0, 10.0, WHITE)
add_shape(s9, MSO_SHAPE.OVAL, 3.166, -1.0, 7.0, 7.0, BG_LIGHT)

add_textbox(s9, 0, 2.8, 13.333, 1.0, "THANK YOU", size=60, color=DARK_BLUE, bold=True, align=PP_ALIGN.CENTER)
add_shape(s9, MSO_SHAPE.RECTANGLE, 6.166, 4.0, 1.0, 0.05, AMBER)
add_textbox(s9, 0, 4.3, 13.333, 0.5, "Jasil N Balussery | Technology / Software Development", size=18, color=TEXT_SUB, align=PP_ALIGN.CENTER)
add_textbox(s9, 0, 4.8, 13.333, 0.5, "April | May | June 2026", size=16, color=TECH_BLUE, bold=True, align=PP_ALIGN.CENTER)

OUTPUT = r"c:\Users\jasil_myg\Desktop\OSG-myG-PORTAL-mainnnnn - Copy\Q2_2026_Work_Summary_Dynamic_Layouts_V3.pptx"
prs.save(OUTPUT)
print(f"Saved: {OUTPUT}")
