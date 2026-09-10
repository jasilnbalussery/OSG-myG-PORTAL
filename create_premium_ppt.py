"""
Q2 2026 Professional Work Summary — Premium Tech / Dark Mode Theme
"""

import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

# ── Theme Colors ──────────────────────────────────────────────
BG_DARK    = RGBColor(11, 15, 25)    # Very dark slate #0B0F19
CARD_DARK  = RGBColor(26, 35, 51)    # Lighter card #1A2333
NEON_CYAN  = RGBColor(0, 240, 255)   # #00F0FF
GOLD       = RGBColor(250, 204, 21)  # #FACC15
WHITE      = RGBColor(255, 255, 255)
LIGHT_GRAY = RGBColor(160, 170, 180) # Subtext
MED_GRAY   = RGBColor(70, 80, 100)   # Borders/Dividers

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

def add_textbox(slide, left, top, width, height, text, size=18, color=WHITE, bold=False, align=PP_ALIGN.LEFT):
    txBox = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(size)
    p.font.color.rgb = color
    p.font.bold = bold
    p.font.name = "Segoe UI" # More modern than Calibri
    p.alignment = align
    return txBox

def add_bullet_list(slide, items, left, top, width, height, size=12, color=LIGHT_GRAY):
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
        p.space_after = Pt(6)
        p.space_before = Pt(2)

def slide_template(slide, title, subtitle):
    # Background
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, 13.333, 7.5, BG_DARK)
    
    # Header Graphics
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0.5, 0.4, 0.1, 0.6, NEON_CYAN)
    
    # Title & Subtitle
    add_textbox(slide, 0.8, 0.35, 12, 0.6, title, size=28, color=WHITE, bold=True)
    if subtitle:
        add_textbox(slide, 0.8, 0.85, 12, 0.4, subtitle, size=14, color=NEON_CYAN)
        
    # Footer
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0.5, 7.2, 12.333, 0.02, MED_GRAY)
    add_textbox(slide, 0.5, 7.22, 12, 0.3, "myG TECHNOLOGY | ENGINEERING & AI | Q2 2026", size=9, color=LIGHT_GRAY, align=PP_ALIGN.LEFT)

# --- SLIDE 1: TITLE ---
s1 = prs.slides.add_slide(BLANK)
add_shape(s1, MSO_SHAPE.RECTANGLE, 0, 0, 13.333, 7.5, BG_DARK)
# Grid accent
add_shape(s1, MSO_SHAPE.RECTANGLE, 12.0, 0, 1.333, 7.5, CARD_DARK)
add_shape(s1, MSO_SHAPE.RECTANGLE, 12.0, 1.5, 0.05, 4.5, NEON_CYAN)
add_shape(s1, MSO_SHAPE.RECTANGLE, 12.0, 1.5, 1.333, 0.05, NEON_CYAN)

add_textbox(s1, 1.0, 2.0, 10.0, 0.5, "QUARTERLY BUSINESS REVIEW", size=14, color=GOLD, bold=True)
add_textbox(s1, 1.0, 2.5, 10.0, 1.5, "Q2 2026 WORK\nSUMMARY", size=60, color=WHITE, bold=True)
add_shape(s1, MSO_SHAPE.RECTANGLE, 1.0, 4.8, 3.5, 0.05, NEON_CYAN)
add_textbox(s1, 1.0, 5.2, 10.0, 0.8, "Jasil N Balussery", size=24, color=WHITE, bold=True)
add_textbox(s1, 1.0, 5.7, 10.0, 0.8, "Technology / Software Development | myG", size=16, color=LIGHT_GRAY)

# --- SLIDE 2: EXEC SUMMARY ---
s2 = prs.slides.add_slide(BLANK)
slide_template(s2, "EXECUTIVE SUMMARY", "HIGH-LEVEL KPIS & DELIVERY METRICS")

add_textbox(s2, 0.5, 1.6, 12.3, 1.0, "During Q2 2026, extensive full-stack development, database engineering, AI/ML integration, and multiple greenfield product launches were delivered across 8 distinct projects. Work spanned Python (Flask/Django), PostgreSQL, JavaScript, Machine Learning, LLM integration, Google APIs, and cloud deployment.", size=16, color=LIGHT_GRAY)

metrics = [("8", "Projects"), ("80+", "Dev Sessions"), ("12.6M+", "DB Rows"), ("26", "Mat Views"), ("< 1s", "AI Response"), ("4", "ML Models")]
for i, (val, lbl) in enumerate(metrics):
    row = i // 3
    col = i % 3
    lx = 0.5 + col * 4.2
    ty = 3.0 + row * 1.8
    add_shape(s2, MSO_SHAPE.ROUNDED_RECTANGLE, lx, ty, 3.8, 1.5, CARD_DARK, MED_GRAY)
    add_shape(s2, MSO_SHAPE.RECTANGLE, lx + 0.2, ty + 0.3, 0.05, 0.9, NEON_CYAN)
    add_textbox(s2, lx + 0.4, ty + 0.35, 3.0, 0.6, val, size=36, color=WHITE, bold=True)
    add_textbox(s2, lx + 0.4, ty + 0.9, 3.0, 0.5, lbl, size=14, color=GOLD, bold=True)

# --- HELPER FOR CONTENT SLIDES ---
def add_card(slide, left, top, width, height, title, points):
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height, CARD_DARK, MED_GRAY)
    add_shape(slide, MSO_SHAPE.RECTANGLE, left + 0.3, top + 0.45, 0.4, 0.05, NEON_CYAN) 
    add_textbox(slide, left + 0.3, top + 0.1, width - 0.4, 0.4, title, size=18, color=WHITE, bold=True)
    add_bullet_list(slide, points, left + 0.3, top + 0.7, width - 0.6, height - 0.8, size=13)

# --- SLIDE 3: OSG Portal ---
s3 = prs.slides.add_slide(BLANK)
slide_template(s3, "01. OSG-myG-PORTAL", "CLAIMS MANAGEMENT & WHATSAPP AUTOMATION")
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
slide_template(s4, "02. ENTERPRISE AI AGENT", "NATURAL LANGUAGE BUSINESS INTELLIGENCE")
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
slide_template(s5, "03. DATABASE ADMINISTRATION", "12.6M+ ROWS · DIGITALOCEAN POSTGRESQL")
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
slide_template(s6, "04. MACHINE LEARNING", "LIVE FORECASTING ENGINES & DATA PROCESSING")
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
slide_template(s7, "05. EVALUATION & EVENT PORTALS", "SHE START & BIGG BOSS/FONEFLIX")
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
slide_template(s8, "06. ALL PROJECTS DELIVERED", "Q2 2026 PORTFOLIO OVERVIEW")
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
add_shape(s8, MSO_SHAPE.RECTANGLE, 0.5, 1.8, 12.33, 0.4, NEON_CYAN)
add_textbox(s8, 0.6, 1.8, 3.0, 0.4, "PROJECT", size=12, color=BG_DARK, bold=True)
add_textbox(s8, 4.0, 1.8, 5.0, 0.4, "DESCRIPTION", size=12, color=BG_DARK, bold=True)
add_textbox(s8, 9.5, 1.8, 2.0, 0.4, "PLATFORM", size=12, color=BG_DARK, bold=True)

for idx, (p_name, p_desc, p_plat) in enumerate(projects):
    ty = 2.3 + idx * 0.55
    bg_color = CARD_DARK if idx % 2 == 0 else BG_DARK
    add_shape(s8, MSO_SHAPE.RECTANGLE, 0.5, ty, 12.33, 0.5, bg_color, MED_GRAY)
    add_textbox(s8, 0.6, ty + 0.1, 3.3, 0.4, p_name, size=13, color=WHITE, bold=True)
    add_textbox(s8, 4.0, ty + 0.1, 5.3, 0.4, p_desc, size=13, color=LIGHT_GRAY)
    add_textbox(s8, 9.5, ty + 0.1, 2.0, 0.4, p_plat, size=13, color=GOLD, bold=True)

# --- SLIDE 9: THANK YOU ---
s9 = prs.slides.add_slide(BLANK)
add_shape(s9, MSO_SHAPE.RECTANGLE, 0, 0, 13.333, 7.5, BG_DARK)
# Grid accent
add_shape(s9, MSO_SHAPE.RECTANGLE, 0, 0, 1.333, 7.5, CARD_DARK)
add_shape(s9, MSO_SHAPE.RECTANGLE, 1.283, 1.5, 0.05, 4.5, NEON_CYAN)
add_shape(s9, MSO_SHAPE.RECTANGLE, 0, 1.5, 1.333, 0.05, NEON_CYAN)

add_textbox(s9, 0, 3.0, 13.333, 1.0, "THANK YOU", size=54, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
add_shape(s9, MSO_SHAPE.RECTANGLE, 6.166, 4.2, 1.0, 0.05, NEON_CYAN)
add_textbox(s9, 0, 4.5, 13.333, 0.5, "Jasil N Balussery | Technology / Software Development", size=18, color=LIGHT_GRAY, align=PP_ALIGN.CENTER)
add_textbox(s9, 0, 5.0, 13.333, 0.5, "April | May | June 2026", size=14, color=GOLD, bold=True, align=PP_ALIGN.CENTER)

OUTPUT = r"c:\Users\jasil_myg\Desktop\OSG-myG-PORTAL-mainnnnn - Copy\Q2_2026_Work_Summary_Premium_Dark.pptx"
prs.save(OUTPUT)
print(f"Saved: {OUTPUT}")
