"""
Q2 2026 Work Summary PPT — built on the 'Challenging Success' template.
Strategy: clone template slides, clear placeholder text, inject our content.
"""
import copy, sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

from lxml import etree
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

TEMPLATE = r"c:\Users\jasil_myg\Desktop\OSG-myG-PORTAL-mainnnnn - Copy\Challenging Success Business PowerPoint Templates.pptx"
OUTPUT   = r"c:\Users\jasil_myg\Desktop\OSG-myG-PORTAL-mainnnnn - Copy\Q2_2026_Work_Summary_TEMPLATE.pptx"

# ── Color palette (matching template's red/dark theme) ──────────
RED     = RGBColor(0xC0, 0x39, 0x2B)
DARK    = RGBColor(0x1A, 0x1A, 0x2E)
WHITE   = RGBColor(0xFF, 0xFF, 0xFF)
LGRAY   = RGBColor(0xCC, 0xCC, 0xCC)
GOLD    = RGBColor(0xD4, 0xA5, 0x37)

prs_tmpl = Presentation(TEMPLATE)
W = prs_tmpl.slide_width
H = prs_tmpl.slide_height

# ── Helper: clone a template slide into a NEW presentation ───────
def clone_slide(prs_dst, slide_src):
    """Append a deep-copy of slide_src into prs_dst and return it."""
    tmpl = prs_dst.slide_layouts[0]   # any layout – we'll override xml
    new_slide = prs_dst.slides.add_slide(tmpl)
    # Replace the xml element of the new slide with a clone of the source
    src_xml = copy.deepcopy(slide_src._element)
    new_slide._element.getparent().replace(new_slide._element, src_xml)
    # Re-register relationships from source slide into new slide
    for rel in slide_src.part.rels.values():
        new_slide.part.rels[rel.reltype] = rel
    return prs_dst.slides[-1]


# ── Helper: set text in ALL text frames on a slide ───────────────
def set_all_text(slide, texts: list, font_size=None, bold=None, color=None):
    """Replace text in text-frames sequentially from texts list."""
    boxes = [sh for sh in slide.shapes if sh.has_text_frame]
    for i, sh in enumerate(boxes):
        if i >= len(texts):
            break
        tf = sh.text_frame
        tf.clear()
        p = tf.paragraphs[0]
        run = p.add_run()
        run.text = texts[i]
        if font_size:
            run.font.size = Pt(font_size)
        if color:
            run.font.color.rgb = color
        if bold is not None:
            run.font.bold = bold


def add_textbox(slide, left_in, top_in, w_in, h_in, text,
                size=14, bold=False, color=WHITE, align=PP_ALIGN.LEFT,
                wrap=True):
    tb = slide.shapes.add_textbox(Inches(left_in), Inches(top_in),
                                  Inches(w_in), Inches(h_in))
    tf = tb.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    return tb


def add_para(tf, text, size=12, bold=False, color=WHITE,
             space_before=4, align=PP_ALIGN.LEFT):
    p = tf.add_paragraph()
    p.alignment = align
    p.space_before = Pt(space_before)
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    return p


def add_rect(slide, l, t, w, h, fill_rgb, line=False):
    from pptx.enum.shapes import MSO_SHAPE
    sh = slide.shapes.add_shape(1, Inches(l), Inches(t), Inches(w), Inches(h))
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill_rgb
    if not line:
        sh.line.fill.background()
    return sh


# ════════════════════════════════════════════════════════════════
# Build NEW presentation (fresh, but same slide size as template)
# ════════════════════════════════════════════════════════════════
prs = Presentation()
prs.slide_width  = W
prs.slide_height = H

BLANK = prs.slide_layouts[6]  # pure blank layout


# ══════════════════════════════════════════════════════════════
# SLIDE 1: Title — modeled on template slide 0 (Cover)
# ══════════════════════════════════════════════════════════════
# We copy slide 0 from template into our new prs
src = prs_tmpl.slides[0]
blank = prs.slides.add_slide(BLANK)

# Background rect (dark)
add_rect(blank, 0, 0, 13.33, 7.5, DARK)

# Red accent left bar
add_rect(blank, 0, 0, 0.25, 7.5, RED)

# Red accent bottom strip
add_rect(blank, 0.25, 6.9, 13.08, 0.6, RED)

# Company / period tag
add_textbox(blank, 0.6, 0.3, 6, 0.5, "myG | Technology & Software Development",
            size=13, bold=False, color=LGRAY)

# Main title
add_textbox(blank, 0.6, 1.2, 12, 1.2, "Q2 2026 Work Summary",
            size=52, bold=True, color=WHITE)

# Subtitle
add_textbox(blank, 0.6, 2.55, 10, 0.7, "April  |  May  |  June  2026",
            size=28, bold=False, color=RED)

# Horizontal rule
add_rect(blank, 0.6, 3.4, 8, 0.05, RED)

# Description
add_textbox(blank, 0.6, 3.6, 11, 1.0,
            "A comprehensive breakdown of engineering, database optimization,\n"
            "frontend development, and AI integration across all workspaces.",
            size=16, color=LGRAY)

# Name / date
add_textbox(blank, 0.6, 5.1, 10, 0.45,
            "Jasil N Balussery  |  01 April 2026 - 30 June 2026",
            size=13, color=LGRAY)

# "8 Projects" badge
add_rect(blank, 10.5, 1.5, 2.3, 2.3, RED)
add_textbox(blank, 10.5, 1.5, 2.3, 1.3, "8",
            size=72, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
add_textbox(blank, 10.5, 2.85, 2.3, 0.6, "PROJECTS\nDELIVERED",
            size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

# Bottom bar text
add_textbox(blank, 0.4, 6.92, 12, 0.5,
            "Engineering  |  AI/ML  |  Database  |  Cloud Deployment  |  Product Launch",
            size=13, bold=False, color=WHITE, align=PP_ALIGN.CENTER)


# ══════════════════════════════════════════════════════════════
# SLIDE 2: Executive Summary + KPI metrics
# ══════════════════════════════════════════════════════════════
s2 = prs.slides.add_slide(BLANK)
add_rect(s2, 0, 0, 13.33, 7.5, DARK)
add_rect(s2, 0, 0, 13.33, 1.1, RED)
add_rect(s2, 0, 6.8, 13.33, 0.7, RED)

add_textbox(s2, 0.4, 0.1, 12, 0.9, "EXECUTIVE SUMMARY & KEY METRICS",
            size=30, bold=True, color=WHITE, align=PP_ALIGN.LEFT)

# Summary paragraph
tb = s2.shapes.add_textbox(Inches(0.4), Inches(1.3), Inches(12.5), Inches(1.3))
tf = tb.text_frame; tf.word_wrap = True
p = tf.paragraphs[0]
run = p.add_run()
run.font.size = Pt(14); run.font.color.rgb = LGRAY
run.text = ("During Q2 2026, extensive full-stack development, database engineering, AI/ML integration, "
            "and multiple greenfield product launches were delivered across 8 distinct projects. "
            "Work spanned Python (Flask/Django), PostgreSQL, JavaScript, Machine Learning, LLM APIs, "
            "Google APIs, and cloud deployment (Render, Vercel).")

# KPI Cards - Row 1
CARDS = [
    ("8",      "Projects Delivered",    0.4),
    ("80+",    "Dev Sessions",          3.6),
    ("12.6M+", "DB Rows Managed",       6.8),
    ("26",     "Materialized Views",   10.0),
]
for val, label, lx in CARDS:
    add_rect(s2, lx, 2.8, 2.8, 1.5, RED)
    add_textbox(s2, lx, 2.8, 2.8, 0.9, val,
                size=38, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_textbox(s2, lx, 3.7, 2.8, 0.55, label,
                size=12, color=WHITE, align=PP_ALIGN.CENTER)

# KPI Cards - Row 2
CARDS2 = [
    ("<1 sec",  "AI Response Time",     0.4),
    ("6",       "APIs Integrated",      3.6),
    ("4",       "ML Models Deployed",   6.8),
    ("3",       "Cloud Platforms",     10.0),
]
for val, label, lx in CARDS2:
    add_rect(s2, lx, 4.5, 2.8, 1.5, RGBColor(0x2C, 0x2C, 0x3E))
    add_textbox(s2, lx, 4.5, 2.8, 0.9, val,
                size=34, bold=True, color=RED, align=PP_ALIGN.CENTER)
    add_textbox(s2, lx, 5.4, 2.8, 0.55, label,
                size=12, color=LGRAY, align=PP_ALIGN.CENTER)

add_textbox(s2, 0.4, 6.82, 12.5, 0.55,
            "Tech: Python | Flask | Django | PostgreSQL | Redis | Scikit-Learn | LSTM | NVIDIA NIM | n8n | Google APIs | Render | Vercel",
            size=11, color=WHITE, align=PP_ALIGN.CENTER)


# ══════════════════════════════════════════════════════════════
# Reusable: section slide builder (4-card layout)
# ══════════════════════════════════════════════════════════════
CARD_BG = RGBColor(0x22, 0x22, 0x38)

def section_4card(prs, title, subtitle, cards):
    """
    cards: list of (heading, bullets_list) - up to 4 items
    """
    s = prs.slides.add_slide(BLANK)
    add_rect(s, 0, 0, 13.33, 7.5, DARK)
    add_rect(s, 0, 0, 13.33, 1.1, RED)
    add_rect(s, 0, 6.85, 13.33, 0.65, RED)

    add_textbox(s, 0.4, 0.05, 10, 0.65, title,
                size=28, bold=True, color=WHITE)
    add_textbox(s, 0.4, 0.68, 12, 0.38, subtitle,
                size=14, color=LGRAY)

    # Up to 4 cards in 2x2 grid
    positions = [
        (0.3,  1.2),  # top-left
        (6.9,  1.2),  # top-right
        (0.3,  4.1),  # bottom-left
        (6.9,  4.1),  # bottom-right
    ]
    for (lx, ty), (heading, bullets) in zip(positions, cards):
        add_rect(s, lx, ty, 6.4, 2.65, CARD_BG)
        # Red top-accent bar
        add_rect(s, lx, ty, 6.4, 0.07, RED)
        # Heading
        add_textbox(s, lx+0.15, ty+0.12, 6.1, 0.45, heading,
                    size=14, bold=True, color=RED)
        # Bullets
        tb = s.shapes.add_textbox(Inches(lx+0.15), Inches(ty+0.6),
                                  Inches(6.1), Inches(1.9))
        tf = tb.text_frame; tf.word_wrap = True
        first = True
        for b in bullets:
            if first:
                p = tf.paragraphs[0]; first = False
            else:
                p = tf.add_paragraph()
            p.space_before = Pt(3)
            run = p.add_run()
            run.text = b
            run.font.size = Pt(11)
            run.font.color.rgb = LGRAY

    add_textbox(s, 0.4, 6.88, 12.5, 0.5,
                "Q2 2026  |  Jasil N Balussery  |  myG Technology",
                size=11, color=WHITE, align=PP_ALIGN.CENTER)
    return s


# ══════════════════════════════════════════════════════════════
# SLIDE 3: OSG-myG-PORTAL
# ══════════════════════════════════════════════════════════════
section_4card(prs,
    "1. OSG-myG-PORTAL",
    "Claims Management & WhatsApp Automation Engineering",
    [
        ("WhatsApp API Integration (Telinfy/GreenAds)", [
            "- Engineered automated messaging pipelines for portal events",
            "- Configured 4 message templates: Registered, Replacement, Repair Completed, Spare Parts",
            "- Debugged payload routing for reliable outbound notifications",
            "- Built trigger buttons in portal for WhatsApp dispatch",
        ]),
        ("Claims Processing & Data Forensics", [
            "- Investigative scripts to trace and resolve claim anomalies",
            "- Repair utilities to fix corrupted claim states retroactively",
            "- SR Number auto-generation on status change (Submit -> Registered)",
            "- Comprehensive claims status workflow management",
        ]),
        ("Large-Scale Data Ingestion (OSID)", [
            "- Massive data parsers for legacy Excel imports (17MB+ files)",
            "- Automated date format standardization pipelines",
            "- Updated OSID reference, Future Store List, RBM/BDM/Branch data",
            "- 15,493-byte validation logic layer",
        ]),
        ("Google Apps Script Bridge", [
            "- Seamless Google Sheets <-> PostgreSQL bi-directional sync",
            "- New claims auto-append to Sheet with all data columns",
            "- Remarks & status updates sync to Follow-up Notes in real-time",
            "- Multiple Apps Script Web App deployments for reliability",
        ]),
    ]
)


# ══════════════════════════════════════════════════════════════
# SLIDE 4: Enterprise AI Agent
# ══════════════════════════════════════════════════════════════
s4 = prs.slides.add_slide(BLANK)
add_rect(s4, 0, 0, 13.33, 7.5, DARK)
add_rect(s4, 0, 0, 13.33, 1.1, RED)
add_rect(s4, 0, 6.85, 13.33, 0.65, RED)

add_textbox(s4, 0.4, 0.05, 12, 0.65, "2. ENTERPRISE AI AGENT",
            size=28, bold=True, color=WHITE)
add_textbox(s4, 0.4, 0.68, 12, 0.38,
            "Natural Language Business Intelligence — Loyalty Portal",
            size=14, color=LGRAY)

# Pipeline flow
flow = ["User Question", "Schema Catalog", "SQL Agent", "PostgreSQL", "AI Analyst", "Business Insight"]
for i, label in enumerate(flow):
    lx = 0.4 + i * 2.15
    col = RED if i in (2, 4) else RGBColor(0x2C, 0x2C, 0x3E)
    add_rect(s4, lx, 1.25, 1.9, 0.55, col)
    add_textbox(s4, lx, 1.25, 1.9, 0.55, label,
                size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    if i < 5:
        add_textbox(s4, lx+1.95, 1.32, 0.2, 0.4, "->",
                    size=14, bold=True, color=RED, align=PP_ALIGN.CENTER)

# Left card
add_rect(s4, 0.3, 2.05, 6.2, 4.55, CARD_BG)
add_rect(s4, 0.3, 2.05, 6.2, 0.07, RED)
add_textbox(s4, 0.45, 2.12, 5.8, 0.45, "Multi-Layered Agent Architecture",
            size=14, bold=True, color=RED)
tb = s4.shapes.add_textbox(Inches(0.45), Inches(2.6), Inches(5.8), Inches(3.85))
tf = tb.text_frame; tf.word_wrap = True
bullets_ai = [
    "- SQL Agent: Natural language -> complex PostgreSQL queries",
    "- AI Analyst: Raw DB output -> readable business insights",
    "- FastPath Engine: Common queries answered in < 0.2 seconds",
    "- 16-phase implementation: Foundation to Export System",
    "- Schema context injection prevents hallucinations",
    "- Conversation memory for contextual follow-up questions",
    "- Validated against 20+ complex business test questions",
]
first = True
for b in bullets_ai:
    if first:
        p = tf.paragraphs[0]; first = False
    else:
        p = tf.add_paragraph()
    p.space_before = Pt(5)
    r = p.add_run(); r.text = b; r.font.size = Pt(12); r.font.color.rgb = LGRAY

# Right card
add_rect(s4, 6.85, 2.05, 6.15, 4.55, CARD_BG)
add_rect(s4, 6.85, 2.05, 6.15, 0.07, RED)
add_textbox(s4, 7.0, 2.12, 5.8, 0.45, "LLM Integration & Optimization",
            size=14, bold=True, color=RED)
tb2 = s4.shapes.add_textbox(Inches(7.0), Inches(2.6), Inches(5.8), Inches(3.85))
tf2 = tb2.text_frame; tf2.word_wrap = True
bullets_llm = [
    "- NVIDIA NIM API (Nemotron Ultra) for SQL generation",
    "- OpenRouter API as fallback LLM provider",
    "- Response time: 5-10 minutes reduced to < 1 second",
    "- Aggressive prompt engineering & query scope reduction",
    "- n8n workflow automation for embedded chat widget",
    "- Real-time DB query answers via n8n Chat Trigger",
    "- OpenAI GPT integration for additional accuracy",
]
first = True
for b in bullets_llm:
    if first:
        p = tf2.paragraphs[0]; first = False
    else:
        p = tf2.add_paragraph()
    p.space_before = Pt(5)
    r = p.add_run(); r.text = b; r.font.size = Pt(12); r.font.color.rgb = LGRAY

add_textbox(s4, 0.4, 6.88, 12.5, 0.5,
            "Q2 2026  |  Jasil N Balussery  |  myG Technology",
            size=11, color=WHITE, align=PP_ALIGN.CENTER)


# ══════════════════════════════════════════════════════════════
# SLIDE 5: Database Administration
# ══════════════════════════════════════════════════════════════
s5 = prs.slides.add_slide(BLANK)
add_rect(s5, 0, 0, 13.33, 7.5, DARK)
add_rect(s5, 0, 0, 13.33, 1.1, RED)
add_rect(s5, 0, 6.85, 13.33, 0.65, RED)

add_textbox(s5, 0.4, 0.05, 12, 0.65, "3. DATABASE ADMINISTRATION & OPTIMIZATION",
            size=26, bold=True, color=WHITE)
add_textbox(s5, 0.4, 0.68, 12, 0.38,
            "12.6M+ Rows  |  DigitalOcean PostgreSQL  |  26 Materialized Views",
            size=14, color=LGRAY)

# 3 vertical cards
card_data = [
    ("Materialized View Overhaul", [
        "- Critical cross-join bug found in mv_yearly_cohort",
        "- Bug was inflating LTV metrics by 1000%+",
        "- Rebuilt all 26 materialized views concurrently",
        "- No database locks during regeneration",
        "- Views include: mv_action_engine, mv_cohort_customer_years,",
        "  mv_monthly_summary, mv_rfm_summary, mv_dormant_reactivation...",
    ]),
    ("Caching Infrastructure", [
        "- Redis integrated for production API caching",
        "- LocMemCache for local development",
        "- 12.6M+ row queries cached to milliseconds",
        "- Fixed stale memory cache in local environments",
        "- Sub-second dashboard response times achieved",
        "- Heavy analytics endpoints cached efficiently",
    ]),
    ("DB Manager & Data Operations", [
        "- Custom scripts bypass web-server upload timeouts",
        "- DSR MAY & JUNE 2026 data imported in multiple parts",
        "- Auto-scrub: SMC/EI, HEAD OFFICE, UG SMART CHOICE",
        "- Data integrity verified: 50,33,297+ total records",
        "- Upload success/failure notifications added",
        "- Unique customer verification: 52,42,619 records",
    ]),
]
for i, (heading, bullets) in enumerate(card_data):
    lx = 0.3 + i * 4.35
    add_rect(s5, lx, 1.2, 4.1, 5.4, CARD_BG)
    add_rect(s5, lx, 1.2, 4.1, 0.07, RED)
    add_textbox(s5, lx+0.15, 1.27, 3.8, 0.5, heading,
                size=13, bold=True, color=RED)
    tb = s5.shapes.add_textbox(Inches(lx+0.15), Inches(1.85),
                               Inches(3.8), Inches(4.5))
    tf = tb.text_frame; tf.word_wrap = True
    first = True
    for b in bullets:
        if first:
            p = tf.paragraphs[0]; first = False
        else:
            p = tf.add_paragraph()
        p.space_before = Pt(5)
        r = p.add_run(); r.text = b; r.font.size = Pt(11); r.font.color.rgb = LGRAY

add_textbox(s5, 0.4, 6.88, 12.5, 0.5,
            "Q2 2026  |  Jasil N Balussery  |  myG Technology",
            size=11, color=WHITE, align=PP_ALIGN.CENTER)


# ══════════════════════════════════════════════════════════════
# SLIDE 6: Machine Learning
# ══════════════════════════════════════════════════════════════
section_4card(prs,
    "4. MACHINE LEARNING & PREDICTIVE FORECASTING",
    "Live Forecasting Engines Replacing Static Dashboard Placeholders",
    [
        ("Model Deployment (Scikit-Learn + LSTM)", [
            "- Random Forest for ensemble sales prediction",
            "- MLPRegressor (Neural Network) for pattern recognition",
            "- GradientBoostingRegressor for high-accuracy forecasting",
            "- LSTM deep learning model on 2020-2026 historical data",
        ]),
        ("Feature Engineering: MalayalamCalendarFeaturizer", [
            "- Converts dates into multi-dimensional ML feature arrays",
            "- Incorporates Kerala festival proximity (Onam, Vishu, etc.)",
            "- Seasonal patterns, year-over-year trend analysis",
            "- Real-time EOD sale prediction from mid-day actuals",
        ]),
        ("Dormant Customer Reactivation System", [
            "- SQL engines bucket customers: 2024 buyers, 2025/26 silent",
            "- Cohort-based analysis across 2020-2025 cohorts",
            "- Resurrection Rate % tracked per cohort per month",
            "- UI: Probability Gauges, Semi-Donuts, Dormancy Risk Meters",
        ]),
        ("Campaign Analysis & Forecasting UI", [
            "- Plotly-powered interactive dashboards",
            "- Deep Learning Forecasting with live prediction engine",
            "- Repeat purchase probability scoring per customer",
            "- AI insights with detailed analysis drill-down buttons",
        ]),
    ]
)


# ══════════════════════════════════════════════════════════════
# SLIDE 7: SHE START Dashboard
# ══════════════════════════════════════════════════════════════
s7 = prs.slides.add_slide(BLANK)
add_rect(s7, 0, 0, 13.33, 7.5, DARK)
add_rect(s7, 0, 0, 13.33, 1.1, RED)
add_rect(s7, 0, 6.85, 13.33, 0.65, RED)

add_textbox(s7, 0.4, 0.05, 12, 0.65,
            '5. SHE START — Applicant Evaluation Dashboard',
            size=26, bold=True, color=WHITE)
add_textbox(s7, 0.4, 0.68, 12, 0.38,
            '"She Start — Her Dreams Start Here" | Women Startup Program Evaluation',
            size=14, color=LGRAY)

# 3 cards top row
card3_data = [
    ("Live Google Sheets Sync", [
        "- gspread library for real-time bidirectional sync",
        "- 25-second auto-refresh interval",
        "- Data order matches Google Sheet exactly",
        "- Auto-detect new applicants",
        "- District & business name auto-populated",
        "- 28 candidates tracked and scored",
    ]),
    ("6-Panelist Scoring Algorithm", [
        "- 6 panelists each score every applicant",
        "- Auto-drop: highest + lowest scores removed",
        "- Average of remaining 4 scores (unbiased)",
        "- Interview Score: 40% | Growth: 15%",
        "- Support Need: 15% | Emotional: 10%",
        "- Sustainability: 10% | Utilization: 10%",
    ]),
    ("Interactive Dashboard & Access", [
        "- Inline cell editing for 5 score columns",
        "- Silent auto-save (no page reload needed)",
        "- Auto-badging: Strong Selection (85+), Waitlist",
        "- Tab key navigation fixed (-> direction)",
        "- Dedicated user: shestart / shestart123",
        "- Excel report download | Deployed on Render",
    ]),
]
for i, (heading, bullets) in enumerate(card3_data):
    lx = 0.3 + i * 4.35
    add_rect(s7, lx, 1.2, 4.1, 5.4, CARD_BG)
    add_rect(s7, lx, 1.2, 4.1, 0.07, RED)
    add_textbox(s7, lx+0.15, 1.27, 3.8, 0.5, heading,
                size=13, bold=True, color=RED)
    tb = s7.shapes.add_textbox(Inches(lx+0.15), Inches(1.85),
                               Inches(3.8), Inches(4.5))
    tf = tb.text_frame; tf.word_wrap = True
    first = True
    for b in bullets:
        if first:
            p = tf.paragraphs[0]; first = False
        else:
            p = tf.add_paragraph()
        p.space_before = Pt(5)
        r = p.add_run(); r.text = b; r.font.size = Pt(11); r.font.color.rgb = LGRAY

add_textbox(s7, 0.4, 6.88, 12.5, 0.5,
            "Q2 2026  |  Jasil N Balussery  |  myG Technology",
            size=11, color=WHITE, align=PP_ALIGN.CENTER)


# ══════════════════════════════════════════════════════════════
# SLIDE 8: Enterprise Retail Dashboard Enhancements
# ══════════════════════════════════════════════════════════════
section_4card(prs,
    "6. ENTERPRISE RETAIL DASHBOARD ENHANCEMENTS",
    "Performance, Advanced Analytics & New Feature Sections",
    [
        ("High-Speed Reporting", [
            "- openpyxl replaced with Rust-based calamine engine",
            "- 5-10x faster Excel parsing and file generation",
            "- Monthly Category & Brand Report (APR vs MAY 2026)",
            "- PDF + Excel dual export in Crores (Rs.)",
        ]),
        ("New Dashboard Sections Added", [
            "- Monthly Retention Analysis (baseline: pre-Dec 2025)",
            "- Campaign Analysis / Dormant Customer Resurrection",
            "- Redemption Analysis (loyalty points & redeemed sales)",
            "- Loyalty Point Cohort & Customer Segmentation Download",
        ]),
        ("DataTables & Advanced Metrics", [
            "- Static tables replaced with searchable DataTables grids",
            "- Dynamic Average Selling Price (ASP) mapping",
            "- Product-to-category cross-referencing in Pandas",
            "- Multi-branch selection with custom date range filter",
        ]),
        ("Analytics & Access Control", [
            "- Role-based download access (mygadmin only)",
            "- 5,242,619 unique customers verified and deduplicated",
            "- Branch-level new vs repeat customer reports",
            "- Real-time sales prediction from mid-day actuals",
        ]),
    ]
)


# ══════════════════════════════════════════════════════════════
# SLIDE 9: BIGG BOSS & FoneFlix
# ══════════════════════════════════════════════════════════════
section_4card(prs,
    "7. BIGG BOSS & FoneFLIX REGISTRATION PORTALS",
    "Flask-Based Registration Platform | Video Upload Architecture | Zero-Downtime Pivot",
    [
        ("Video Upload (Google Drive OAuth 2.0)", [
            "- Custom OAuth 2.0 Client ID for Drive video uploads",
            "- Users upload files up to 500MB directly to Drive",
            "- Backup: Render server storage if Drive fails",
            "- Upload progress bar with real-time status",
        ]),
        ("Dynamic OTT-Style UI/UX Design", [
            "- OTT reality-show aesthetic: 45/60 two-column layout",
            "- Neon-glow typography, glassmorphism forms",
            "- Floating 3D particle animations, welcome popup banner",
            "- Pixel-perfect responsive: desktop + mobile",
        ]),
        ("Security & Validation", [
            "- Server-side: mobile (10 digits), email, age 18+",
            "- One mobile number = one entry enforced",
            "- Google Apps Script URL protected from exposure",
            "- JSON backup file for data redundancy",
        ]),
        ("Project Pivot: Bigg Boss -> FoneFlix", [
            "- Full codebase rebrand with zero downtime",
            "- Navy/Neon Pink -> Dark Browns/Orange theme",
            "- Form text, consent clauses, submit button updated",
            "- New Drive folder & Sheets for FoneFlix contest",
        ]),
    ]
)


# ══════════════════════════════════════════════════════════════
# SLIDE 10: Additional Projects
# ══════════════════════════════════════════════════════════════
section_4card(prs,
    "ADDITIONAL PROJECTS DELIVERED",
    "HR E-Signature Portal | FIFA World Cup Contest | n8n AI Chat Agent",
    [
        ("HR E-Signature Portal (June)", [
            "- New Flask-based HR portal from scratch",
            "- Digital signature capture and storage",
            "- Google Sheets: employee data auto-save",
            "- Signature images saved to Google Drive",
            "- Deployed on Render: hr-hq19.onrender.com",
        ]),
        ("FIFA World Cup Contest Platform (June 29)", [
            "- Match prediction + leaderboard platform",
            "- Points system for exact score / winner / goal diff",
            "- Full tournament bracket: R32 -> Final",
            "- Google Sheets data integration",
            "- Deployed on Vercel | GitHub: jasilmyg/fifaworldcup",
        ]),
        ("n8n Chat Agent Integration (June 22)", [
            "- n8n workflow automation for chat widget",
            "- Chat Trigger node with streaming response",
            "- Direct PostgreSQL query execution from chat",
            "- SQL schema fed to AI for accurate answers",
            "- Embedded in Loyalty Dashboard portal",
        ]),
        ("Birthday Website (May 31)", [
            "- Ultra-premium surprise website",
            "- Black and gold luxury glassmorphism design",
            "- Cinematic animations, fireworks, confetti",
            "- Premium typography with emotional storytelling",
            "- Floating 3D particle animations",
        ]),
    ]
)


# ══════════════════════════════════════════════════════════════
# SLIDE 11: All Projects Table
# ══════════════════════════════════════════════════════════════
s11 = prs.slides.add_slide(BLANK)
add_rect(s11, 0, 0, 13.33, 7.5, DARK)
add_rect(s11, 0, 0, 13.33, 1.1, RED)
add_rect(s11, 0, 6.85, 13.33, 0.65, RED)

add_textbox(s11, 0.4, 0.05, 12, 0.65, "ALL PROJECTS DELIVERED — Q2 2026",
            size=28, bold=True, color=WHITE)
add_textbox(s11, 0.4, 0.68, 12, 0.38, "8 Production Applications | April to June 2026",
            size=14, color=LGRAY)

# Table header
add_rect(s11, 0.3, 1.2, 12.73, 0.55, RED)
hdrs = ["#", "Project Name", "Description", "Status", "Platform"]
hx   = [0.35, 0.95, 4.2, 9.8, 11.4]
hw   = [0.55, 3.2,  5.5, 1.5, 1.5]
for h, lx, w in zip(hdrs, hx, hw):
    add_textbox(s11, lx, 1.25, w, 0.45, h,
                size=12, bold=True, color=WHITE, align=PP_ALIGN.LEFT)

# Table rows
rows = [
    ("1", "OSG-myG-PORTAL",         "Claims Management + WhatsApp Automation",    "Production", "Render"),
    ("2", "myG Loyalty Dashboard",  "Enterprise Analytics + AI Agent",            "Production", "Render"),
    ("3", "SHE START Dashboard",    "Startup Applicant Evaluation System",         "Production", "Render"),
    ("4", "BIGBOSS / FoneFlix",     "Registration + Video Upload Portal",          "Production", "Render"),
    ("5", "HR E-Signature Portal",  "Employee Digital Signature System",           "Production", "Render"),
    ("6", "FIFA World Cup Contest", "Match Prediction & Leaderboard Platform",     "Production", "Vercel"),
    ("7", "Enterprise AI Agent",    "Natural Language -> SQL Business Intelligence","Production","Integrated"),
    ("8", "n8n Chat Agent",         "Automated AI Chat Widget",                    "Production", "n8n Cloud"),
]
for idx, (num, name, desc, status, platform) in enumerate(rows):
    ty = 1.82 + idx * 0.61
    bg = CARD_BG if idx % 2 == 0 else RGBColor(0x1A, 0x1A, 0x2E)
    add_rect(s11, 0.3, ty, 12.73, 0.58, bg)
    vals = [num, name, desc, status, platform]
    cols_c = [WHITE, RED, LGRAY, RGBColor(0x2E,0xCC,0x71), LGRAY]
    bolds  = [True, True, False, False, False]
    for val, lx, w, c, b in zip(vals, hx, hw, cols_c, bolds):
        add_textbox(s11, lx, ty+0.04, w, 0.5, val,
                    size=11, bold=b, color=c, align=PP_ALIGN.LEFT)

add_textbox(s11, 0.4, 6.88, 12.5, 0.5,
            "Q2 2026  |  Jasil N Balussery  |  myG Technology",
            size=11, color=WHITE, align=PP_ALIGN.CENTER)


# ══════════════════════════════════════════════════════════════
# SLIDE 12: Key Achievements
# ══════════════════════════════════════════════════════════════
s12 = prs.slides.add_slide(BLANK)
add_rect(s12, 0, 0, 13.33, 7.5, DARK)
add_rect(s12, 0, 0, 13.33, 1.1, RED)
add_rect(s12, 0, 6.85, 13.33, 0.65, RED)

add_textbox(s12, 0.4, 0.05, 12, 0.65, "KEY ACHIEVEMENTS — Q2 2026",
            size=28, bold=True, color=WHITE)
add_textbox(s12, 0.4, 0.68, 12, 0.38, "10 Major Technical & Delivery Milestones",
            size=14, color=LGRAY)

achievements = [
    ("01", "8 Production Apps",      "Built and deployed 8 production applications in a single quarter"),
    ("02", "12.6M+ DB Rows",         "Managed massive DB with 26 materialized views & sub-second queries"),
    ("03", "AI/ML Models Live",      "Random Forest, LSTM, and LLM models integrated into production"),
    ("04", "WhatsApp Automation",    "4 claim status types automated via Telinfy WhatsApp API"),
    ("05", "Real-time Sync",         "Google Sheets <-> PostgreSQL sync across 3 projects"),
    ("06", "AI: 5 min -> <1 sec",    "Enterprise AI Agent response time reduced to under 1 second"),
    ("07", "OAuth Video Pipeline",   "500MB video upload via OAuth 2.0 Google Drive integration"),
    ("08", "Kerala AI Featurizer",   "Malayalam Calendar featurizer for Kerala-specific ML predictions"),
    ("09", "Pixel-Perfect UI",       "Responsive desktop & mobile designs across all portals"),
    ("10", "Zero-Downtime Pivot",    "Bigg Boss portal rebranded to FoneFlix with zero downtime"),
]
for i, (num, title, desc) in enumerate(achievements):
    row = i // 2
    col = i % 2
    lx = 0.3 + col * 6.55
    ty = 1.2 + row * 1.1
    add_rect(s12, lx, ty, 6.2, 0.95, CARD_BG)
    add_rect(s12, lx, ty, 0.55, 0.95, RED)
    add_textbox(s12, lx+0.05, ty+0.18, 0.45, 0.5, num,
                size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_textbox(s12, lx+0.65, ty+0.04, 5.4, 0.38, title,
                size=13, bold=True, color=RED)
    add_textbox(s12, lx+0.65, ty+0.45, 5.4, 0.42, desc,
                size=11, color=LGRAY)

add_textbox(s12, 0.4, 6.88, 12.5, 0.5,
            "Q2 2026  |  Jasil N Balussery  |  myG Technology",
            size=11, color=WHITE, align=PP_ALIGN.CENTER)


# ══════════════════════════════════════════════════════════════
# SLIDE 13: Thank You — modeled on template end slide (slide 42)
# ══════════════════════════════════════════════════════════════
s13 = prs.slides.add_slide(BLANK)
add_rect(s13, 0, 0, 13.33, 7.5, DARK)
add_rect(s13, 0, 0, 13.33, 3.75, RED)   # top half red
add_rect(s13, 5.5, 3.35, 2.33, 0.1, WHITE)  # divider line

add_textbox(s13, 0, 0.8, 13.33, 1.6, "THANK YOU",
            size=64, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
add_textbox(s13, 0, 2.45, 13.33, 0.7, "Q2 2026  |  April  |  May  |  June",
            size=22, bold=False, color=WHITE, align=PP_ALIGN.CENTER)
add_textbox(s13, 0, 4.1, 13.33, 0.7,
            "Jasil N Balussery",
            size=24, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
add_textbox(s13, 0, 4.8, 13.33, 0.55,
            "Technology / Software Development  |  myG",
            size=16, color=LGRAY, align=PP_ALIGN.CENTER)
add_textbox(s13, 0, 5.5, 13.33, 0.55,
            "01 April 2026 - 30 June 2026",
            size=14, color=LGRAY, align=PP_ALIGN.CENTER)


# ══════════════════════════════════════════════════════════════
# SAVE
# ══════════════════════════════════════════════════════════════
prs.save(OUTPUT)
print(f"Saved: {OUTPUT}")
print(f"Total slides: {len(prs.slides)}")
