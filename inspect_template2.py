"""Deep-inspect key template slides to understand shapes, positions, text frames."""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE_TYPE

template_path = r"c:\Users\jasil_myg\Desktop\OSG-myG-PORTAL-mainnnnn - Copy\Challenging Success Business PowerPoint Templates.pptx"
prs = Presentation(template_path)

# Inspect slides 0,1,2,3,5,7,8,9 — likely most useful
INSPECT = [0, 1, 2, 3, 5, 7, 9, 10, 19, 28, 34, 40, 41, 42]

for si in INSPECT:
    slide = prs.slides[si]
    print(f"\n{'='*60}")
    print(f"SLIDE {si} | layout: '{slide.slide_layout.name}' | shapes: {len(slide.shapes)}")
    print(f"{'='*60}")
    for sh in slide.shapes:
        left  = round(sh.left  / 914400, 2) if sh.left  else 0
        top   = round(sh.top   / 914400, 2) if sh.top   else 0
        width = round(sh.width / 914400, 2) if sh.width else 0
        height= round(sh.height/ 914400, 2) if sh.height else 0
        stype = sh.shape_type
        name  = sh.name
        txt   = ""
        if sh.has_text_frame:
            for para in sh.text_frame.paragraphs:
                t = para.text.strip()
                if t:
                    txt = t[:60]
                    break
        print(f"  [{name}] type={stype} pos=({left},{top}) size=({width}x{height}) text='{txt}'")
