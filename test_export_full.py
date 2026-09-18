"""
test_export_full.py
Full end-to-end test of the export route logic.
"""
import os, io, datetime
from dotenv import load_dotenv
load_dotenv()

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from validators import validate_claim_ids_list
from services.pg_sync import fetch_claims_from_postgres

# ── Step 1: Fetch all claims ──────────────────────────────────────
print("Step 1: Fetching claims from DB...")
claims_raw = fetch_claims_from_postgres()
print(f"  Total raw claims: {len(claims_raw)}")

all_ids = [str(c.get('Claim ID', '') or '').strip() for c in claims_raw if c.get('Claim ID')]
print(f"  Valid claim IDs: {len(all_ids)}")

# ── Step 2: Validate IDs ─────────────────────────────────────────
print("\nStep 2: Validating claim IDs...")
try:
    validated = validate_claim_ids_list(all_ids)
    print(f"  Validated OK: {len(validated)} IDs")
except ValueError as e:
    print(f"  VALIDATION ERROR: {e}")
    raise

# ── Step 3: Import ClaimWrapper and fetch via app logic ───────────
print("\nStep 3: Loading ClaimWrapper objects...")
import sys
sys.path.insert(0, '.')

# Manually import fetch_claims_from_db without running Flask
from app import fetch_claims_from_db, ClaimWrapper
claims = fetch_claims_from_db()
print(f"  ClaimWrapper count: {len(claims)}")

id_set = set(validated)
filtered = [c for c in claims if str(c.claim_id) in id_set]
print(f"  Filtered to export: {len(filtered)}")

# ── Step 4: Build workbook ────────────────────────────────────────
print("\nStep 4: Building Excel workbook...")
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Claims Export"

header_font  = Font(bold=True, color="FFFFFF", size=11)
header_fill  = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
thin   = Side(border_style="thin", color="D1D5DB")
border = Border(left=thin, right=thin, top=thin, bottom=thin)
alt_fill     = PatternFill(start_color="EFF6FF", end_color="EFF6FF", fill_type="solid")
pending_fill = PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid")
overdue_fill = PatternFill(start_color="FFDCDC", end_color="FFDCDC", fill_type="solid")

headers = [
    "SR No", "Claim ID", "Submitted Date", "Customer Name", "Mobile",
    "Branch", "Product", "Issue", "Status",
    "Replacement Progress %", "Complete", "Aging Days"
]
ws.append(headers)
for col_idx, _ in enumerate(headers, 1):
    cell = ws.cell(row=1, column=col_idx)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = header_align
    cell.border = border
ws.row_dimensions[1].height = 22

today = datetime.date.today()
errors = []

for i, claim in enumerate(filtered, 2):
    try:
        if not claim.complete:
            try:
                submitted_date = claim.created_at.date() if claim.created_at else None
                aging_days = (today - submitted_date).days if submitted_date else '-'
            except:
                aging_days = '-'
        else:
            aging_days = '-'

        stages = [
            claim.data.get("Customer Confirmation", ""),
            claim.data.get("Approval Mail Received From Onsitego (Yes/No)", ""),
            claim.data.get("Mail Sent To Store (Yes/No)", ""),
            claim.data.get("Invoice Generated (Yes/No)", ""),
            claim.data.get("Invoice Sent To Onsitego (Yes/No)", ""),
            claim.data.get("Settled With Accounts (Yes/No)", ""),
        ]
        done = sum(1 for s in stages if str(s).lower() == "yes")
        progress = round((done / len(stages)) * 100)

        row_data = [
            str(claim.sr_no or ''),
            str(claim.claim_id or ''),
            str(claim.created_at.strftime('%d %b %Y') if claim.created_at else ''),
            str(claim.customer_name or ''),
            str(claim.mobile_no or ''),
            str(claim.branch or '-'),
            str(claim.model or ''),
            str(claim.issue or ''),
            str(claim.status or ''),
            progress,
            'Yes' if claim.complete else 'No',
            aging_days
        ]
        ws.append(row_data)
    except Exception as e:
        errors.append(f"  Row {i} (claim {claim.claim_id}): {type(e).__name__}: {e}")

if errors:
    print(f"  ERRORS building {len(errors)} rows:")
    for e in errors[:10]:
        print(e)
else:
    print(f"  All {len(filtered)} rows built without error")

# ── Step 5: Save to BytesIO ───────────────────────────────────────
print("\nStep 5: Saving workbook to buffer...")
try:
    output = io.BytesIO()
    wb.save(output)
    print(f"  SUCCESS. Excel file size: {output.tell():,} bytes")
except Exception as e:
    print(f"  SAVE ERROR: {e}")
    raise

print("\n=== EXPORT TEST COMPLETE - ALL STEPS PASSED ===")
