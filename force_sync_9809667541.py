"""
Force sync remarks + onsitego from Google Sheet for a specific claim.
"""
import os, sys
sys.path.insert(0, r'c:\Users\jasil_myg\Desktop\OSG-myG-PORTAL-mainnnnn - Copy')
from dotenv import load_dotenv
load_dotenv(r'c:\Users\jasil_myg\Desktop\OSG-myG-PORTAL-mainnnnn - Copy\.env')

import requests
from services.pg_sync import fetch_claims_from_postgres, upsert_claim_to_postgres

TARGET_MOBILE = "9809667541"
TARGET_CLM    = "CLM-1789445732"

web_app_url = os.environ.get("WEB_APP_URL")
if not web_app_url:
    print("[ERROR] WEB_APP_URL not set in .env")
    sys.exit(1)

print(f"Fetching Google Sheet data...")
resp = requests.get(web_app_url, timeout=30)
if resp.status_code != 200:
    print(f"[ERROR] Sheet fetch failed: {resp.status_code}")
    sys.exit(1)

sheet_data = resp.json()
print(f"Sheet returned {len(sheet_data)} rows")

# Find our claim
sheet_row = None
for row in sheet_data:
    cid = str(row.get("Claim ID") or row.get("claim_id") or "").strip().replace(" ", "-")
    if cid == TARGET_CLM:
        sheet_row = row
        break

if not sheet_row:
    print(f"[ERROR] Claim {TARGET_CLM} not found in Google Sheet!")
    sys.exit(1)

print(f"\nFound claim in sheet: {TARGET_CLM}")

# Extract remarks and onsitego — first non-empty value wins (sheet has both 'REMARKS' and 'Remarks')
s_remarks  = ""
s_onsitego = ""
for k, v in sheet_row.items():
    kl = str(k).strip().lower()
    if kl == "remarks":
        val = str(v).strip()
        if not s_remarks and val and val.lower() not in ('nan', 'none', 'nat'):
            s_remarks = val
    elif "onsitego" in kl and "status" in kl:
        val = str(v).strip()
        if not s_onsitego and val and val.lower() not in ('nan', 'none', 'nat'):
            s_onsitego = val

print(f"Sheet Remarks      : {repr(s_remarks)}")
print(f"Sheet Onsitego Sts : {repr(s_onsitego)}")

if not s_remarks and not s_onsitego:
    print("\nNothing to sync — both are empty in the sheet.")
    sys.exit(0)

# Force upsert
partial_row = {"Claim ID": TARGET_CLM}
if s_remarks:
    partial_row["Remarks"] = s_remarks
if s_onsitego:
    partial_row["ONSITEGO - STATUS"] = s_onsitego

print(f"\nUpserting: {partial_row}")
result = upsert_claim_to_postgres(partial_row, source="sheet")
print(f"Result: {result}")
print("\nDone! Refresh the claim page to see the updated Follow Up History.")
