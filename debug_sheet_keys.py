"""
Debug: Print ALL keys/values from the Google Sheet row for CLM-1789445732
to find the exact key used for Remarks column.
"""
import os, sys
sys.path.insert(0, r'c:\Users\jasil_myg\Desktop\OSG-myG-PORTAL-mainnnnn - Copy')
from dotenv import load_dotenv
load_dotenv(r'c:\Users\jasil_myg\Desktop\OSG-myG-PORTAL-mainnnnn - Copy\.env')

import requests

TARGET_CLM = "CLM-1789445732"

web_app_url = os.environ.get("WEB_APP_URL")
resp = requests.get(web_app_url, timeout=30)
sheet_data = resp.json()

for row in sheet_data:
    cid = str(row.get("Claim ID") or row.get("claim_id") or "").strip().replace(" ", "-")
    if cid == TARGET_CLM:
        print(f"=== ALL KEYS for {TARGET_CLM} ===")
        for k, v in row.items():
            print(f"  {repr(k):45s} => {repr(str(v))}")
        break
