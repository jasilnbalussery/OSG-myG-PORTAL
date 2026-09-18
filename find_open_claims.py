"""
EXACT replica of the JS openClaims filter - no extra exclusions.
JS code:
  if (s === 'settled' || s === 'closed' || s === 'repair completed' || 
      s === 'rejected' || s.includes('no issue')) return false;
  if (s.includes('replacement') && s.includes('approved')) return !c.complete;
  return true;
NOTE: 'cancelled' is NOT in the JS exclusion list!
"""
import os, sys
from dotenv import load_dotenv
load_dotenv()
sys.path.insert(0, '.')

from app import fetch_claims_from_db

print("Loading claims...")
claims = fetch_claims_from_db()
print(f"Total: {len(claims)}\n")

open_claims = []

for c in claims:
    s = (c.status or '').lower().strip()
    
    # JS exclusions EXACTLY
    if s in ('settled', 'closed', 'repair completed', 'rejected'):
        continue
    if 'no issue' in s:
        continue
    
    # Replacement approved check
    if 'replacement' in s and 'approved' in s:
        if not c.complete:
            open_claims.append(c)
        continue
    
    # Everything else = OPEN (including 'cancelled', 'registered', 'follow up', etc.)
    open_claims.append(c)

print(f"=== TOTAL OPEN (JS exact simulation): {len(open_claims)} ===\n")

# Group by status
by_status = {}
for c in open_claims:
    st = c.status or 'UNKNOWN'
    by_status.setdefault(st, []).append(c)

print("By status:")
for status, items in sorted(by_status.items(), key=lambda x: -len(x[1])):
    print(f"  {status:<45}: {len(items)}")

print()
# Show any unexpected (not Follow Up, not Replacement Approved)
unexpected = [c for c in open_claims 
              if (c.status or '').lower() not in ('follow up',) 
              and 'replacement' not in (c.status or '').lower()]

if unexpected:
    print(f"=== NON-FOLLOWUP NON-REPLACEMENT OPEN CLAIMS: {len(unexpected)} ===")
    print(f"{'CLM_ID':<20} {'CUSTOMER':<24} {'DATE':<12} {'STATUS':<20} {'COMPLETE'}")
    print('-'*90)
    for c in unexpected:
        print(f"{str(c.claim_id):<20} {str(c.customer_name)[:24]:<24} "
              f"{str(c.created_at)[:10]:<12} {str(c.status):<20} {c.complete}")
