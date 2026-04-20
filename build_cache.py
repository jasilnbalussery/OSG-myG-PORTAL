"""
One-time script: Generate cache.pkl from the Excel file.
Run this locally whenever the Excel file is updated, then git commit cache.pkl.
This ensures Render always has a pre-built cache and avoids startup delays.
"""
import pandas as pd
import pickle
import os
import sys

EXCEL_FILE = "Onsitego OSID updated upto Jan 2026.xlsx"
CACHE_FILE = "cache.pkl"

if not os.path.exists(EXCEL_FILE):
    print(f"ERROR: Excel file not found: {EXCEL_FILE}")
    sys.exit(1)

print(f"Reading Excel file: {EXCEL_FILE}...")
cols_to_use = ['Customer', 'Mobile No', 'Invoice No', 'Store Name', 'Model', 'Serial No', 'OSID', 'Date']
try:
    df = pd.read_excel(EXCEL_FILE, usecols=cols_to_use, engine='openpyxl')
except Exception as e:
    print(f"Could not read with usecols ({e}), reading all columns...")
    df = pd.read_excel(EXCEL_FILE, engine='openpyxl')

print(f"Loaded {len(df)} rows. Normalizing...")
df.columns = [str(c).strip().lower() for c in df.columns]

mob_col = None
for c in df.columns:
    if "mobile" in c or "phone" in c:
        mob_col = c
        break

if not mob_col:
    print("ERROR: No mobile column found in Excel!")
    sys.exit(1)

df = df.dropna(subset=[mob_col])
df['target_mobile_str'] = (
    df[mob_col]
    .astype(str)
    .str.replace(r'\.0$', '', regex=True)
    .str.strip()
)

print("Building index...")
# Column lookup helper
def col_lookup(df, variations):
    for v in variations:
        if v in df.columns:
            return v
    return None

name_col = col_lookup(df, ["customer", "customer name"])
inv_col = col_lookup(df, ["invoice no", "invoice", "invoice_no"])
mod_col = col_lookup(df, ["model"])
ser_col = col_lookup(df, ["serial no", "serialno", "serial_no"])
osid_col = col_lookup(df, ["osid"])
br_col = col_lookup(df, ["store name", "store_name", "branch", "branch name"])

index = {}
for row in df.to_dict('records'):
    mob = str(row.get('target_mobile_str', ''))
    if not mob: continue
    if mob not in index:
        index[mob] = {"name": str(row.get(name_col, "Unknown")), "products": []}
    index[mob]["products"].append({
        "invoice": str(row.get(inv_col, "")),
        "model": str(row.get(mod_col, "")),
        "serial": str(row.get(ser_col, "")),
        "osid": str(row.get(osid_col, "")),
        "branch": str(row.get(br_col, "Main Branch"))
    })

print(f"Index built: {len(index)} unique mobile numbers")

with open(CACHE_FILE, 'wb') as f:
    pickle.dump(index, f, protocol=4)

size_kb = os.path.getsize(CACHE_FILE) / 1024
print(f"Saved to {CACHE_FILE} ({size_kb:.1f} KB)")
print("Done! Now commit cache.pkl to GitHub:")
print("  git add cache.pkl && git commit -m 'Update customer cache' && git push")
