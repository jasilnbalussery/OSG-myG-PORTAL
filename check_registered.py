import os
from dotenv import load_dotenv
load_dotenv()
import psycopg2

conn = psycopg2.connect(os.environ.get('DATABASE_URL'))
cur = conn.cursor()

print('=== REGISTERED CLAIMS ===\n')

cur.execute("""
    SELECT claim_id, customer_name, mobile_number, date, 
           address, model, issue, status, remarks,
           assigned_staff
    FROM claims
    WHERE LOWER(status) = 'registered'
    OR LOWER(status) LIKE '%register%'
    ORDER BY date DESC
""")
rows = cur.fetchall()
print(f'Total Registered claims: {len(rows)}\n')

print(f"{'CLM_ID':<20} {'CUSTOMER':<22} {'MOBILE':<13} {'DATE':<12} {'STATUS':<15} {'ASSIGNED'}")
print('-'*100)
for r in rows:
    print(f"{str(r[0]):<20} {str(r[1])[:22]:<22} {str(r[2]):<13} {str(r[3])[:10]:<12} {str(r[7]):<15} {str(r[9] or 'Unassigned')}")

print()
print('--- Full Details ---')
for r in rows:
    print(f"\nClaim ID  : {r[0]}")
    print(f"Customer  : {r[1]}")
    print(f"Mobile    : {r[2]}")
    print(f"Date      : {r[3]}")
    print(f"Address   : {r[4]}")
    print(f"Product   : {r[5]}")
    print(f"Issue     : {r[6]}")
    print(f"Status    : {r[7]}")
    print(f"Remarks   : {r[8]}")
    print(f"Assigned  : {r[9]}")

conn.close()
print('\nDone.')
