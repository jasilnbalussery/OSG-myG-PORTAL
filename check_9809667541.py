import os, sys
sys.path.insert(0, r'c:\Users\jasil_myg\Desktop\OSG-myG-PORTAL-mainnnnn - Copy')
from dotenv import load_dotenv
load_dotenv(r'c:\Users\jasil_myg\Desktop\OSG-myG-PORTAL-mainnnnn - Copy\.env')
import psycopg2
import psycopg2.extras

conn = psycopg2.connect(os.environ['DATABASE_URL'])
cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
cur.execute("SELECT claim_id, mobile_number, remarks, follow_up_notes FROM claims WHERE mobile_number LIKE %s", ('%9809667541%',))
rows = cur.fetchall()
print(f'Found {len(rows)} claims for 9809667541')
for r in rows:
    print('---')
    print(f'Claim ID     : {r["claim_id"]}')
    print(f'Mobile       : {r["mobile_number"]}')
    print(f'Remarks      : {repr(r["remarks"])}')
    print(f'Follow Notes : {repr(r["follow_up_notes"])}')
conn.close()
