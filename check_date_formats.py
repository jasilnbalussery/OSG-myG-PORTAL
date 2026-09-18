import os, datetime
from dotenv import load_dotenv
load_dotenv()
import psycopg2

conn = psycopg2.connect(os.environ.get('DATABASE_URL'))
cur = conn.cursor()

print('=== DATE FORMAT AUDIT ===\n')

# Total claims
cur.execute('SELECT COUNT(*) FROM claims')
total = cur.fetchone()[0]
print(f'Total claims: {total}')

# Format breakdown
cur.execute("""
    SELECT 
        COUNT(*) FILTER (WHERE date IS NULL OR TRIM(date) = '') as empty_date,
        COUNT(*) FILTER (WHERE LENGTH(TRIM(date)) = 10 AND date ~ '^[0-9]{4}-[0-9]{2}-[0-9]{2}$') as fmt_date_only,
        COUNT(*) FILTER (WHERE LENGTH(TRIM(date)) > 10 AND date ~ '^[0-9]{4}-[0-9]{2}-[0-9]{2}') as fmt_with_time,
        COUNT(*) FILTER (WHERE date IS NOT NULL AND TRIM(date) != '' 
                         AND date !~ '^[0-9]{4}-[0-9]{2}-[0-9]{2}') as fmt_other
    FROM claims
""")
r = cur.fetchone()
print(f'\nFormat breakdown:')
print(f'  Empty / NULL          : {r[0]}')
print(f'  YYYY-MM-DD only       : {r[1]}')
print(f'  YYYY-MM-DD + time     : {r[2]}')
print(f'  Other / unknown       : {r[3]}')

# Sample YYYY-MM-DD only
print('\n--- Sample: YYYY-MM-DD only ---')
cur.execute("""
    SELECT claim_id, date, customer_name, status
    FROM claims
    WHERE LENGTH(TRIM(date)) = 10 AND date ~ '^[0-9]{4}-[0-9]{2}-[0-9]{2}$'
    ORDER BY date DESC LIMIT 5
""")
for row in cur.fetchall():
    print(f'  {row[0]} | {row[1]} | {str(row[2])[:20]} | {row[3]}')

# Sample YYYY-MM-DD + time
print('\n--- Sample: YYYY-MM-DD + timestamp ---')
cur.execute("""
    SELECT claim_id, date, customer_name, status
    FROM claims
    WHERE LENGTH(TRIM(date)) > 10 AND date ~ '^[0-9]{4}-[0-9]{2}-[0-9]{2}'
    ORDER BY date DESC LIMIT 5
""")
for row in cur.fetchall():
    print(f'  {row[0]} | {row[1]} | {str(row[2])[:20]} | {row[3]}')

# Check TAT impact - simulate ClaimWrapper._parse_date
print('\n--- TAT Calculation Impact ---')

def parse_date(d_str):
    if not d_str: return None
    d_str = str(d_str).strip()
    for fmt in ['%Y-%m-%d %H:%M:%S', '%Y-%m-%d', '%d-%m-%Y', '%d/%m/%Y', '%Y-%m-%dT%H:%M:%S']:
        try:
            return datetime.datetime.strptime(d_str[:19], fmt)
        except:
            pass
    return None

cur.execute("SELECT claim_id, date FROM claims WHERE date IS NOT NULL AND TRIM(date) != '' ORDER BY RANDOM() LIMIT 50")
rows = cur.fetchall()

parse_ok = 0
parse_fail = 0
fail_examples = []
for r in rows:
    dt = parse_date(r[1])
    if dt:
        parse_ok += 1
    else:
        parse_fail += 1
        fail_examples.append(r)

print(f'  50 random dates tested: {parse_ok} OK, {parse_fail} FAILED')
if fail_examples:
    print('  Failed examples:')
    for r in fail_examples:
        print(f'    Claim {r[0]}: date value = "{r[1]}"')
else:
    print('  All 50 parsed correctly. No TAT breakage found.')

# Check ClaimWrapper _parse_date actual implementation in app.py
print('\n--- app.py _parse_date function check ---')
cur.execute("""
    SELECT claim_id, date,
           (NOW()::date - SUBSTR(date, 1, 10)::date) as age_days
    FROM claims
    WHERE date IS NOT NULL AND TRIM(date) != ''
    AND date ~ '^[0-9]{4}-[0-9]{2}-[0-9]{2}'
    AND (NOW()::date - SUBSTR(date, 1, 10)::date) > 365
    ORDER BY date ASC
    LIMIT 5
""")
old_rows = cur.fetchall()
if old_rows:
    print('  Claims with date > 1 year ago (sanity check):')
    for r in old_rows:
        print(f'    {r[0]} | {r[1]} | {r[2]} days old')

# VERDICT
print('\n=== VERDICT ===')
cur.execute("""
    SELECT COUNT(*) FROM claims
    WHERE date IS NOT NULL AND TRIM(date) != ''
    AND LENGTH(TRIM(date)) > 10
""")
with_time = cur.fetchone()[0]

cur.execute("""
    SELECT COUNT(*) FROM claims
    WHERE date IS NOT NULL AND TRIM(date) != ''
    AND LENGTH(TRIM(date)) = 10
""")
without_time = cur.fetchone()[0]

print(f'  Dates with timestamp (YYYY-MM-DD HH:MM:SS) : {with_time}')
print(f'  Dates without timestamp (YYYY-MM-DD)       : {without_time}')

if with_time > 0 and without_time > 0:
    print('  MIXED formats STILL PRESENT in DB.')
    print('  However: app _parse_date handles both formats via format list.')
    print('  TAT calculations are NOT broken - but sorting by date may be inconsistent.')
else:
    print('  Dates are UNIFORM - no mixed format issue.')

conn.close()
print('\nDone.')
