"""
backfill_remarks_to_notes.py
-----------------------------
One-time script to backfill existing `remarks` values into `follow_up_notes`
for ALL claims in the DB that have a remark but haven't logged it yet.

Run:
    python backfill_remarks_to_notes.py

What it does:
  - Reads every claim that has a non-empty `remarks` column
  - Checks if "[REMARK]:" already appears in `follow_up_notes`
  - If not, appends: [dd/mm/yyyy, hh:mm:ss am/pm] [REMARK]: <remark value>
  - Prints a summary at the end
"""

import os
import datetime
import psycopg2
import psycopg2.extras
from dotenv import load_dotenv

load_dotenv()

DATABASE_URL = os.environ.get("DATABASE_URL")
if not DATABASE_URL:
    raise EnvironmentError("DATABASE_URL is not set in .env")

conn = psycopg2.connect(DATABASE_URL)
conn.autocommit = False

try:
    with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
        # Fetch all claims that have a non-empty remarks
        cur.execute("""
            SELECT claim_id, remarks, follow_up_notes
            FROM claims
            WHERE remarks IS NOT NULL AND trim(remarks) != ''
              AND trim(remarks) NOT IN ('nan', 'none', 'nat')
        """)
        rows = cur.fetchall()

    print(f"Found {len(rows)} claims with existing remarks.\n")

    updated = 0
    skipped = 0

    with conn.cursor() as cur:
        for row in rows:
            claim_id   = row["claim_id"] or ""
            remarks    = (row["remarks"] or "").strip()
            notes      = (row["follow_up_notes"] or "").strip()

            # Skip if this remark was already logged
            if "[REMARK]:" in notes and remarks.lower() in notes.lower():
                skipped += 1
                continue

            # Build the timestamped entry (use a generic backfill timestamp)
            ts = datetime.datetime.now().strftime('%d/%m/%Y, %I:%M:%S %p').lower()
            new_entry = f"[{ts}] [REMARK]: {remarks}"

            updated_notes = (notes + "\n" + new_entry).strip()

            cur.execute(
                'UPDATE claims SET follow_up_notes = %s WHERE claim_id = %s',
                (updated_notes, claim_id)
            )
            print(f"  [OK] [{claim_id}] Appended remark: \"{remarks[:60]}{'...' if len(remarks) > 60 else ''}\"")
            updated += 1

    conn.commit()
    print(f"\n--- Backfill complete ---")
    print(f"  Updated : {updated} claims")
    print(f"  Skipped : {skipped} claims (remark already in notes)")

except Exception as e:
    conn.rollback()
    print(f"\n[ERROR]: {e}")
    raise
finally:
    conn.close()
