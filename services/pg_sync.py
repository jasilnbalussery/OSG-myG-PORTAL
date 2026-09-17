"""
pg_sync.py — PostgreSQL Service
------------------------------------------------------
Manages claim data in a PostgreSQL database hosted on DigitalOcean.

Usage:
    from services.pg_sync import fetch_claims_from_postgres, upsert_claim_to_postgres

Environment Variables required:
    DATABASE_URL  — Full PostgreSQL connection string
"""

import os
import json
import logging
import datetime
import requests
import psycopg2
import psycopg2.extras
from psycopg2 import sql

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Column definitions — these map 1-to-1 with the Google Sheet headers
SHEET_COLUMNS = [
    "Claim ID",
    "Customer Name",
    "Mobile Number",
    "Address",
    "Product",
    "Invoice Number",
    "Serial Number",
    "SR No",
    "Model",
    "OSID",
    "Issue",
    "Branch",
    "Follow Up - Dates",
    "Follow Up - Notes",
    "Claim Settled Date",
    "Remarks",
    "Status",
    # Replacement workflow
    "Replacement: Confirmation Pending",
    "Replacement: OSG Approval",
    "Replacement: Mail to Store",
    "Replacement: Invoice Generated",
    "Replacement: Invoice Sent to OSG",
    "Replacement: Settled with Accounts",
    "Replacement: Settlement Mail to Accounts",
    # Dates for replacement steps
    "Approval Mail Received Date",
    "Mail Sent To Store Date",
    "Invoice Generated Date",
    "Invoice Sent To Onsitego Date",
    # Other
    "Complete",
    "Settled Time (TAT)",
    "Assigned Staff",
    "Feedback Rating",
    "Repair Feedback Completed (Yes/No)",
    "Settlement Mail to Accounts(Yes/No)",
    "Last Updated Timestamp",
    "Last_Notified_Status",
    
    # Alternative legacy column names used by app.py
    "Customer Confirmation",
    "Approval Mail Received From Onsitego (Yes/No)",
    "Mail Sent To Store (Yes/No)",
    "Invoice Generated (Yes/No)",
    "Invoice Sent To Onsitego (Yes/No)",
    "Settlement Mail to Accounts(Yes/No)",
    "Settlement Mail to Accounts Date",
    "Settled With Accounts (Yes/No)",
    "Complete (Yes/No)",
]

# Postgres-safe column name mapping  (Sheet Header → DB column name)
def _pg_col(name: str) -> str:
    """Convert a Sheet column header to a safe PostgreSQL column name."""
    import re
    clean_name = name.strip().lower()
    clean_name = (
        clean_name
        .replace(" ", "_")
        .replace("-", "_")
        .replace("(", "")
        .replace(")", "")
        .replace("/", "_")
        .replace(":", "")
        .replace(".", "")
        .replace(",", "")
        .replace("'", "")
        .replace("?", "")
        .replace("!", "")
    )
    # Collapse multiple consecutive underscores into a single underscore
    # This prevents "Follow Up - Dates" becoming "follow_up___dates" instead of "follow_up_dates"
    clean_name = re.sub(r'_+', '_', clean_name)
    return clean_name.strip("_")


# Pre-build the column map at import time
COL_MAP = {col: _pg_col(col) for col in SHEET_COLUMNS}
PG_COLS = list(COL_MAP.values())  # ordered list of DB column names


# ---------------------------------------------------------------------------
# Database helpers
# ---------------------------------------------------------------------------

def _get_connection():
    """Return a new psycopg2 connection using DATABASE_URL."""
    db_url = os.environ.get("DATABASE_URL")
    if not db_url:
        raise EnvironmentError("DATABASE_URL environment variable is not set.")
    return psycopg2.connect(db_url)


def ensure_table_exists(conn):
    """
    Create the `claims` table if it doesn't exist.
    All columns are TEXT to faithfully preserve every value from the Sheet.
    """
    col_defs = ",\n    ".join(
        f'"{pg_col}" TEXT' for pg_col in PG_COLS if pg_col != "claim_id"
    )
    ddl = f"""
        CREATE TABLE IF NOT EXISTS claims (
            claim_id TEXT PRIMARY KEY,
            {col_defs},
            synced_at TIMESTAMPTZ DEFAULT NOW()
        );
    """
    with conn.cursor() as cur:
        cur.execute(ddl)
    conn.commit()
    logger.info("[PG_SYNC] Ensured `claims` table exists.")


def _add_missing_columns(conn, sheet_row_keys: list):
    """
    Dynamically add any columns found in the sheet that don't yet exist
    in the Postgres table. Keeps the DB in sync as the sheet grows.
    """
    with conn.cursor() as cur:
        cur.execute(
            "SELECT column_name FROM information_schema.columns WHERE table_name = 'claims';"
        )
        existing = {row[0] for row in cur.fetchall()}

    new_cols = []
    for raw_key in sheet_row_keys:
        pg = _pg_col(raw_key)
        if pg and pg not in existing and pg != "claim_id":
            new_cols.append((raw_key, pg))

    if new_cols:
        with conn.cursor() as cur:
            for raw_key, pg in new_cols:
                try:
                    cur.execute(f'ALTER TABLE claims ADD COLUMN IF NOT EXISTS "{pg}" TEXT;')
                    logger.info(f"[PG_SYNC] Auto-added column: {pg} (from sheet: {raw_key})")
                except Exception as e:
                    logger.warning(f"[PG_SYNC] Could not add column {pg}: {e}")
        conn.commit()


# ---------------------------------------------------------------------------
# Workflow columns that must never be overwritten with NULL by a Google Sheets
# sync.  The portal sets these via the Edit-Claim form; if the Sheet doesn't
# carry them (NULL incoming), we keep whatever value is already in the DB.
# ---------------------------------------------------------------------------
PROTECTED_WORKFLOW_COLS = {
    # New-style column names (portal saves here)
    "customer_confirmation",
    "approval_mail_received_from_onsitego_yes_no",
    "mail_sent_to_store_yes_no",
    "invoice_generated_yes_no",
    "invoice_sent_to_onsitego_yes_no",
    "settlement_mail_to_accounts_yes_no",
    "settled_with_accounts_yes_no",
    "complete_yes_no",
    # Old-style column names (legacy Sheet columns)
    "replacement_confirmation_pending",
    "replacement_osg_approval",
    "replacement_mail_to_store",
    "replacement_invoice_generated",
    "replacement_invoice_sent_to_osg",
    "replacement_settled_with_accounts",
    "replacement_settlement_mail_to_accounts",
    # Dates for workflow steps
    "approval_mail_received_date",
    "mail_sent_to_store_date",
    "invoice_generated_date",
    "invoice_sent_to_onsitego_date",
    # Completion flag
    "complete",
}

# ---------------------------------------------------------------------------
# Core Insert/Update function
# ---------------------------------------------------------------------------

def upsert_claim_to_postgres(claim_data: dict, source: str = "sheet") -> dict:
    """
    Upsert a single claim into PostgreSQL.

    Args:
        claim_data (dict): The claim data dictionary (keys matching sheet headers).
        source (str): 'portal' when called from the portal UI, 'sheet' when called
                      from the Google Sheet webhook or background poller.
                      Portal writes stamp portal_last_updated and always win on status.
                      Sheet writes skip status change if portal updated within 30s.
    Returns:
        dict with keys: success (bool), error (str|None)
    """
    try:
        conn = _get_connection()
    except Exception as e:
        logger.error(f"[PG_DB] DB connection failed: {e}")
        return {"success": False, "error": str(e)}

    try:
        ensure_table_exists(conn)
        # Ensure portal_last_updated sentinel column exists before the INSERT uses it
        _add_missing_columns(conn, list(claim_data.keys()) + ["portal_last_updated", "synced_at"])

        with conn.cursor() as cur:
            claim_id = str(claim_data.get("Claim ID") or claim_data.get("claim_id") or "").strip()
            if not claim_id:
                # Generate a claim ID if missing
                import time
                claim_id = f"CLM-{int(time.time())}"
                claim_data["Claim ID"] = claim_id

            # Check if claim is new
            cur.execute("SELECT * FROM claims WHERE claim_id = %s", (claim_id,))
            existing_claim_tuple = cur.fetchone()
            is_new_claim = existing_claim_tuple is None
            
            existing_dict = {}
            if not is_new_claim:
                col_names = [desc[0] for desc in cur.description]
                existing_dict = dict(zip(col_names, existing_claim_tuple))

            # Dynamically append Remarks and ONSITEGO STATUS to Follow Up Notes
            # Note: sheet may have both 'REMARKS' (uppercase) and 'Remarks' columns — pick first non-empty
            remarks_val = ""
            remarks_key = ""
            for k, v in claim_data.items():
                if str(k).strip().lower() == "remarks":
                    val = str(v).strip()
                    if not remarks_val and val and val.lower() not in ('nan', 'none', 'nat'):
                        remarks_val = val  # first non-empty wins
                        remarks_key = k
            
            # Find onsitego status from keys (case-insensitive)
            onsitego_val = ""
            onsitego_key = ""
            for k, v in claim_data.items():
                if "onsitego" in str(k).lower() and "status" in str(k).lower():
                    onsitego_val = str(v).strip()
                    onsitego_key = k
                    break
                    
            notes_val = str(claim_data.get("Follow Up - Notes") or claim_data.get("follow_up___notes") or claim_data.get("follow_up_notes") or "").strip()
            if not notes_val and not is_new_claim:
                # Check both old (triple-underscore) and new (single-underscore) notes columns
                notes_val = str(
                    existing_dict.get("follow_up_notes") or
                    existing_dict.get("follow_up___notes") or
                    ""
                ).strip()
            
            if remarks_val.lower() in ('nan', 'none', 'nat'): remarks_val = ""
            if onsitego_val.lower() in ('nan', 'none', 'nat'): onsitego_val = ""
            if notes_val.lower() in ('nan', 'none', 'nat'): notes_val = ""
            
            # Get old values
            old_remarks = str(existing_dict.get("remarks") or "").strip() if not is_new_claim else ""
            pg_onsitego_col = _pg_col(onsitego_key) if onsitego_key else None
            old_onsitego = str(existing_dict.get(pg_onsitego_col) or "").strip() if pg_onsitego_col and not is_new_claim else ""
            
            ts = datetime.datetime.now().strftime('%d/%m/%Y, %I:%M:%S %p').lower()
            
            appended = False
            if remarks_val and remarks_val.lower() != old_remarks.lower():
                notes_val += f"\n[{ts}] [REMARK]: {remarks_val}"
                appended = True
            if onsitego_val and onsitego_val.lower() != old_onsitego.lower() and onsitego_val.lower() not in notes_val.lower():
                notes_val += f"\n[{ts}] [ONSITEGO STATUS]: {onsitego_val}"
                appended = True
                
            claim_data["Follow Up - Notes"] = notes_val.strip()
            
            if appended:
                status_key = "Status"
                for k in claim_data.keys():
                    if str(k).strip().lower() == "status":
                        status_key = k
                        break

                # Capture the ORIGINAL incoming sheet status BEFORE any mutation.
                # (Reading it after the force-to-Follow-Up would give "Follow Up",
                # making the protection check below useless — that was the old bug.)
                original_sheet_status = str(claim_data.get(status_key, "")).strip()
                old_db_status = str(existing_dict.get("status") or "").strip()

                TERMINAL_STATUSES = {
                    "repair completed", "replacement approved", "replacement closed",
                    "rejected", "closed", "settled", "cancelled", "no issue/oncall resolution",
                }
                PROTECTED_STATUSES = {
                    "follow up", "replacement approved", "replacement closed",
                    "repair completed", "rejected", "closed", "settled", "cancelled",
                    "no issue/oncall resolution",
                }
                DOWNGRADE_STATUSES = {"", "registered", "submitted"}

                # Race Condition Guard (Cause 2 Fix):
                # If this is a SHEET/POLLER write and the portal updated this claim
                # within the last 30 seconds, unconditionally preserve the DB status.
                # This prevents the sheet poller from overwriting a portal status
                # change that hasn't fully committed yet (race condition).
                GRACE_PERIOD_SECONDS = 30
                if source == "sheet" and not is_new_claim:
                    portal_ts_raw = existing_dict.get("portal_last_updated")
                    if portal_ts_raw:
                        try:
                            if isinstance(portal_ts_raw, str):
                                portal_ts = datetime.datetime.fromisoformat(portal_ts_raw.replace("Z", "+00:00"))
                                # Make naive for comparison
                                portal_ts = portal_ts.replace(tzinfo=None)
                            else:
                                portal_ts = portal_ts_raw.replace(tzinfo=None)
                            seconds_since_portal = (datetime.datetime.utcnow() - portal_ts).total_seconds()
                            if seconds_since_portal < GRACE_PERIOD_SECONDS:
                                claim_data[status_key] = old_db_status
                                logger.info(
                                    f"[PG_SYNC] Race condition guard triggered for {claim_id}: "
                                    f"portal updated {seconds_since_portal:.1f}s ago — preserving "
                                    f"DB status '{old_db_status}' over sheet source."
                                )
                        except Exception as ts_err:
                            logger.warning(f"[PG_SYNC] Could not parse portal_last_updated for {claim_id}: {ts_err}")

                # Rule 1: If the DB already holds a protected/elevated status,
                #         NEVER downgrade it — regardless of what the Sheet says.
                #         Just log the remark/onsitego note and leave the status alone.
                if old_db_status.lower() in PROTECTED_STATUSES:
                    claim_data[status_key] = old_db_status  # preserve DB status
                    logger.info(
                        f"[PG_SYNC] Status protected for {claim_id}: "
                        f"DB='{old_db_status}' | Sheet='{original_sheet_status}' "
                        f"→ keeping DB status (append triggered Follow-Up guard)"
                    )

                # Rule 2: Only force "Follow Up" if the incoming sheet status is NOT
                #         terminal AND the DB status is also not protected.
                elif original_sheet_status.lower() not in TERMINAL_STATUSES:
                    claim_data[status_key] = "Follow Up"

            else:
                # If not appended, protect elevated/terminal statuses from being
                # downgraded by a Google Sheets sync that still shows "Registered".
                if not is_new_claim:
                    old_status = str(existing_dict.get("status") or "").strip()
                    sheet_status = ""
                    status_key = "Status"
                    for k, v in claim_data.items():
                        if str(k).strip().lower() == "status":
                            sheet_status = str(v).strip()
                            status_key = k
                            break

                    PROTECTED_STATUSES = {
                        "follow up",
                        "replacement approved",
                        "replacement closed",
                        "repair completed",
                        "rejected",
                        "closed",
                        "settled",
                        "cancelled",
                    }
                    DOWNGRADE_STATUSES = {"", "registered", "submitted"}

                    if old_status.lower() in PROTECTED_STATUSES and sheet_status.lower() in DOWNGRADE_STATUSES:
                        claim_data[status_key] = old_status  # preserve existing DB status

            # Build column→value dict for this row
            col_vals = {"claim_id": claim_id}
            for raw_key, val in claim_data.items():
                pg = _pg_col(raw_key)
                if pg and pg != "claim_id":
                    v_str = str(val) if val is not None else None
                    if pg == 'date' and v_str and 'T' in v_str and v_str.endswith('Z'):
                        try:
                            import pytz
                            s_clean = v_str.replace('T', ' ')[:19]
                            dt_utc = datetime.datetime.strptime(s_clean, '%Y-%m-%d %H:%M:%S')
                            dt_utc = dt_utc.replace(tzinfo=pytz.UTC)
                            ist = pytz.timezone('Asia/Kolkata')
                            v_str = dt_utc.astimezone(ist).strftime('%Y-%m-%d')
                        except Exception:
                            pass
                    col_vals[pg] = v_str

            # Always stamp the sync time
            col_vals["synced_at"] = datetime.datetime.utcnow().isoformat()

            # Portal writes stamp portal_last_updated so the race condition guard
            # can detect when the sheet poller fires too soon after a portal save.
            if source == "portal":
                col_vals["portal_last_updated"] = datetime.datetime.utcnow().isoformat()

            cols = list(col_vals.keys())
            vals = [col_vals[c] for c in cols]

            # Build the upsert (INSERT … ON CONFLICT … DO UPDATE SET …)
            # For PROTECTED_WORKFLOW_COLS: use COALESCE so that a NULL from the
            # Sheet never overwrites a value the portal already saved.
            # For all other columns: normal EXCLUDED.col replacement.
            insert_cols = sql.SQL(", ").join(sql.Identifier(c) for c in cols)
            placeholders = sql.SQL(", ").join(sql.Placeholder() * len(cols))

            update_parts = []
            for c in cols:
                if c in ["claim_id", "date", "submitted_date"]:
                    continue  # never overwrite these
                if c in PROTECTED_WORKFLOW_COLS:
                    # COALESCE: keep existing value if incoming is NULL
                    update_parts.append(
                        sql.SQL("{col} = COALESCE(EXCLUDED.{col}, claims.{col})").format(
                            col=sql.Identifier(c)
                        )
                    )
                else:
                    update_parts.append(
                        sql.SQL("{col} = EXCLUDED.{col}").format(col=sql.Identifier(c))
                    )
            update_set = sql.SQL(", ").join(update_parts)

            upsert_query = sql.SQL(
                "INSERT INTO claims ({cols}) VALUES ({vals}) "
                "ON CONFLICT (claim_id) DO UPDATE SET {updates}"
            ).format(cols=insert_cols, vals=placeholders, updates=update_set)

            cur.execute(upsert_query, vals)
        conn.commit()
        return {"success": True, "error": None, "claim_id": claim_id}

    except Exception as e:
        conn.rollback()
        logger.error(f"[PG_DB] Failed to upsert claim: {e}", exc_info=True)
        return {"success": False, "error": str(e)}
    finally:
        conn.close()


# ---------------------------------------------------------------------------
# Read back from PostgreSQL
# ---------------------------------------------------------------------------

def fetch_claims_from_postgres() -> list:
    """
    Read all claims from PostgreSQL and return them as a list of dicts
    (same shape as the Google Sheets JSON, so ClaimWrapper works unchanged).
    """
    try:
        conn = _get_connection()
    except Exception as e:
        logger.error(f"[PG_SYNC] DB connection failed for read: {e}")
        return []

    try:
        with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
            cur.execute("SELECT * FROM claims ORDER BY synced_at DESC;")
            rows = cur.fetchall()

        # Convert RealDictRow → plain dict, and map PG column names back to
        # original Sheet header names so ClaimWrapper keeps working.
        reverse_map = {v: k for k, v in COL_MAP.items()}
        reverse_map["claim_id"] = "Claim ID"

        result = []
        for row in rows:
            d = {}
            for pg_col, val in row.items():
                sheet_key = reverse_map.get(pg_col, pg_col)
                d[sheet_key] = val
            result.append(d)

        logger.info(f"[PG_SYNC] Read {len(result)} claims from PostgreSQL.")
        return result

    except Exception as e:
        logger.error(f"[PG_SYNC] Failed to read from PostgreSQL: {e}", exc_info=True)
        return []
    finally:
        conn.close()


def test_connection() -> dict:
    """Test the PostgreSQL connection and return status info."""
    try:
        conn = _get_connection()
        with conn.cursor() as cur:
            cur.execute("SELECT version();")
            version = cur.fetchone()[0]
            cur.execute("SELECT COUNT(*) FROM claims;") if _table_exists(conn) else None
        conn.close()
        return {"success": True, "version": version}
    except Exception as e:
        return {"success": False, "error": str(e)}


def _table_exists(conn) -> bool:
    with conn.cursor() as cur:
        cur.execute(
            "SELECT EXISTS (SELECT 1 FROM information_schema.tables WHERE table_name = 'claims');"
        )
        return cur.fetchone()[0]
