from app import fetch_claims_from_sheet, get_ist_now
import datetime
import traceback

claims = fetch_claims_from_sheet(force_refresh=True)
now = get_ist_now().replace(tzinfo=None)

report_stats = {
    'gst_invoice': {'lt5': 0, 'gt5': 0, 'gt10': 0, 'total': 0},
    'grand_total_replacement': 0
}

count = 0
for c in claims:
    age = (now - c.created_at.replace(tzinfo=None)).days if c.created_at else 0
    repl_age = age
    settled_date_raw = c.claim_settled_date
    if settled_date_raw and str(settled_date_raw).strip() not in ('', 'nan', 'None'):
        try:
            settled_dt = datetime.datetime.strptime(str(settled_date_raw).strip()[:10], '%Y-%m-%d')
            repl_age = (now - settled_dt).days
        except Exception:
            try:
                settled_dt = datetime.datetime.strptime(str(settled_date_raw).strip()[:10], '%d-%m-%Y')
                repl_age = (now - settled_dt).days
            except: pass

    status = (c.status or "").strip().lower()

    if "replacement" in status or c.mail_sent_to_store:
        if c.settled_with_accounts:
            pass
        elif c.settlement_mail_accounts:
            pass
        elif c.invoice_sent_osg:
            pass
        elif c.invoice_generated:
            report_stats['gst_invoice']['total'] += 1
            report_stats['grand_total_replacement'] += 1
            
            def _parse_date(raw):
                if not raw or str(raw).strip() in ('', 'nan', 'None'): return None
                s = str(raw).strip()[:10]
                dt = None
                for fmt in ('%Y-%m-%d', '%d-%m-%Y', '%d/%m/%Y'):
                    try:
                        dt = datetime.datetime.strptime(s, fmt)
                        break
                    except: continue
                
                if dt and (dt - now).days > 1:
                    try:
                        if '-' in s and len(s.split('-')[0]) == 4:
                            dt = datetime.datetime.strptime(s, '%Y-%d-%m')
                    except: pass
                return dt

            inv_gen_dt = _parse_date(c.invoice_generated_date)
            gst_age = -1
            if inv_gen_dt:
                gst_age = max(0, (now - inv_gen_dt).days)
                dt_used = "Invoice"
            else:
                store_dt = _parse_date(c.mail_sent_to_store_date)
                if store_dt:
                    gst_age = max(0, (now - store_dt).days)
                    dt_used = "Store"
                else:
                    gst_age = max(0, repl_age)
                    dt_used = "Repl Age"

            if gst_age <= 5: 
                report_stats['gst_invoice']['lt5'] += 1
            elif gst_age <= 10: 
                report_stats['gst_invoice']['gt5'] += 1
            else: 
                report_stats['gst_invoice']['gt10'] += 1
                
            print(f"Claim: {c.claim_id} | Mobile: {c.mobile_no} | Age: {gst_age} | Used: {dt_used} | Raw Inv: {c.invoice_generated_date}")

print("\nFINAL BUCKETS:", report_stats['gst_invoice'])
