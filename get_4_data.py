import sys
import datetime
import json
sys.path.append(r'c:\Users\jasil_myg\Desktop\OSG-myG-PORTAL-mainnnnn')
import app

with app.app.app_context():
    claims = app.fetch_claims_from_sheet(force_refresh=True)
    now = app.get_ist_now().replace(tzinfo=None)
    res = []
    
    for c in claims:
        age = (now - c.created_at.replace(tzinfo=None)).days if c.created_at else 0
        settled_date_raw = c.claim_settled_date
        repl_age = age
        if settled_date_raw and str(settled_date_raw).strip() not in ('', 'nan', 'None'):
            try:
                settled_dt = datetime.datetime.strptime(str(settled_date_raw).strip()[:10], '%Y-%m-%d')
                repl_age = (now - settled_dt).days
            except Exception:
                try:
                    settled_dt = datetime.datetime.strptime(str(settled_date_raw).strip()[:10], '%d-%m-%Y')
                    repl_age = (now - settled_dt).days
                except Exception:
                    repl_age = age
                    
        status = (c.status or '').strip().lower()
        if 'replacement' in status or c.mail_sent_to_store:
            if c.settled_with_accounts:
                pass
            elif c.settlement_mail_accounts:
                pass
            elif c.invoice_sent_osg:
                if repl_age > 10:
                    res.append({
                        'Claim ID': c.claim_id,
                        'OSID': c.osid,
                        'Customer Name': c.customer_name,
                        'Mobile': c.mobile_no,
                        'Model': c.model,
                        'Status': c.status,
                        'Age (Days)': repl_age
                    })
    
    print(json.dumps(res, indent=2))
