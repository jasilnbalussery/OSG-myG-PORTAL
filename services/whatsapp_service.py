import re
import requests
import os
import logging
from dotenv import load_dotenv

load_dotenv()

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

API_URL = "https://hub.telinfy.com/unified/developer/api/v1/whatsapp/campaigns/send"
API_KEY = os.getenv("TELFINY_API_KEY")

# ─────────────────────────────────────────────────────────────────────────────
# TEMPLATE REGISTRY
# Copy exact body text from your Telfiny dashboard for each approved template.
# This is used to:
#   1. Count the {{N}} variables each template requires
#   2. Validate parameters BEFORE calling the API (prevents H31008 errors)
#
# To add a new template: paste its exact body text as a new entry below.
# ─────────────────────────────────────────────────────────────────────────────
TEMPLATE_REGISTRY = {
    "myg_onsitego_registered": (
        "പ്രിയ {{1}}NAME,\n"
        "നിങ്ങളുടെ {{2}}TV യുടെ EXTENDED വാറന്റി സർവീസ് രജിസ്റ്റർ ചെയ്തിട്ടുണ്ട്. SR നമ്പർ: {{3}}Complaint Number. കൂടുതൽ വിവരങ്ങൾക്കും സംശയങ്ങൾക്കും ഞങ്ങളുടെ കസ്റ്റമർ കെയർ നമ്പറായ 9249 001 001 ൽ ബന്ധപ്പെടാവുന്നതാണ്."
    ),
    "myg_onsitego_registered_main": (
        "പ്രിയ *{{1}}*,\n"
        "നിങ്ങളുടെ {{2}} യുടെ EXTENDED വാറൻ്റി സർവീസ് രജിസ്റ്റർ ചെയ്തിട്ടുണ്ട്. SR No: *{{3}}*. കൂടുതൽ വിവരങ്ങൾക്കും സംശയങ്ങൾക്കും ഞങ്ങളുടെ കസ്റ്റമർ കെയർ നമ്പറായ 9249 001 001 ൽ ബന്ധപ്പെടാവുന്നതാണ്."
    ),
    "osg_clm_registered": (
        "Hello {{1}}, Your service request for *{{2}}* under OSG Extended Warranty "
        "has been successfully registered. *Service ID: {{3}}* Our technician will "
        "contact you shortly. For any assistance, please contact our customer care "
        "at 9249001001. Regards, Service Coordinator, myG"
    ),
    "osg_clm_repair_completed": (
        "Hello {{1}}, We are happy to inform you that the service for *{{2}}* "
        "*(ID: {{3}})* has been completed and the case is now closed. We hope you "
        "are satisfied with our service. For further support, feel free to call us "
        "at 9249001001. Thank you for choosing myG! Best Regards, Service Coordinator, myG"
    ),
    "myg_onsitego_replacement_main": (
        "Dear *{{1}}*,\n"
        "നിങ്ങളുടെ *{{2}}* യുടെ വാറൻ്റി ക്ലെയിം APPROVE ആയിട്ടുണ്ട്. കൂടുതൽ വിവരങ്ങൾക്കായി ഞങ്ങളുടെ ഭാഗത്തു നിന്ന് താങ്കൾക്ക് CALL ലഭിക്കുന്നതാണ്. CUSTOMER CARE : 9249001001"
    ),
    "osg_clm_replacements_approved": (
        "Hi *{{1}}*! Good news: your extended warranty claim for a replacement has "
        "been approved. Our team will contact you shortly with the next steps and "
        "more information. Thank you for your patience! We hope you are satisfied "
        "with our service. For further support, feel free to call us at 9249001001. "
        "Thank you for choosing myG! Best Regards, Service Coordinator, myG"
    ),
    "osg_clm_reject": (
        "Hello *{{1}}*, regarding your extended warranty claim: unfortunately, it "
        "has been declined at this time. Our team will contact you shortly to provide "
        "more information and discuss the next steps. Thank you for your understanding. "
        "For further support, feel free to call us at 9249001001. Best Regards, "
        "Service Coordinator, myG"
    ),
    "myg_onsitego_repair_completed_main": (
        "Dear *{{1}}*,\n"
        "Your complaint (SR No: *{{2}}*) has been Completed..\n"
        "ഞങ്ങളുടെ സർവീസ് സംബന്ധിച്ച് നിങ്ങളുടെ വിലയേറിയ അഭിപ്രായങ്ങൾ 9249 001 001 എന്ന ഞങ്ങളുടെ കസ്റ്റമർ കെയർ നമ്പറിൽ അറിയിക്കാവുന്നതാണ്."
    ),
    "myg_onsitego_part_order_main": (
        "Dear *{{1}}*,\n"
        "നിങ്ങളുടെ *{{2}}* സർവീസിന് ആവശ്യമായ സ്പെയർ പാർട്ട് ഓർഡർ ചെയ്തിട്ടുണ്ട്. പാർട്ട് എത്തിയാൽ നിങ്ങളെ അറിയിക്കുകയും സർവീസ് പൂർത്തിയാക്കുകയും ചെയ്യുന്നതാണ്. കൂടുതൽ വിവരങ്ങൾക്ക് 9249 001 001 ൽ ബന്ധപ്പെടാവുന്നതാണ്."
    ),
}


def _count_template_vars(body_text: str) -> int:
    """Count the number of unique {{N}} placeholders in a template body string."""
    matches = re.findall(r'\{\{(\d+)\}\}', body_text)
    return len(set(matches))  # count unique variable positions (e.g. {{1}}, {{2}}, {{3}})


def validate_template_params(template_name: str, params: list):
    """
    Validate that the provided params match the template's requirements.

    Checks:
      1. Template exists in the registry.
      2. Number of params equals the number of {{N}} variables in the template.
      3. No individual param is empty, None, or whitespace-only.

    Returns:
      (True, "OK")                    — if all checks pass
      (False, "<reason string>")      — if any check fails
    """
    # 1. Template must be registered
    if template_name not in TEMPLATE_REGISTRY:
        return False, (
            f"Template '{template_name}' is not in the local registry. "
            f"Add its body text to TEMPLATE_REGISTRY in whatsapp_service.py."
        )

    body = TEMPLATE_REGISTRY[template_name]
    expected_count = _count_template_vars(body)
    provided_count = len(params)

    # 2. Param count must match exactly
    if provided_count != expected_count:
        return False, (
            f"Template '{template_name}' requires {expected_count} parameter(s) "
            f"but {provided_count} were provided. "
            f"Params given: {params}"
        )

    # 3. No individual param may be empty / None / "None" / whitespace
    for i, p in enumerate(params, start=1):
        val = str(p).strip() if p is not None else ""
        if val == "" or val.lower() == "none":
            return False, (
                f"Template '{template_name}': parameter {{{{i}}}} (position {i}) "
                f"is empty or None — sending would cause H31008. "
                f"All params: {params}"
            )

    return True, "OK"


def send_whatsapp_message(mobile: str, template_name: str, params: list) -> dict:
    """
    Sends a WhatsApp template message via the Telfiny API.

    Before sending, validates:
      - Template is registered and approved locally
      - Param count matches template variable count
      - No param is empty / None

    Args:
        mobile:        Recipient phone number (10-digit or with country code)
        template_name: Exact Telfiny template name (e.g. 'osg_clm_registered')
        params:        List of strings for {{1}}, {{2}}, ... in the template body

    Returns:
        dict with 'status_code' and 'response' on success,
        or 'error' and 'blocked': True if validation failed.
    """
    try:
        # ── CUTOFF DATE CHECK ─────────────────────────────────────────────────
        from datetime import date
        cutoff_date = date(2026, 7, 4)
        if date.today() <= cutoff_date:
            logger.info(f"[WHATSAPP_BLOCKED] Messages are disabled until after {cutoff_date.strftime('%d-%m-%Y')}.")
            return {"error": f"Blocked by cutoff date {cutoff_date.strftime('%d-%m-%Y')}", "blocked": True}

        # ── API KEY CHECK ─────────────────────────────────────────────────────
        if not API_KEY:
            logger.warning("[WHATSAPP] API Key is missing — check .env (TELFINY_API_KEY).")
            return {"error": "Missing API Key"}

        # ── VALIDATE PARAMS BEFORE SENDING ───────────────────────────────────
        is_valid, reason = validate_template_params(template_name, params)
        if not is_valid:
            logger.warning(f"[WHATSAPP_BLOCKED] Validation failed: {reason}")
            return {"error": reason, "blocked": True}

        # ── FORMAT MOBILE NUMBER ──────────────────────────────────────────────
        mobile_str = str(mobile).strip()
        if len(mobile_str) == 10 and mobile_str.isdigit():
            mobile_str = f"91{mobile_str}"

        # ── BUILD PAYLOAD ─────────────────────────────────────────────────────
        lang_code = "ml" if template_name in ["myg_onsitego_registered", "myg_onsitego_registered_main", "myg_onsitego_replacement_main", "myg_onsitego_repair_completed_main", "myg_onsitego_part_order_main"] else "en"
        
        # Inject Webhook URL (Trying multiple common keys since it's undocumented)
        webhook = os.getenv('WEBHOOK_URL', 'https://your-domain.com/api/webhooks/telfiny')
        
        payload = {
            "to": mobile_str,
            "phoneNumber": mobile_str,
            "type": "template",
            "webhookUrl": webhook,
            "webhook_url": webhook,
            "callbackUrl": webhook,
            "callback_url": webhook,
            "dlr_url": webhook,
            "template": {
                "name": template_name,
                "language": {"code": lang_code},
                "components": [
                    {
                        "type": "body",
                        "parameters": [
                            {"type": "text", "text": str(p)} for p in params
                        ]
                    }
                ]
            }
        }

        headers = {
            "Authorization": f"Bearer {API_KEY}",
            "x-api-key": API_KEY,
            "Content-Type": "application/json"
        }

        logger.info(
            f"[WHATSAPP] Sending → to={mobile_str} | "
            f"template='{template_name}' | params={params}"
        )

        # ── CALL API ──────────────────────────────────────────────────────────
        response = requests.post(API_URL, json=payload, headers=headers, timeout=10)

        logger.info(f"[WHATSAPP] Response status: {response.status_code}")
        try:
            json_response = response.json()
            logger.info(f"[WHATSAPP] Response body: {json_response}")
        except Exception:
            json_response = {"text_response": response.text}
            logger.info(f"[WHATSAPP] Response body (non-JSON): {response.text}")

        result = {
            "status_code": response.status_code,
            "response": json_response
        }

        # Surface the actual API error message when status is not 2xx
        if response.status_code not in [200, 201, 202]:
            api_err = (
                json_response.get("message")
                or json_response.get("error")
                or json_response.get("detail")
                or json_response.get("text_response")
                or f"HTTP {response.status_code}"
            )
            result["error"] = f"[{response.status_code}] {api_err}"
            logger.warning(f"[WHATSAPP] API error: {result['error']}")

        return result

    except Exception as e:
        logger.error(f"[WHATSAPP] Service error: {e}")
        return {"error": str(e)}
