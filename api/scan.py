"""Vercel serverless function: POST a receipt image -> structured line item.

Uses Claude vision (Anthropic API) to read a receipt photo and return JSON the
front-end can drop straight into a line row. Requires the ANTHROPIC_API_KEY
environment variable (set it in Vercel project settings). The model is
overridable via ANTHROPIC_MODEL.

Request  JSON: { image_base64, media_type, report_type }   ("job" | "office")
Response JSON: { configured, vendor, date, amount, category, confidence, raw }
"""
import json
import os
import re
import sys
from http.server import BaseHTTPRequestHandler

DEFAULT_MODEL = os.environ.get("ANTHROPIC_MODEL", "claude-sonnet-4-6")

JOB_CATEGORIES = ["airfare", "ground_transport", "car_rental_gas", "lodging",
                  "telephone_fax", "parking", "misc"]
OFFICE_CATEGORIES = ["office_expense", "credit_card"]


def _prompt(report_type: str) -> str:
    cats = JOB_CATEGORIES if report_type != "office" else OFFICE_CATEGORIES
    return (
        "You are reading a single expense receipt for an expense report. "
        "Extract the following and respond with STRICT JSON only (no prose, no "
        "markdown fences):\n"
        '{"vendor": string, "date": "YYYY-MM-DD" or null, '
        '"amount": number (grand total incl. tax/tip), '
        f'"category": one of {cats}, '
        '"confidence": number between 0 and 1}\n'
        "Pick the single best category. If unsure about any field, still return "
        "your best guess and lower the confidence. amount must be a plain number "
        "with no currency symbol."
    )


def _extract_json(text: str) -> dict:
    text = text.strip()
    text = re.sub(r"^```(?:json)?|```$", "", text, flags=re.MULTILINE).strip()
    m = re.search(r"\{.*\}", text, flags=re.DOTALL)
    if not m:
        raise ValueError(f"Model did not return JSON: {text[:200]}")
    return json.loads(m.group(0))


class handler(BaseHTTPRequestHandler):
    def do_POST(self):
        try:
            length = int(self.headers.get("content-length") or 0)
            data = json.loads(self.rfile.read(length) if length else b"{}")
            image_b64 = data.get("image_base64")
            media_type = data.get("media_type", "image/jpeg")
            report_type = (data.get("report_type") or "job").lower()
            if not image_b64:
                raise ValueError("image_base64 is required")

            api_key = os.environ.get("ANTHROPIC_API_KEY")
            if not api_key:
                return self._json({
                    "configured": False,
                    "message": "Receipt scanning is off. Set ANTHROPIC_API_KEY in "
                               "your Vercel project env vars to enable it.",
                }, 200)

            import anthropic  # imported lazily so generate.py never needs it
            client = anthropic.Anthropic(api_key=api_key)
            if media_type == "application/pdf":
                receipt_block = {"type": "document", "source": {
                    "type": "base64", "media_type": "application/pdf", "data": image_b64}}
            else:
                receipt_block = {"type": "image", "source": {
                    "type": "base64", "media_type": media_type, "data": image_b64}}
            resp = client.messages.create(
                model=DEFAULT_MODEL,
                max_tokens=512,
                messages=[{
                    "role": "user",
                    "content": [receipt_block, {"type": "text", "text": _prompt(report_type)}],
                }],
            )
            text = "".join(b.text for b in resp.content if getattr(b, "type", "") == "text")
            parsed = _extract_json(text)
            parsed["configured"] = True
            parsed.setdefault("confidence", 0.5)
            self._json(parsed, 200)
        except Exception as e:  # noqa: BLE001
            self._json({"configured": True, "error": str(e)}, 400)

    def _json(self, obj: dict, code: int):
        payload = json.dumps(obj).encode()
        self.send_response(code)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(payload)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(payload)
