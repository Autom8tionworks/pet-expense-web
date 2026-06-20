"""Vercel serverless function: POST submission JSON -> download filled .xlsx.

No email, no login. The browser posts the form as JSON; we fill the correct
master template in-memory and stream the workbook back as a file download.
"""
import json
import os
import sys
from http.server import BaseHTTPRequestHandler

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _lib.build import generate_xlsx  # noqa: E402

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


class handler(BaseHTTPRequestHandler):
    def do_POST(self):
        try:
            length = int(self.headers.get("content-length") or 0)
            raw = self.rfile.read(length) if length else b"{}"
            data = json.loads(raw or b"{}")
            filename, xlsx = generate_xlsx(data)

            self.send_response(200)
            self.send_header("Content-Type", XLSX_MIME)
            self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
            self.send_header("Content-Length", str(len(xlsx)))
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            self.wfile.write(xlsx)
        except Exception as e:  # noqa: BLE001
            self._error(str(e))

    def _error(self, message: str, code: int = 400):
        payload = json.dumps({"error": message}).encode()
        self.send_response(code)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(payload)))
        self.end_headers()
        self.wfile.write(payload)
