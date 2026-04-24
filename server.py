#!/usr/bin/env python3
from __future__ import annotations

import argparse
import base64
import html
import hmac
import json
import ipaddress
import os
from email.parser import BytesParser
from email.policy import default
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse

from generate_report import (
    build_csv_report_html,
    build_empty_report_html,
    build_report_html,
)


MAX_UPLOAD_BYTES = 25 * 1024 * 1024
BASIC_AUTH_USER = os.getenv("BASIC_AUTH_USER", "")
BASIC_AUTH_PASSWORD = os.getenv("BASIC_AUTH_PASSWORD", "")
ALLOWED_IPS = [item.strip() for item in os.getenv("ALLOWED_IPS", "").split(",") if item.strip()]

def parse_uploaded_file(headers, body: bytes) -> tuple[str, bytes] | tuple[None, None]:
    content_type = headers.get("Content-Type", "")
    if "multipart/form-data" not in content_type:
        return None, None

    message = BytesParser(policy=default).parsebytes(
        (
            f"Content-Type: {content_type}\r\n"
            "MIME-Version: 1.0\r\n\r\n"
        ).encode("utf-8") + body
    )

    for part in message.iter_parts():
        if part.get_content_disposition() != "form-data":
            continue
        if part.get_param("name", header="content-disposition") != "spreadsheet":
            continue
        filename = part.get_filename()
        payload = part.get_payload(decode=True) or b""
        if not filename:
            return None, None
        return filename, payload
    return None, None


def render_home_page(message: str = "", is_error: bool = False) -> str:
    alert_html = ""
    if message:
        alert_html = f'<div class="alert{" is-error" if is_error else ""}">{html.escape(message)}</div>'

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Weekly Report Automation</title>
  <style>
    :root {{
      --bg: #f3f7fc;
      --panel: #ffffff;
      --text: #102033;
      --muted: #5b6b7a;
      --line: #d8e2ef;
      --accent: #1264d6;
      --shadow: 0 24px 48px rgba(15, 23, 42, 0.10);
      --success: #e8f7ee;
      --danger: #fdecec;
      --danger-text: #9b1c1c;
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      font-family: "Avenir Next", "Segoe UI", sans-serif;
      color: var(--text);
      background:
        radial-gradient(circle at top left, rgba(18, 100, 214, 0.16), transparent 24%),
        linear-gradient(180deg, #eef4fb 0%, var(--bg) 24%, var(--bg) 100%);
    }}
    .page {{
      width: min(980px, calc(100vw - 28px));
      margin: 0 auto;
      padding: 28px 0 40px;
    }}
    .hero {{
      padding: 28px;
      border-radius: 28px;
      background: linear-gradient(135deg, rgba(16, 32, 51, 0.96), rgba(18, 100, 214, 0.94));
      color: white;
      box-shadow: var(--shadow);
    }}
    .hero-top {{
      display: flex;
      justify-content: space-between;
      gap: 16px;
      align-items: flex-start;
      flex-wrap: wrap;
    }}
    h1 {{
      margin: 0 0 12px;
      font-size: clamp(28px, 4vw, 42px);
      letter-spacing: -0.03em;
    }}
    h2 {{
      margin: 0 0 12px;
      font-size: 24px;
      letter-spacing: -0.02em;
    }}
    p {{
      margin: 8px 0;
      line-height: 1.55;
    }}
    .section {{
      margin-top: 22px;
      padding: 22px;
      border-radius: 24px;
      background: rgba(255, 255, 255, 0.92);
      border: 1px solid rgba(216, 226, 239, 0.92);
      box-shadow: var(--shadow);
    }}
    .upload-form {{
      display: grid;
      gap: 14px;
      margin-top: 18px;
    }}
    .upload-row {{
      display: flex;
      gap: 12px;
      flex-wrap: wrap;
      align-items: center;
    }}
    .upload-input {{
      flex: 1 1 320px;
      min-width: 280px;
      border: 1px dashed rgba(255, 255, 255, 0.45);
      border-radius: 16px;
      padding: 14px;
      background: rgba(255, 255, 255, 0.10);
      color: white;
      font: inherit;
    }}
    .upload-button {{
      display: inline-flex;
      align-items: center;
      justify-content: center;
      padding: 12px 16px;
      border-radius: 14px;
      border: 1px solid rgba(255, 255, 255, 0.18);
      background: rgba(255, 255, 255, 0.14);
      color: white;
      text-decoration: none;
      font-weight: 700;
      font: inherit;
      cursor: pointer;
    }}
    .upload-help {{
      font-size: 14px;
      color: rgba(255, 255, 255, 0.82);
    }}
    .theme-toggle-wrap {{
      display: grid;
      gap: 10px;
      min-width: 210px;
    }}
    .theme-toggle-row {{
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 12px;
    }}
    .theme-toggle-label {{
      font-size: 12px;
      letter-spacing: 0.08em;
      text-transform: uppercase;
      color: rgba(255, 255, 255, 0.82);
      font-weight: 700;
    }}
    .theme-toggle-state {{
      font-size: 13px;
      color: rgba(255, 255, 255, 0.88);
    }}
    .theme-toggle {{
      appearance: none;
      border: 1px solid rgba(255, 255, 255, 0.20);
      background: rgba(255, 255, 255, 0.14);
      color: white;
      border-radius: 999px;
      padding: 5px;
      width: 100%;
      cursor: pointer;
      display: flex;
      align-items: center;
      justify-content: flex-start;
    }}
    .theme-toggle-thumb {{
      width: 50%;
      min-width: 90px;
      padding: 10px 12px;
      border-radius: 999px;
      background: rgba(255, 255, 255, 0.92);
      color: #102033;
      font-size: 13px;
      font-weight: 800;
      text-align: center;
      transition: transform 180ms ease, background 180ms ease, color 180ms ease;
    }}
    .alert {{
      margin-top: 16px;
      padding: 14px 16px;
      border-radius: 16px;
      background: var(--success);
      border: 1px solid #bfe3cb;
      color: #14532d;
      font-weight: 600;
    }}
    .alert.is-error {{
      background: var(--danger);
      border-color: #f6caca;
      color: var(--danger-text);
    }}
    .home-actions {{
      display: flex;
      gap: 12px;
      flex-wrap: wrap;
      margin-top: 16px;
    }}
    .secondary-button {{
      display: inline-flex;
      align-items: center;
      justify-content: center;
      padding: 12px 16px;
      border-radius: 14px;
      border: 1px solid rgba(255, 255, 255, 0.22);
      background: rgba(255, 255, 255, 0.08);
      color: white;
      text-decoration: none;
      font-weight: 700;
      font: inherit;
      cursor: pointer;
    }}
    code {{
      background: #eef4ff;
      padding: 2px 6px;
      border-radius: 8px;
    }}
    body.theme-corona {{
      --bg: #0f1015;
      --panel: #191c24;
      --text: #f5f5f5;
      --muted: #a1aab8;
      --line: #2c2e33;
      --accent: #0090e7;
      --shadow: 0 18px 34px rgba(0, 0, 0, 0.28);
      background:
        radial-gradient(circle at 0% 0%, rgba(0, 144, 231, 0.16), transparent 22%),
        radial-gradient(circle at 100% 0%, rgba(0, 210, 91, 0.12), transparent 20%),
        linear-gradient(180deg, #0a0b0f 0%, #0f1015 100%);
    }}
    body.theme-corona .hero,
    body.theme-corona .section {{
      background: #191c24;
      color: #f5f5f5;
      border-color: #2c2e33;
      box-shadow: none;
    }}
    body.theme-corona .hero {{
      background: linear-gradient(135deg, #191c24, #111318 58%, #151922);
    }}
    body.theme-corona .upload-input,
    body.theme-corona .theme-toggle {{
      background: #0f1015;
      color: #f5f5f5;
      border-color: #2c2e33;
    }}
    body.theme-corona .theme-toggle-thumb {{
      transform: translateX(100%);
      background: linear-gradient(135deg, #0090e7, #0069aa);
      color: white;
    }}
    body.theme-corona .upload-button {{
      background: linear-gradient(135deg, #0090e7, #0069aa);
      border-color: transparent;
    }}
    body.theme-corona .secondary-button {{
      background: rgba(255, 255, 255, 0.06);
      border-color: #2c2e33;
    }}
    body.theme-corona code {{
      background: #0f1015;
      color: #8fd8ff;
    }}
    @media (max-width: 720px) {{
      .page {{
        width: min(100vw - 18px, 980px);
      }}
      .hero-top,
      .upload-row {{
        flex-direction: column;
        align-items: stretch;
      }}
      .upload-button {{
        width: 100%;
      }}
    }}
  </style>
</head>
<body>
  <div class="page">
    <section class="hero">
      <div class="hero-top">
        <div>
          <h1>Weekly Report Automation</h1>
          <p>Upload an <code>.xlsx</code>, <code>Intervention Log .csv</code>, or <code>Flat Time .csv</code> and the system generates the dashboard immediately.</p>
          <p>You can also open the dashboard without a workbook and use the Flat Time tab with CSV files only.</p>
        </div>
        <div class="theme-toggle-wrap">
          <div class="theme-toggle-row">
            <span class="theme-toggle-label">Theme</span>
            <span id="theme-toggle-state" class="theme-toggle-state">Classic</span>
          </div>
          <button id="theme-toggle" class="theme-toggle" type="button" aria-label="Toggle theme">
            <span id="theme-toggle-thumb" class="theme-toggle-thumb">Classic</span>
          </button>
        </div>
      </div>
      <form class="upload-form" method="post" action="/upload" enctype="multipart/form-data">
        <div class="upload-row">
          <input class="upload-input" type="file" name="spreadsheet" accept=".xlsx,.csv,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,text/csv" required>
          <button class="upload-button" type="submit">Upload And Open Report</button>
        </div>
        <div class="upload-help">The file is processed in memory only, so no Excel data is kept on the server.</div>
      </form>
      <div class="home-actions">
        <a class="secondary-button" href="/dashboard">Open Dashboard Without Workbook</a>
      </div>
      {alert_html}
    </section>

    <section class="section">
      <h2>How to use</h2>
      <p>1. <code>upload a .xlsx file</code>, <code>Intervention Log .csv</code>, or <code>Flat Time .csv</code>.</p>
      <p>2. or open <code>the dashboard without a workbook</code> and use <code>Flat Time</code> with CSVs.</p>
      <p>3. Review the dashboard in the browser.</p>
      <p>4. Uploaded files are not stored after processing.</p>
    </section>
  </div>
  <script>
    const THEME_STORAGE_KEY = "weekly-report-theme";
    const themeToggle = document.getElementById("theme-toggle");
    const themeToggleState = document.getElementById("theme-toggle-state");
    const themeToggleThumb = document.getElementById("theme-toggle-thumb");

    function applyTheme(theme) {{
      const resolvedTheme = theme === "corona" ? "corona" : "classic";
      document.body.classList.toggle("theme-corona", resolvedTheme === "corona");
      themeToggleState.textContent = resolvedTheme === "corona" ? "Corona" : "Classic";
      themeToggleThumb.textContent = resolvedTheme === "corona" ? "Corona" : "Classic";
      localStorage.setItem(THEME_STORAGE_KEY, resolvedTheme);
    }}

    themeToggle.addEventListener("click", () => {{
      applyTheme(document.body.classList.contains("theme-corona") ? "classic" : "corona");
    }});

    applyTheme(localStorage.getItem(THEME_STORAGE_KEY) || "classic");
  </script>
</body>
</html>"""


def render_message_page(title: str, message: str, status_label: str = "Message") -> str:
    return f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{html.escape(title)}</title>
  <style>
    body {{
      margin: 0;
      font-family: "Avenir Next", "Segoe UI", sans-serif;
      color: #102033;
      background: linear-gradient(180deg, #eef4fb 0%, #f3f7fc 100%);
    }}
    .page {{
      width: min(760px, calc(100vw - 28px));
      margin: 0 auto;
      padding: 40px 0;
    }}
    .panel {{
      background: rgba(255, 255, 255, 0.94);
      border: 1px solid rgba(216, 226, 239, 0.92);
      border-radius: 24px;
      box-shadow: 0 24px 48px rgba(15, 23, 42, 0.10);
      padding: 28px;
    }}
    .eyebrow {{
      display: inline-flex;
      padding: 6px 10px;
      border-radius: 999px;
      background: #eef4ff;
      color: #1264d6;
      font-size: 12px;
      font-weight: 800;
      letter-spacing: 0.08em;
      text-transform: uppercase;
    }}
    h1 {{
      margin: 14px 0 10px;
      font-size: clamp(28px, 4vw, 38px);
      letter-spacing: -0.03em;
    }}
    p {{
      margin: 8px 0;
      line-height: 1.55;
      color: #546579;
    }}
    a {{
      display: inline-flex;
      align-items: center;
      justify-content: center;
      margin-top: 16px;
      padding: 12px 16px;
      border-radius: 14px;
      background: #1264d6;
      color: white;
      text-decoration: none;
      font-weight: 700;
    }}
  </style>
</head>
<body>
  <div class="page">
    <div class="panel">
      <div class="eyebrow">{html.escape(status_label)}</div>
      <h1>{html.escape(title)}</h1>
      <p>{html.escape(message)}</p>
      <a href="/">Return to dashboard</a>
    </div>
  </div>
</body>
</html>"""


class ReportHandler(BaseHTTPRequestHandler):
    server_version = "LocalReportServer/3.0"

    def _client_ip(self) -> str:
        forwarded_for = self.headers.get("X-Forwarded-For", "").split(",")[0].strip()
        return forwarded_for or self.client_address[0]

    def _request_is_https(self) -> bool:
        return self.headers.get("X-Forwarded-Proto", "").lower() == "https"

    def _is_ip_allowed(self) -> bool:
        if not ALLOWED_IPS:
            return True
        client_ip = ipaddress.ip_address(self._client_ip())
        for value in ALLOWED_IPS:
            try:
                if "/" in value:
                    if client_ip in ipaddress.ip_network(value, strict=False):
                        return True
                elif client_ip == ipaddress.ip_address(value):
                    return True
            except ValueError:
                continue
        return False

    def _is_authenticated(self) -> bool:
        if not BASIC_AUTH_USER or not BASIC_AUTH_PASSWORD:
            return True
        header = self.headers.get("Authorization", "")
        if not header.startswith("Basic "):
            return False
        try:
            decoded = base64.b64decode(header.split(" ", 1)[1]).decode("utf-8")
        except Exception:  # noqa: BLE001
            return False
        username, separator, password = decoded.partition(":")
        return (
            separator == ":"
            and hmac.compare_digest(username, BASIC_AUTH_USER)
            and hmac.compare_digest(password, BASIC_AUTH_PASSWORD)
        )

    def _send_auth_challenge(self) -> None:
        self.send_response(HTTPStatus.UNAUTHORIZED)
        self.send_header("WWW-Authenticate", 'Basic realm="Weekly Report"')
        self.send_header("Cache-Control", "no-store")
        self.send_header("Content-Length", "0")
        self._send_security_headers()
        self.end_headers()

    def _send_security_headers(self) -> None:
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("X-Frame-Options", "DENY")
        self.send_header("Referrer-Policy", "no-referrer")
        self.send_header("Permissions-Policy", "camera=(), microphone=(), geolocation=()")
        self.send_header(
            "Content-Security-Policy",
            "default-src 'self'; img-src 'self' data:; style-src 'self' 'unsafe-inline'; "
            "script-src 'self' 'unsafe-inline'; connect-src 'self'; font-src 'self' data:; "
            "object-src 'none'; base-uri 'self'; frame-ancestors 'none'; form-action 'self'",
        )
        if self._request_is_https():
            self.send_header("Strict-Transport-Security", "max-age=31536000; includeSubDomains")

    def _send_text(self, status: int, body: str, content_type: str = "text/html; charset=utf-8") -> None:
        encoded = body.encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(encoded)))
        self.send_header("Cache-Control", "no-store")
        self._send_security_headers()
        self.end_headers()
        self.wfile.write(encoded)

    def _send_json(self, payload: object, status: int = HTTPStatus.OK) -> None:
        self._send_text(status, json.dumps(payload, ensure_ascii=False, indent=2), "application/json; charset=utf-8")

    def do_GET(self) -> None:
        parsed = urlparse(self.path)

        if parsed.path != "/health" and not self._is_ip_allowed():
            self._send_text(HTTPStatus.FORBIDDEN, "<h1>403</h1><p>Access denied.</p>")
            return

        if parsed.path != "/health" and not self._is_authenticated():
            self._send_auth_challenge()
            return

        if parsed.path in {"/", "/dashboard"}:
            self._send_text(
                HTTPStatus.OK,
                build_empty_report_html(flat_time_payload={"datasets": []}),
            )
            return

        if parsed.path == "/health":
            self._send_json({"status": "ok"})
            return

        self._send_text(HTTPStatus.NOT_FOUND, "<h1>404</h1><p>Page not found.</p>")

    def do_POST(self) -> None:
        parsed = urlparse(self.path)

        if not self._is_ip_allowed():
            self._send_text(HTTPStatus.FORBIDDEN, "<h1>403</h1><p>Access denied.</p>")
            return

        if not self._is_authenticated():
            self._send_auth_challenge()
            return

        if parsed.path != "/upload":
            self._send_text(HTTPStatus.NOT_FOUND, "<h1>404</h1><p>Page not found.</p>")
            return

        try:
            content_length = int(self.headers.get("Content-Length", "0"))
        except ValueError:
            content_length = 0

        if content_length <= 0:
            self._send_text(HTTPStatus.BAD_REQUEST, render_message_page("Upload failed", "No file was sent.", "Error"))
            return

        if content_length > MAX_UPLOAD_BYTES:
            self._send_text(
                HTTPStatus.REQUEST_ENTITY_TOO_LARGE,
                render_message_page("Upload failed", "The uploaded file is too large. Please keep it under 25 MB.", "Error"),
            )
            return

        body = self.rfile.read(content_length)
        filename, payload = parse_uploaded_file(self.headers, body)

        if not filename or not payload:
            self._send_text(
                HTTPStatus.BAD_REQUEST,
                render_message_page("Upload failed", "Please choose a valid .xlsx or .csv file before uploading.", "Error"),
            )
            return

        lower_name = filename.lower()

        try:
            if lower_name.endswith(".xlsx"):
                html_report = build_report_html(
                    filename,
                    payload,
                    flat_time_payload={"datasets": []},
                )
            elif lower_name.endswith(".csv"):
                html_report = build_csv_report_html(
                    filename,
                    payload,
                    flat_time_payload={"datasets": []},
                )
            else:
                self._send_text(
                    HTTPStatus.BAD_REQUEST,
                    render_message_page("Upload failed", "Only .xlsx and .csv files are supported.", "Error"),
                )
                return
        except Exception as exc:  # noqa: BLE001
            self._send_text(
                HTTPStatus.INTERNAL_SERVER_ERROR,
                render_message_page("Upload failed", f"Could not process the uploaded file: {exc}", "Error"),
            )
            return

        self._send_text(HTTPStatus.OK, html_report)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Run the local interactive report server.")
    parser.add_argument(
        "--host",
        default=os.getenv("HOST", "0.0.0.0"),
        help="Host to bind the server to. Default: HOST env or 0.0.0.0",
    )
    parser.add_argument(
        "--port",
        type=int,
        default=int(os.getenv("PORT", "8000")),
        help="Port to bind the server to. Default: PORT env or 8000",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    server = ThreadingHTTPServer((args.host, args.port), ReportHandler)
    print(f"Local report server running at http://{args.host}:{args.port}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
