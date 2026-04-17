#!/usr/bin/env python3
from __future__ import annotations

import argparse
import html
import json
import os
from email.parser import BytesParser
from email.policy import default
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse

from generate_report import build_report_html


MAX_UPLOAD_BYTES = 25 * 1024 * 1024


def parse_uploaded_spreadsheet(headers, body: bytes) -> tuple[str, bytes] | tuple[None, None]:
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
    code {{
      background: #eef4ff;
      padding: 2px 6px;
      border-radius: 8px;
    }}
    @media (max-width: 720px) {{
      .page {{
        width: min(100vw - 18px, 980px);
      }}
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
      <h1>Weekly Report Automation</h1>
      <p>Upload an <code>.xlsx</code> file and the system generates the dashboard immediately.</p>
      <p>No uploaded spreadsheet is stored on the server after processing.</p>
      <form class="upload-form" method="post" action="/upload" enctype="multipart/form-data">
        <div class="upload-row">
          <input class="upload-input" type="file" name="spreadsheet" accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" required>
          <button class="upload-button" type="submit">Upload And Open Report</button>
        </div>
        <div class="upload-help">The file is processed in memory only, so no Excel data is kept on the server.</div>
      </form>
      {alert_html}
    </section>

    <section class="section">
      <h2>How to use</h2>
      <p>1. Start the server with <code>python3 server.py</code>.</p>
      <p>2. Open <code>http://127.0.0.1:8000</code> in the browser.</p>
      <p>3. Upload an <code>.xlsx</code> file and the report will open directly.</p>
    </section>
  </div>
</body>
</html>"""


class ReportHandler(BaseHTTPRequestHandler):
    server_version = "LocalReportServer/3.0"

    def _send_text(self, status: int, body: str, content_type: str = "text/html; charset=utf-8") -> None:
        encoded = body.encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(encoded)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(encoded)

    def _send_json(self, payload: object, status: int = HTTPStatus.OK) -> None:
        self._send_text(status, json.dumps(payload, ensure_ascii=False, indent=2), "application/json; charset=utf-8")

    def do_GET(self) -> None:
        parsed = urlparse(self.path)

        if parsed.path == "/":
            self._send_text(HTTPStatus.OK, render_home_page())
            return

        if parsed.path == "/health":
            self._send_json({"status": "ok"})
            return

        self._send_text(HTTPStatus.NOT_FOUND, "<h1>404</h1><p>Page not found.</p>")

    def do_POST(self) -> None:
        parsed = urlparse(self.path)

        if parsed.path != "/upload":
            self._send_text(HTTPStatus.NOT_FOUND, "<h1>404</h1><p>Page not found.</p>")
            return

        try:
            content_length = int(self.headers.get("Content-Length", "0"))
        except ValueError:
            content_length = 0

        if content_length <= 0:
            self._send_text(HTTPStatus.BAD_REQUEST, render_home_page("No file was sent.", True))
            return

        if content_length > MAX_UPLOAD_BYTES:
            self._send_text(
                HTTPStatus.REQUEST_ENTITY_TOO_LARGE,
                render_home_page("The uploaded file is too large. Please keep it under 25 MB.", True),
            )
            return

        body = self.rfile.read(content_length)
        filename, payload = parse_uploaded_spreadsheet(self.headers, body)

        if not filename or not payload:
            self._send_text(
                HTTPStatus.BAD_REQUEST,
                render_home_page("Please choose a valid .xlsx file before uploading.", True),
            )
            return

        if not filename.lower().endswith(".xlsx"):
            self._send_text(
                HTTPStatus.BAD_REQUEST,
                render_home_page("Only .xlsx files are supported.", True),
            )
            return

        try:
            html_report = build_report_html(filename, payload)
        except Exception as exc:  # noqa: BLE001
            self._send_text(
                HTTPStatus.INTERNAL_SERVER_ERROR,
                render_home_page(f"Could not process the spreadsheet: {exc}", True),
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
