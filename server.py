#!/usr/bin/env python3
from __future__ import annotations

import argparse
import html
import json
import os
import re
import tempfile
from datetime import datetime
from email.parser import BytesParser
from email.policy import default
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, quote, urlparse

from generate_report import build_report


ROOT_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = ROOT_DIR / "report_output"
OUTPUT_FILE = OUTPUT_DIR / "report.html"
UPLOADS_DIR = Path(tempfile.gettempdir()) / "clinica_report_uploads"
MAX_UPLOAD_BYTES = 25 * 1024 * 1024


def list_spreadsheets() -> list[Path]:
    if not UPLOADS_DIR.exists():
        return []
    return sorted(
        [path for path in UPLOADS_DIR.glob("*.xlsx") if not path.name.startswith("~$")],
        key=lambda path: path.stat().st_mtime,
        reverse=True,
    )


def resolve_spreadsheet(filename: str | None) -> Path | None:
    spreadsheets = list_spreadsheets()
    if not spreadsheets:
        return None
    if not filename:
        return spreadsheets[0]
    for path in spreadsheets:
        if path.name == filename:
            return path
    return None


def sanitize_filename(filename: str) -> str:
    cleaned = Path(filename).name.strip()
    stem = re.sub(r"[^A-Za-z0-9._-]+", "-", Path(cleaned).stem).strip("._-")
    suffix = Path(cleaned).suffix.lower()
    if suffix != ".xlsx":
      suffix = ".xlsx"
    if not stem:
        stem = "report"
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    return f"{timestamp}-{stem}{suffix}"


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


def save_uploaded_spreadsheet(filename: str, content: bytes) -> Path:
    UPLOADS_DIR.mkdir(parents=True, exist_ok=True)
    output_name = sanitize_filename(filename)
    output_path = UPLOADS_DIR / output_name
    output_path.write_bytes(content)
    return output_path


def render_home_page(selected_name: str | None = None, message: str = "", is_error: bool = False) -> str:
    spreadsheets = list_spreadsheets()
    cards = []
    for spreadsheet in spreadsheets[:8]:
        is_selected = spreadsheet.name == selected_name
        open_href = "/dashboard?file=" + quote(spreadsheet.name)
        modified_at = datetime.fromtimestamp(spreadsheet.stat().st_mtime).strftime("%Y-%m-%d %H:%M")
        cards.append(
            f"""
            <a class="file-card{' is-selected' if is_selected else ''}" href="{open_href}">
              <div class="file-name">{html.escape(spreadsheet.name)}</div>
              <div class="file-meta">Uploaded at {html.escape(modified_at)}</div>
            </a>
            """
        )

    cards_html = "".join(cards) if cards else '<div class="empty">No spreadsheet uploaded yet. Use the upload area above to generate your first report.</div>'
    latest_dashboard_href = "/dashboard?file=" + quote(spreadsheets[0].name) if spreadsheets else "#"
    latest_button_class = "button" if spreadsheets else "button is-disabled"
    alert_html = ""
    if message:
        alert_html = f'<div class="alert{" is-error" if is_error else ""}">{html.escape(message)}</div>'

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Report Automation</title>
  <style>
    :root {{
      --bg: #f3f7fc;
      --panel: #ffffff;
      --text: #102033;
      --muted: #5b6b7a;
      --line: #d8e2ef;
      --accent: #1264d6;
      --accent-dark: #102f6b;
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
      width: min(1100px, calc(100vw - 28px));
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
    .actions {{
      display: flex;
      gap: 12px;
      flex-wrap: wrap;
      margin-top: 18px;
    }}
    .button,
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
    .button.secondary {{
      background: white;
      color: var(--accent);
      border-color: #bfdbfe;
    }}
    .button.is-disabled {{
      pointer-events: none;
      opacity: 0.55;
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
    .upload-help {{
      font-size: 14px;
      color: rgba(255, 255, 255, 0.82);
    }}
    .grid {{
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
      gap: 16px;
      margin-top: 18px;
    }}
    .file-card {{
      display: block;
      padding: 18px;
      border-radius: 20px;
      background: linear-gradient(180deg, #ffffff, #f8fbff);
      border: 1px solid var(--line);
      color: inherit;
      text-decoration: none;
      transition: transform 140ms ease, border-color 140ms ease;
    }}
    .file-card:hover,
    .file-card.is-selected {{
      transform: translateY(-1px);
      border-color: #8bb5f8;
    }}
    .file-name {{
      font-weight: 700;
      margin-bottom: 8px;
      word-break: break-word;
    }}
    .file-meta {{
      color: var(--muted);
      font-size: 14px;
    }}
    .empty {{
      padding: 18px;
      border-radius: 18px;
      border: 1px dashed var(--line);
      color: var(--muted);
      background: #fbfdff;
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
        width: min(100vw - 18px, 1100px);
      }}
      .upload-row {{
        flex-direction: column;
        align-items: stretch;
      }}
      .upload-button,
      .button {{
        width: 100%;
      }}
    }}
  </style>
</head>
<body>
  <div class="page">
    <section class="hero">
      <h1>Report Automation</h1>
      <p>This version is ready for publishing because the application no longer depends on Excel files living inside the project folder.</p>
      <p>Upload an <code>.xlsx</code> file, generate the dashboard, and open the full interactive report directly in the browser.</p>
      <form class="upload-form" method="post" action="/upload" enctype="multipart/form-data">
        <div class="upload-row">
          <input class="upload-input" type="file" name="spreadsheet" accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" required>
          <button class="upload-button" type="submit">Upload And Open Report</button>
        </div>
        <div class="upload-help">Uploaded spreadsheets are stored outside the project folder, so your published code stays clean.</div>
      </form>
      {alert_html}
      <div class="actions">
        <a class="{latest_button_class}" href="{latest_dashboard_href}">Open latest uploaded report</a>
        <a class="button secondary" href="/api/files">View uploaded files</a>
      </div>
    </section>

    <section class="section">
      <h2>Recent uploads</h2>
      <p>Each upload generates the same interactive dashboard without keeping the Excel workbook inside the project repository.</p>
      <div class="grid">{cards_html}</div>
    </section>

    <section class="section">
      <h2>How to use</h2>
      <p>1. Start the server with <code>python3 server.py</code>.</p>
      <p>2. Open <code>http://127.0.0.1:8000</code> in the browser.</p>
      <p>3. Upload an <code>.xlsx</code> file and the app will generate the dashboard automatically.</p>
    </section>
  </div>
</body>
</html>"""


class ReportHandler(BaseHTTPRequestHandler):
    server_version = "LocalReportServer/2.0"

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

    def _redirect(self, location: str) -> None:
        self.send_response(HTTPStatus.SEE_OTHER)
        self.send_header("Location", location)
        self.send_header("Cache-Control", "no-store")
        self.end_headers()

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        query = parse_qs(parsed.query)

        if parsed.path == "/":
            selected = query.get("file", [None])[0]
            message = query.get("message", [""])[0]
            is_error = query.get("error", ["0"])[0] == "1"
            self._send_text(HTTPStatus.OK, render_home_page(selected, message, is_error))
            return

        if parsed.path == "/api/files":
            files = list_spreadsheets()
            self._send_json(
                {
                    "files": [path.name for path in files],
                    "default": files[0].name if files else None,
                    "storage": str(UPLOADS_DIR),
                }
            )
            return

        if parsed.path == "/health":
            self._send_json({"status": "ok"})
            return

        if parsed.path == "/dashboard":
            selected_name = query.get("file", [None])[0]
            spreadsheet = resolve_spreadsheet(selected_name)
            if spreadsheet is None:
                self._send_text(
                    HTTPStatus.NOT_FOUND,
                    render_home_page(
                        selected_name,
                        "Upload a spreadsheet first to generate the report.",
                        True,
                    ),
                )
                return

            OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
            build_report(spreadsheet, OUTPUT_FILE)
            self._send_text(HTTPStatus.OK, OUTPUT_FILE.read_text(encoding="utf-8"))
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
            self._send_text(HTTPStatus.BAD_REQUEST, render_home_page(message="No file was sent.", is_error=True))
            return

        if content_length > MAX_UPLOAD_BYTES:
            self._send_text(
                HTTPStatus.REQUEST_ENTITY_TOO_LARGE,
                render_home_page(message="The uploaded file is too large. Please keep it under 25 MB.", is_error=True),
            )
            return

        body = self.rfile.read(content_length)
        filename, payload = parse_uploaded_spreadsheet(self.headers, body)

        if not filename or not payload:
            self._send_text(
                HTTPStatus.BAD_REQUEST,
                render_home_page(message="Please choose a valid .xlsx file before uploading.", is_error=True),
            )
            return

        if Path(filename).suffix.lower() != ".xlsx":
            self._send_text(
                HTTPStatus.BAD_REQUEST,
                render_home_page(message="Only .xlsx files are supported.", is_error=True),
            )
            return

        try:
            saved_file = save_uploaded_spreadsheet(filename, payload)
            OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
            build_report(saved_file, OUTPUT_FILE)
        except Exception as exc:  # noqa: BLE001
            self._send_text(
                HTTPStatus.INTERNAL_SERVER_ERROR,
                render_home_page(message=f"Could not process the spreadsheet: {exc}", is_error=True),
            )
            return

        self._redirect("/dashboard?file=" + quote(saved_file.name))


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
    UPLOADS_DIR.mkdir(parents=True, exist_ok=True)
    args = parse_args()
    server = ThreadingHTTPServer((args.host, args.port), ReportHandler)
    print(f"Local report server running at http://{args.host}:{args.port}")
    print(f"Uploaded spreadsheets are stored at {UPLOADS_DIR}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
