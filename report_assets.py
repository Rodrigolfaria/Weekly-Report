from __future__ import annotations

from pathlib import Path

# The final dashboard HTML remains standalone.
# We now keep CSS/JS/template sources in separate files and assemble them at import time,
# which makes the split reversible without changing the generated report contract.
_ASSET_DIR = Path(__file__).with_name("report_template_assets")
_TEMPLATE_PATH = _ASSET_DIR / "template.html"
_STYLE_PATH = _ASSET_DIR / "report.css"
_SCRIPT_DIR = _ASSET_DIR / "js"
_SCRIPT_PATHS = [
    _SCRIPT_DIR / "00-state-and-ingest.js",
    _SCRIPT_DIR / "10-shared-ui.js",
    _SCRIPT_DIR / "20-weekly-report.js",
    _SCRIPT_DIR / "30-flat-time-analysis.js",
    _SCRIPT_DIR / "40-dashboard-bootstrap.js",
]


def _read_asset(path: Path) -> str:
    if not path.exists():
        raise FileNotFoundError(f"Missing report asset: {path}")
    return path.read_text(encoding="utf-8")


def _build_html_template() -> str:
    template = _read_asset(_TEMPLATE_PATH)
    css = _read_asset(_STYLE_PATH)
    js = "\n\n".join(_read_asset(path).rstrip() for path in _SCRIPT_PATHS) + "\n"
    return template.replace("__REPORT_CSS__", css).replace("__REPORT_JS__", js)


HTML_TEMPLATE = _build_html_template()
