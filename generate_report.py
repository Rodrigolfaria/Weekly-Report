#!/usr/bin/env python3
from __future__ import annotations

import argparse
from pathlib import Path

# Compatibility facade:
# - server.py and any local scripts can keep importing from generate_report.py
# - implementation now lives in report_assets.py, report_parsers.py, and report_builder.py
# - rolling back is straightforward because the public function names are preserved here
from report_assets import HTML_TEMPLATE
from report_builder import (
    build_csv_sheet,
    build_csv_report_html,
    build_dashboard_payload,
    build_empty_report_html,
    build_flat_time_csv_report_html,
    build_html,
    build_intervention_csv_report_html,
    build_report,
    build_report_html,
)
from report_flat_time import (
    extract_flat_time_section_size,
    load_activity_code_translations,
    load_flat_time_directory,
    looks_like_flat_time_csv,
    parse_flat_time_matrix_rows,
    parse_flat_time_rows,
)
from report_parsers import (
    DOCUMENT_REL_NS,
    NS,
    PACKAGE_REL_NS,
    SPREADSHEET_NS,
    SheetData,
    as_number,
    build_intervention_rows,
    column_letters_to_index,
    excel_serial_to_datetime,
    infer_header_row,
    iso_date,
    iso_month,
    iso_week,
    load_sheet_rows,
    load_shared_strings,
    load_workbook,
    normalize_header,
    normalize_text,
    parse_datetime_value,
    read_cell_value,
    rows_to_dicts,
    sorted_unique,
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Generate an interactive HTML report from an Excel spreadsheet.")
    parser.add_argument(
        "spreadsheet",
        nargs="?",
        help="Path to the .xlsx file. Defaults to the first .xlsx found in the current directory.",
    )
    parser.add_argument(
        "--output",
        help="Path to the generated HTML report. Defaults to report_output/report.html",
    )
    return parser.parse_args()


def find_default_spreadsheet(cwd: Path) -> Path:
    candidates = sorted(cwd.glob("*.xlsx"))
    if not candidates:
        raise FileNotFoundError("No .xlsx files found in the current directory.")
    return candidates[0]


def main() -> int:
    args = parse_args()
    cwd = Path.cwd()
    spreadsheet = Path(args.spreadsheet).expanduser() if args.spreadsheet else find_default_spreadsheet(cwd)
    if not spreadsheet.is_absolute():
        spreadsheet = cwd / spreadsheet
    if not spreadsheet.exists():
        raise FileNotFoundError(f"Spreadsheet not found: {spreadsheet}")

    output = Path(args.output).expanduser() if args.output else cwd / "report_output" / "report.html"
    if not output.is_absolute():
        output = cwd / output

    build_report(spreadsheet, output)
    print(output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
