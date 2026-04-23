from __future__ import annotations

import csv
import html
import json
from datetime import datetime
from pathlib import Path
from typing import Any

from report_assets import HTML_TEMPLATE
from report_flat_time import load_activity_code_translations
from report_parsers import (
    SheetData,
    build_cost_avoidance_rows,
    build_intervention_rows,
    load_workbook,
    rows_to_dicts,
    sorted_unique,
)

def build_dashboard_payload(
    spreadsheet_name: str,
    generated_at: str,
    intervention_rows: list[dict[str, Any]],
    cost_avoidance_rows: list[dict[str, Any]],
    flat_time_payload: dict[str, Any] | None = None,
    activity_code_translations: dict[str, Any] | None = None,
) -> dict[str, Any]:
    dates = [row["date"] for row in intervention_rows if row["date"]]
    payload = {
        "meta": {
            "sourceFile": spreadsheet_name,
            "generatedAt": generated_at,
            "minDate": min(dates) if dates else "",
            "maxDate": max(dates) if dates else "",
            "monitoringStartDate": "2025-11-11",
        },
        "filters": {
            "weeks": sorted_unique([row["week"] for row in intervention_rows]),
            "months": sorted_unique([row["month"] for row in intervention_rows]),
            "rigs": sorted_unique([row["rigName"] for row in intervention_rows]),
            "fields": sorted_unique([row["field"] for row in intervention_rows]),
            "wells": sorted_unique([row["wellName"] for row in intervention_rows]),
            "categories": sorted_unique([row["category"] for row in intervention_rows]),
            "types": sorted_unique([row["type"] for row in intervention_rows]),
            "apps": sorted_unique([row["app"] for row in intervention_rows]),
            "reps": sorted_unique([row["rtesRep"] for row in intervention_rows]),
        },
        "interventions": intervention_rows,
        "costAvoidance": cost_avoidance_rows,
    }
    if flat_time_payload is not None:
        payload["flatTime"] = flat_time_payload
    if activity_code_translations is not None:
        payload["activityCodeTranslations"] = activity_code_translations
    return payload


def build_html(title: str, source_file: str, generated_at: str, payload: dict[str, Any]) -> str:
    output = HTML_TEMPLATE
    output = output.replace("__TITLE__", html.escape(title))
    output = output.replace("__SOURCE_FILE__", html.escape(source_file))
    output = output.replace("__GENERATED_AT__", html.escape(generated_at))
    output = output.replace("__DATA_JSON__", json.dumps(payload, separators=(",", ":"), ensure_ascii=True))
    return output


def build_report_html(spreadsheet_name: str, spreadsheet_bytes: bytes, flat_time_payload: dict[str, Any] | None = None) -> str:
    sheets = load_workbook(spreadsheet_bytes)
    sheet_map = {sheet.name: sheet for sheet in sheets}

    interventions_source = sheet_map.get("Intervention Log", SheetData("Intervention Log", [], [])).rows
    cost_avoidance_source = sheet_map.get("RTES CA", SheetData("RTES CA", [], [])).rows

    intervention_rows = build_intervention_rows(interventions_source)
    cost_avoidance_rows = build_cost_avoidance_rows(cost_avoidance_source)

    generated_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    payload = build_dashboard_payload(
        spreadsheet_name=spreadsheet_name,
        generated_at=generated_at,
        intervention_rows=intervention_rows,
        cost_avoidance_rows=cost_avoidance_rows,
        flat_time_payload=flat_time_payload,
        activity_code_translations=load_activity_code_translations(),
    )

    return build_html(
        title=f"Interactive Report - {spreadsheet_name}",
        source_file=spreadsheet_name,
        generated_at=generated_at,
        payload=payload,
    )


def build_csv_sheet(csv_name: str, csv_bytes: bytes) -> SheetData:
    rows = list(csv.reader(csv_bytes.decode("utf-8-sig", errors="ignore").splitlines()))
    return rows_to_dicts(csv_name, rows)


def build_intervention_csv_report_html(csv_name: str, csv_bytes: bytes, flat_time_payload: dict[str, Any] | None = None) -> str:
    csv_sheet = build_csv_sheet(csv_name, csv_bytes)
    intervention_rows = build_intervention_rows(csv_sheet.rows)
    generated_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    payload = build_dashboard_payload(
        spreadsheet_name=f"{csv_name} (Intervention Log CSV only)",
        generated_at=generated_at,
        intervention_rows=intervention_rows,
        cost_avoidance_rows=[],
        flat_time_payload=flat_time_payload,
        activity_code_translations=load_activity_code_translations(),
    )

    return build_html(
        title=f"Interactive Report - {csv_name}",
        source_file=f"{csv_name} (CSV mode)",
        generated_at=generated_at,
        payload=payload,
    )


def build_empty_report_html(flat_time_payload: dict[str, Any] | None = None) -> str:
    generated_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    payload = build_dashboard_payload(
        spreadsheet_name="No workbook loaded",
        generated_at=generated_at,
        intervention_rows=[],
        cost_avoidance_rows=[],
        flat_time_payload=flat_time_payload if flat_time_payload is not None else {"datasets": []},
        activity_code_translations=load_activity_code_translations(),
    )

    return build_html(
        title="Interactive Report - No workbook loaded",
        source_file="No workbook loaded",
        generated_at=generated_at,
        payload=payload,
    )


def build_report(spreadsheet: Path, output_path: Path) -> None:
    html_output = build_report_html(
        spreadsheet.name,
        spreadsheet.read_bytes(),
        flat_time_payload={"datasets": []},
    )
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(html_output, encoding="utf-8")
