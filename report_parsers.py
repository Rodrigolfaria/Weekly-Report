from __future__ import annotations

import csv
import re
from dataclasses import dataclass
from datetime import datetime, timedelta
from io import BytesIO
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from report_flat_time import (
    extract_flat_time_section_size,
    load_activity_code_translations,
    load_flat_time_directory,
    parse_flat_time_rows,
)

SPREADSHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
DOCUMENT_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {"a": SPREADSHEET_NS, "r": DOCUMENT_REL_NS, "pr": PACKAGE_REL_NS}

@dataclass
class SheetData:
    name: str
    headers: list[str]
    rows: list[dict[str, str]]


CANONICAL_INTERVENTION_CATEGORIES = {
    "operational compliance": "Operational Compliance",
    "optimization": "Optimization",
    "stuck pipe": "Stuck pipe",
    "well control": "Well control",
    "reporting": "Reporting",
}


def excel_serial_to_datetime(value: str) -> datetime | None:
    try:
        serial = float(value)
    except (TypeError, ValueError):
        return None
    return datetime(1899, 12, 30) + timedelta(days=serial)


def parse_datetime_value(value: str | None) -> datetime | None:
    text = str(value or "").strip()
    if not text:
        return None

    serial_date = excel_serial_to_datetime(text)
    if serial_date:
        return serial_date

    normalized = (
        text.replace(".", "/")
        .replace("-", "/")
        .replace("\\", "/")
        .replace("T", " ")
    )
    normalized = re.sub(r"\s+", " ", normalized).strip()
    normalized = normalized.split(" ")[0]

    for fmt in (
        "%m/%d/%Y",
        "%m/%d/%y",
        "%d/%m/%Y",
        "%d/%m/%y",
        "%Y/%m/%d",
    ):
        try:
            return datetime.strptime(normalized, fmt)
        except ValueError:
            continue
    return None


def canonical_intervention_category(value: str | None) -> str:
    text = normalize_text(value)
    if not text:
        return ""
    return CANONICAL_INTERVENTION_CATEGORIES.get(text.lower(), text)


def column_letters_to_index(cell_ref: str) -> int:
    letters = "".join(char for char in cell_ref if char.isalpha())
    result = 0
    for char in letters:
        result = (result * 26) + (ord(char.upper()) - 64)
    return result - 1


def load_shared_strings(workbook: ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in workbook.namelist():
        return []
    root = ET.fromstring(workbook.read("xl/sharedStrings.xml"))
    items = []
    for item in root.findall("a:si", NS):
        text = "".join(node.text or "" for node in item.findall(".//a:t", NS))
        items.append(text)
    return items


def read_cell_value(cell: ET.Element, shared_strings: list[str]) -> str:
    cell_type = cell.attrib.get("t")
    if cell_type == "inlineStr":
        return "".join(node.text or "" for node in cell.findall(".//a:t", NS)).strip()

    value_node = cell.find("a:v", NS)
    if value_node is None or value_node.text is None:
        return ""

    raw_value = value_node.text.strip()
    if cell_type == "s":
        try:
            return shared_strings[int(raw_value)]
        except (ValueError, IndexError):
            return raw_value
    if cell_type == "b":
        return "TRUE" if raw_value == "1" else "FALSE"
    return raw_value


def load_sheet_rows(workbook: ZipFile, sheet_path: str, shared_strings: list[str]) -> list[list[str]]:
    root = ET.fromstring(workbook.read(sheet_path))
    rows: list[list[str]] = []
    for row in root.findall(".//a:sheetData/a:row", NS):
        values: list[str] = []
        for cell in row.findall("a:c", NS):
            ref = cell.attrib.get("r", "")
            index = column_letters_to_index(ref) if ref else len(values)
            while len(values) <= index:
                values.append("")
            values[index] = read_cell_value(cell, shared_strings)
        rows.append(values)
    return rows


def normalize_header(value: str, index: int) -> str:
    text = " ".join((value or "").replace("\xa0", " ").split()).strip()
    return text or f"Column {index + 1}"


def infer_header_row(rows: list[list[str]]) -> int:
    best_index = 0
    best_score = -1
    for index, row in enumerate(rows[:10]):
        normalized = [normalize_header(cell, idx) for idx, cell in enumerate(row)]
        non_empty = sum(1 for cell in row if str(cell).strip())
        unique = len(set(cell for cell in normalized if cell))
        score = (non_empty * 10) + unique
        if score > best_score:
            best_index = index
            best_score = score
    return best_index


def rows_to_dicts(name: str, rows: list[list[str]]) -> SheetData:
    if not rows:
        return SheetData(name=name, headers=[], rows=[])

    header_index = infer_header_row(rows)
    headers = [normalize_header(cell, idx) for idx, cell in enumerate(rows[header_index])]
    width = len(headers)
    records: list[dict[str, str]] = []

    for raw_row in rows[header_index + 1 :]:
        padded = list(raw_row[:width]) + [""] * max(0, width - len(raw_row))
        record = {headers[idx]: (padded[idx] or "").strip() for idx in range(width)}
        if any(value for value in record.values()):
            records.append(record)

    return SheetData(name=name, headers=headers, rows=records)


def load_workbook(source: Path | bytes | BytesIO) -> list[SheetData]:
    workbook_source = BytesIO(source) if isinstance(source, bytes) else source
    with ZipFile(workbook_source) as workbook:
        shared_strings = load_shared_strings(workbook)
        workbook_root = ET.fromstring(workbook.read("xl/workbook.xml"))
        rels_root = ET.fromstring(workbook.read("xl/_rels/workbook.xml.rels"))
        rel_map = {
            rel.attrib["Id"]: f"xl/{rel.attrib['Target']}"
            for rel in rels_root.findall("pr:Relationship", NS)
        }

        sheets = []
        sheets_node = workbook_root.find("a:sheets", NS)
        for sheet in list(sheets_node) if sheets_node is not None else []:
            name = sheet.attrib["name"]
            rel_id = sheet.attrib[f"{{{DOCUMENT_REL_NS}}}id"]
            rows = load_sheet_rows(workbook, rel_map[rel_id], shared_strings)
            sheets.append(rows_to_dicts(name, rows))
        return sheets


def as_number(value: str | None) -> float:
    if value is None:
        return 0.0
    text = str(value).replace(",", "").replace("$", "").replace("\xa0", "").strip()
    if not text:
        return 0.0
    try:
        return float(text)
    except ValueError:
        return 0.0


def iso_date(value: str | None) -> str:
    date_value = parse_datetime_value(value)
    if not date_value:
        return ""
    return date_value.strftime("%Y-%m-%d")


def iso_week(value: str | None) -> str:
    date_value = parse_datetime_value(value)
    if not date_value:
        return ""
    iso_year, iso_week_number, _ = date_value.isocalendar()
    return f"{iso_year}-W{iso_week_number:02d}"


def iso_month(value: str | None) -> str:
    date_value = parse_datetime_value(value)
    if not date_value:
        return ""
    return date_value.strftime("%Y-%m")


def normalize_text(value: str | None) -> str:
    return " ".join(str(value or "").replace("\xa0", " ").split()).strip()


def build_intervention_rows(rows: list[dict[str, str]]) -> list[dict[str, Any]]:
    output: list[dict[str, Any]] = []
    for row in rows:
        date = iso_date(row.get("Date"))
        record = {
            "index": normalize_text(row.get("Interventions Count")),
            "date": date,
            "week": iso_week(row.get("Date")),
            "month": iso_month(row.get("Date")),
            "rigName": normalize_text(row.get("Rig Name")),
            "field": normalize_text(row.get("Field")),
            "wellName": normalize_text(row.get("Well Name")),
            "holeSize": normalize_text(row.get("Hole Size")),
            "engDept": normalize_text(row.get("ENG. Dept")),
            "optDept": normalize_text(row.get("Opt. Dept")),
            "category": canonical_intervention_category(row.get("Intervention Category")),
            "type": normalize_text(row.get("Intervention Type")),
            "eventIndex": normalize_text(row.get("Event Index")),
            "app": normalize_text(row.get("Corva App")),
            "parameter": normalize_text(row.get("Parameter")),
            "expected": normalize_text(row.get("Expected")),
            "actual": normalize_text(row.get("Actual")),
            "description": normalize_text(row.get("Intervention Description")),
            "recommendation": normalize_text(row.get("Recommendation")),
            "justification": normalize_text(row.get("Cost Saving/Potential Cost Avoidance Justification")),
            "validationText": normalize_text(row.get("RTOC/RDH Validation (Y/N)")),
            "isValidated": normalize_text(row.get("RTOC/RDH Validation (Y/N)")).lower() in {"yes", "y", "true"},
            "rtocComments": normalize_text(row.get("RTOC Comments")),
            "rtocCommunication": normalize_text(row.get("RTOC to Rig Communication")),
            "rigAction": normalize_text(row.get("Rig Taken Action")),
            "rigComment": normalize_text(row.get("Rig Comment")),
            "rtocLeadName": normalize_text(row.get("RTOC lead name")),
            "costSavingHours": as_number(row.get("Cost Saving (hrs)")),
            "potentialAvoidanceHours": as_number(row.get("Potential Cost Avoidance (hrs)")),
            "rigSpreadRate": as_number(row.get("Rig Spread Rate ($/day)")),
            "costSavingValue": as_number(row.get("$ Cost Saving")),
            "potentialAvoidanceValue": as_number(row.get("$ Potential Cost Avoidance")),
        }
        search_fields = [
            record["date"],
            record["week"],
            record["month"],
            record["rigName"],
            record["field"],
            record["wellName"],
            record["category"],
            record["type"],
            record["app"],
            record["parameter"],
            record["description"],
            record["recommendation"],
            record["justification"],
            record["rtocComments"],
            record["rigComment"],
        ]
        record["searchText"] = " ".join(value.lower() for value in search_fields if value)
        if (
            (record["rigName"] or record["wellName"])
            and any(
                [
                    record["date"],
                    record["rigName"],
                    record["field"],
                    record["wellName"],
                    record["category"],
                    record["type"],
                    record["app"],
                    record["description"],
                ]
            )
        ):
            output.append(record)
    return output


def build_cost_avoidance_rows(rows: list[dict[str, str]]) -> list[dict[str, Any]]:
    output: list[dict[str, Any]] = []
    for row in rows:
        record = {
            "startDate": iso_date(row.get("MONITOR_STA_DTTM")),
            "endDate": iso_date(row.get("MONITOR_END_DTTM")),
            "rig": normalize_text(row.get("Rig")),
            "well": normalize_text(row.get("Well")),
            "daysSaved": as_number(row.get("Days Saved")),
            "costAvoidanceValue": as_number(row.get("Cost Avoidance US$")),
            "costIncluded": normalize_text(row.get("cost included")),
            "caDisplay": normalize_text(row.get("RTOC CA")),
        }
        search_fields = [
            record["startDate"],
            record["endDate"],
            record["rig"],
            record["well"],
            record["costIncluded"],
            record["caDisplay"],
        ]
        record["searchText"] = " ".join(value.lower() for value in search_fields if value)
        if any([record["startDate"], record["endDate"], record["rig"], record["well"]]):
            output.append(record)
    return output


def build_cost_avoidance_rows_from_interventions(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    output: list[dict[str, Any]] = []
    for row in rows:
        cost_saving_hours = float(row.get("costSavingHours", 0.0) or 0.0)
        potential_avoidance_hours = float(row.get("potentialAvoidanceHours", 0.0) or 0.0)
        cost_saving_value = float(row.get("costSavingValue", 0.0) or 0.0)
        potential_avoidance_value = float(row.get("potentialAvoidanceValue", 0.0) or 0.0)
        total_hours = cost_saving_hours + potential_avoidance_hours
        total_value = cost_saving_value + potential_avoidance_value
        if total_hours <= 0 and total_value <= 0:
            continue

        record = {
            "startDate": row.get("date", ""),
            "endDate": row.get("date", ""),
            "rig": row.get("rigName", ""),
            "well": row.get("wellName", ""),
            "daysSaved": total_hours / 24 if total_hours else 0.0,
            "costAvoidanceValue": total_value,
            "costIncluded": "Intervention Log derived",
            "caDisplay": row.get("justification", "") or row.get("description", "") or row.get("recommendation", ""),
        }
        search_fields = [
            record["startDate"],
            record["endDate"],
            record["rig"],
            record["well"],
            record["costIncluded"],
            record["caDisplay"],
        ]
        record["searchText"] = " ".join(value.lower() for value in search_fields if value)
        output.append(record)
    return output


def sorted_unique(values: list[str]) -> list[str]:
    return sorted({value for value in values if value})
