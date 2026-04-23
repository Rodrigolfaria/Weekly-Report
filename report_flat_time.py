from __future__ import annotations

import csv
import re
from pathlib import Path
from typing import Any


def as_number(value: str | None) -> float:
    try:
        return float((value or "0").replace(",", "").replace("$", "").replace("\xa0", "").strip())
    except ValueError:
        return 0.0


def extract_flat_time_section_size(activity_name: str) -> str:
    match = re.match(r"^(\d+(?:\.\d+)?)(?=-)", activity_name or "")
    return match.group(1) if match else "__no_section__"


def parse_flat_time_rows(file_name: str, rows: list[list[str]]) -> dict[str, Any] | None:
    normalized_rows = [[(cell or "").strip() for cell in row] for row in rows]
    subject_well = next((row[1].strip() for row in normalized_rows if len(row) > 1 and row[0] == "Subject Well"), file_name)
    groups: list[dict[str, Any]] = []
    current_group: dict[str, Any] | None = None

    for row in normalized_rows:
        first = row[0] if row else ""
        if first == "Group Name":
            if current_group and current_group["activities"]:
                groups.append(current_group)
            current_group = {
                "groupName": row[1] if len(row) > 1 else "Unknown",
                "activities": [],
                "totalSubjectHours": 0.0,
                "totalMeanHours": 0.0,
                "totalMedianHours": 0.0,
            }
            continue

        if current_group is None or not first or first in {"Group Type", "Activity"}:
            continue

        if first == "Total":
            current_group["totalSubjectHours"] = as_number(row[1] if len(row) > 1 else 0)
            continue

        current_group["activities"].append(
            {
                "activity": first,
                "sectionSize": extract_flat_time_section_size(first),
                "subjectHours": as_number(row[1] if len(row) > 1 else 0),
                "meanHours": as_number(row[2] if len(row) > 2 else 0),
                "medianHours": as_number(row[3] if len(row) > 3 else 0),
            }
        )

    if current_group and current_group["activities"]:
        groups.append(current_group)

    if not groups:
        return None

    for group in groups:
        if not group["totalSubjectHours"]:
            group["totalSubjectHours"] = sum(activity["subjectHours"] for activity in group["activities"])
        group["totalMeanHours"] = sum(activity["meanHours"] for activity in group["activities"])
        group["totalMedianHours"] = sum(activity["medianHours"] for activity in group["activities"])

    return {
        "id": f"{Path(file_name).stem.lower().replace(' ', '-').replace('_', '-')}-{subject_well.lower().replace(' ', '-')}",
        "fileName": file_name,
        "subjectWell": subject_well,
        "groups": groups,
        "totalSubjectHours": sum(group["totalSubjectHours"] for group in groups),
        "totalMeanHours": sum(group["totalMeanHours"] for group in groups),
        "totalMedianHours": sum(group["totalMedianHours"] for group in groups),
    }


def load_flat_time_directory(directory: Path) -> dict[str, Any]:
    datasets: list[dict[str, Any]] = []
    for path in sorted(directory.glob("*.csv")):
        if path.name.startswith("."):
            continue
        try:
            with path.open(newline="", encoding="utf-8-sig") as handle:
                rows = list(csv.reader(handle))
        except OSError:
            continue
        parsed = parse_flat_time_rows(path.name, rows)
        if parsed:
            datasets.append(parsed)
    return {"datasets": datasets}


def load_activity_code_translations() -> dict[str, Any]:
    candidates = [
        Path.cwd() / "Aramco Activity Codes.csv",
        Path(__file__).with_name("Aramco Activity Codes.csv"),
        Path.home() / "Downloads" / "Aramco Activity Codes.csv",
    ]

    source_path = next((path for path in candidates if path.exists()), None)
    if source_path is None:
        return {
            "loaded": False,
            "source": "",
            "wellSections": {},
            "operations": {},
            "activities": {},
            "generic": {},
        }

    well_sections: dict[str, str] = {}
    operations: dict[str, str] = {}
    activities: dict[str, str] = {}
    generic: dict[str, str] = {}

    try:
        with source_path.open(newline="", encoding="utf-8-sig") as handle:
            for row in csv.DictReader(handle):
                code = (row.get("Client Code") or "").strip()
                description = (row.get("Client Code Description") or "").strip()
                code_type = (row.get("Corva Code Type") or "").strip().lower()
                if not code or not description:
                    continue
                generic.setdefault(code, description)
                if code_type == "well sections":
                    well_sections.setdefault(code, description)
                elif code_type == "operation":
                    operations.setdefault(code, description)
                elif code_type == "activity":
                    activities.setdefault(code, description)
    except OSError:
        return {
            "loaded": False,
            "source": "",
            "wellSections": {},
            "operations": {},
            "activities": {},
            "generic": {},
        }

    return {
        "loaded": True,
        "source": source_path.name,
        "wellSections": well_sections,
        "operations": operations,
        "activities": activities,
        "generic": generic,
    }
