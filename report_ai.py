from __future__ import annotations

import json
import os
import re
from typing import Any
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

from app_config import get_weeklyreport_agent_name, get_weeklyreport_model, has_weeklyreport_ai


GROQ_CHAT_COMPLETIONS_URL = "https://api.groq.com/openai/v1/chat/completions"


def _build_system_prompt(agent_name: str) -> str:
    return (
        f"You are {agent_name}, an English-language drilling performance analyst focused on flat time. "
        "Use only the provided structured context. Do not invent wells, times, or benchmarks. "
        "Write a concise, polished executive report for operations and management. "
        "Always explain why the selected activities or sections were chosen, how much time can likely be recovered, "
        "and what practical actions should be taken next. Prefer direct, quantitative statements, short section headings, "
        "and clean bullet points. Never use more than 2 decimal places for hours or days. "
        "If the context is insufficient, say so clearly."
    )


def _round_numeric_values(value: Any) -> Any:
    if isinstance(value, bool):
        return value
    if isinstance(value, float):
        return round(value, 2)
    if isinstance(value, list):
        return [_round_numeric_values(item) for item in value]
    if isinstance(value, dict):
        return {key: _round_numeric_values(item) for key, item in value.items()}
    return value


def _build_user_prompt(scope: str, specific_work: str, context: dict[str, Any]) -> str:
    focus_text = specific_work.strip() if specific_work.strip() else "No extra job focus was provided."
    instructions = {
        "selected-well": "Focus the report on the selected well. Explain the highest-burden sections, main activities, offset comparison, and the most credible time recovery actions.",
        "selected-well-section": "Focus the report on the selected well within the currently selected section. Explain the section burden, offset benchmark, and the most relevant activities for that section.",
        "selected-activities": "Focus the report on the selected activities as one combined job. Sum the selected activities and explain the combined benchmark, actual time, ideal time, and actions to recover time.",
        "current-scope": "Summarize the full currently filtered comparison set. Prioritize the worst wells, most recoverable sections, and the best repeated opportunities across the loaded offsets.",
    }
    scope_instruction = instructions.get(scope, instructions["selected-well"])
    context_json = json.dumps(_round_numeric_values(context), ensure_ascii=False, indent=2)
    return (
        "Write the report in English.\n\n"
        "Required structure:\n"
        "1. Executive Summary\n"
        "2. Selected Job / Scope\n"
        "3. Main Findings\n"
        "4. Time Recovery Opportunities\n"
        "5. Recommended Actions\n\n"
        "Rules:\n"
        "- Use hours first and days in parentheses when meaningful.\n"
        "- Use no more than 2 decimal places for hours and days.\n"
        "- Keep it decision-ready, not academic.\n"
        "- Sound professional and operational, not conversational.\n"
        "- Be more explanatory than a short dashboard caption. Expand the reasoning with concrete numbers.\n"
        "- In Main Findings and Time Recovery Opportunities, explain not only what is high, but why it matters operationally.\n"
        "- Compare the selected well against offsets or ideal targets whenever that context exists.\n"
        "- If the context includes sections and drivers, explain which sections are structurally driving the loss and which activities are causing it.\n"
        "- Make the report feel presentation-ready for leadership and drilling teams.\n"
        "- Use bullets where it helps readability.\n"
        "- If there is a selected well, name it explicitly.\n"
        "- If there is a selected section, name it explicitly.\n"
        "- If there are selected activities, treat them as one combined job and sum the implications.\n\n"
        "- Do not add a separate Confidence or Caveats section.\n\n"
        f"Scope instruction: {scope_instruction}\n"
        f"Specific work / job focus: {focus_text}\n\n"
        "Structured context:\n"
        f"{context_json}"
    )


def _remove_confidence_section(report: str) -> str:
    cleaned = re.sub(
        r"(?is)\n*#{0,3}\s*6\.\s*confidence\s*/\s*caveats.*\Z",
        "",
        report,
    )
    return cleaned.strip()


def generate_flat_time_ai_report(context: dict[str, Any], scope: str, specific_work: str = "") -> dict[str, str]:
    if not has_weeklyreport_ai():
        raise RuntimeError("WeeklyReport is not configured on this server.")

    api_key = os.getenv("GROQ_API_KEY", "").strip()
    if not api_key:
        raise RuntimeError("WeeklyReport is not configured on this server.")

    agent_name = get_weeklyreport_agent_name()
    model = get_weeklyreport_model()
    payload = {
        "model": model,
        "temperature": 0.2,
        "messages": [
            {"role": "system", "content": _build_system_prompt(agent_name)},
            {"role": "user", "content": _build_user_prompt(scope, specific_work, context)},
        ],
    }
    request = Request(
        GROQ_CHAT_COMPLETIONS_URL,
        data=json.dumps(payload).encode("utf-8"),
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
            "User-Agent": f"{agent_name}/1.0 (+local weekly report server)",
        },
        method="POST",
    )

    try:
        with urlopen(request, timeout=75) as response:
            result = json.loads(response.read().decode("utf-8"))
    except HTTPError as exc:
        try:
            response_payload = json.loads(exc.read().decode("utf-8"))
            message = response_payload.get("error", {}).get("message", "")
        except Exception:  # noqa: BLE001
            message = ""
        raise RuntimeError(message or "WeeklyReport could not generate the report right now.") from exc
    except URLError as exc:
        raise RuntimeError("WeeklyReport could not reach the report service right now.") from exc

    report = (
        result.get("choices", [{}])[0]
        .get("message", {})
        .get("content", "")
        .strip()
    )
    report = _remove_confidence_section(report)
    if not report:
        raise RuntimeError("WeeklyReport returned an empty response.")

    return {
        "agent": agent_name,
        "model": model,
        "report": report,
    }
