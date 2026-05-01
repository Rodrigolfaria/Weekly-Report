from __future__ import annotations

import os
from pathlib import Path


def load_local_env() -> None:
    candidates = [
        Path.cwd() / ".env.local",
        Path(__file__).resolve().parent / ".env.local",
    ]

    for path in candidates:
        if not path.exists():
            continue
        for line in path.read_text(encoding="utf-8").splitlines():
            stripped = line.strip()
            if not stripped or stripped.startswith("#") or "=" not in stripped:
                continue
            key, value = stripped.split("=", 1)
            key = key.strip()
            value = value.strip().strip('"').strip("'")
            if key and key not in os.environ:
                os.environ[key] = value
        break


load_local_env()


def get_weeklyreport_agent_name() -> str:
    return os.getenv("WEEKLYREPORT_AGENT_NAME", "WeeklyReport").strip() or "WeeklyReport"


def get_weeklyreport_model() -> str:
    return os.getenv("WEEKLYREPORT_MODEL", "llama-3.3-70b-versatile").strip() or "llama-3.3-70b-versatile"


def has_weeklyreport_ai() -> bool:
    return bool(os.getenv("GROQ_API_KEY", "").strip())
