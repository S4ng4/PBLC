from __future__ import annotations

import re


def parse_money(value: str | float | int | None) -> float:
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip()
    s = s.replace("$", "").replace(",", "").strip()
    s = re.sub(r"\s+", "", s)
    if not s:
        return 0.0
    if s.startswith("(") and s.endswith(")"):
        inner = s[1:-1].replace("$", "").replace(",", "").strip()
        try:
            return -float(inner)
        except ValueError:
            return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0
