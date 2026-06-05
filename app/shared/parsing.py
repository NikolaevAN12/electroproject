from __future__ import annotations


def parse_number(raw: str) -> float | None:
    text = raw.strip().replace(" ", "").replace(",", ".")
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None

