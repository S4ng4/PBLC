from __future__ import annotations

import re
from typing import List, Tuple

from money import parse_money


def normalize_ws(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip())


def find_label_value(text: str, label: str) -> str:
    """Return text after `Label` on same line, if present."""
    pattern = re.compile(rf"{re.escape(label)}\s*[:\s]\s*(.+)$", re.IGNORECASE | re.MULTILINE)
    m = pattern.search(text)
    if not m:
        return ""
    return normalize_ws(m.group(1))


def _label_match_variants(label: str) -> List[str]:
    """
    PLCB PDFs use both 'Sales Tax (6%):' and 'Sales Tax(6%):' (space before '(' optional).
    Build a small set of substrings to look for in each line.
    """
    base = label.strip()
    out: List[str] = [base]
    collapsed = re.sub(r"(\S)\s+\(", r"\1(", base)
    if collapsed not in out:
        out.append(collapsed)
    expanded = re.sub(r"(\S)\(", r"\1 (", base)
    if expanded not in out:
        out.append(expanded)
    return out


def find_money_after_label(text: str, label: str) -> float:
    """
    Find a monetary value on the same line as label, or immediately following line.
    Handles '(25.14)' style negatives.
    """
    variants = [v.lower() for v in _label_match_variants(label)]
    lines = text.splitlines()
    for i, line in enumerate(lines):
        low = line.lower()
        if not any(v in low for v in variants):
            continue
        tail = line.split(":", 1)[-1] if ":" in line else line
        amounts = re.findall(r"\(?\$?\s*-?\d[\d,]*(?:\.\d+)?\)?", tail)
        if amounts:
            return parse_money(amounts[-1])
        if i + 1 < len(lines):
            nxt = lines[i + 1].strip()
            if re.search(r"\$?\(?\d", nxt):
                am = re.findall(r"\(?\$?\s*-?\d[\d,]*(?:\.\d+)?\)?", nxt)
                if am:
                    return parse_money(am[0])
    for v in variants:
        for m in re.finditer(
            rf"{re.escape(v)}[^\n$]*(\(?\$?\s*-?\d[\d,]*(?:\.\d+)?\)?)", text, flags=re.IGNORECASE
        ):
            return parse_money(m.group(1))
    return 0.0


def is_probable_product_start(line: str) -> bool:
    """Heuristic: line number + PLCB-style numeric item code."""
    s = line.strip()
    return bool(re.match(r"^\d{1,3}\s+\d{4,10}\b", s))


def merge_wrapped_product_block(lines: List[str], start_idx: int) -> Tuple[List[str], int]:
    """
    From lines[start_idx], consume product rows; join continuation lines that do not
    start a new numbered item row (multi-line descriptions from PDF extraction).
    """
    merged: List[str] = []
    i = start_idx
    n = len(lines)
    while i < n:
        line = lines[i].strip()
        if not line:
            i += 1
            continue
        if not is_probable_product_start(line):
            break
        buf = line
        i += 1
        while i < n:
            nxt = lines[i].strip()
            if not nxt:
                i += 1
                break
            if is_probable_product_start(nxt):
                break
            buf = f"{buf} {nxt}"
            i += 1
        merged.append(buf)
    return merged, i


def split_trailing_money_cells(tail: str) -> Tuple[List[str], List[float]]:
    """
    From the end of a product row string, peel off numeric tokens (money/qty).
    Returns (text_tokens_from_start, numbers_from_end_reversed_then_ordered).
    """
    s = normalize_ws(tail)
    nums: List[float] = []
    parts = s.split(" ")
    while parts:
        tok = parts[-1]
        if re.fullmatch(r"\(?\$?\s*-?\d[\d,]*(?:\.\d+)?\)?", tok.replace(",", "")) or re.fullmatch(
            r"\d+", tok.replace(",", "")
        ):
            nums.append(parse_money(tok))
            parts.pop()
            continue
        if re.fullmatch(r"[A-Za-z]+", tok) and parts:
            break
        break
    nums.reverse()
    return parts, nums
