"""Logica condivisa tra GUI Tk e interfaccia web (senza importare tkinter)."""
from __future__ import annotations

import re
from pathlib import Path

from calculator.lsc_wholesale import evaluate_margin_edge_sum_p as sum_p_lw
from calculator.special_order import evaluate_margin_edge_sum_p as sum_p_so
from models import InvoiceSummary, verify_totals
from parser.detector import detect_invoice_type
from parser.lsc_wholesale import parse_lsc_wholesale_invoice
from parser.special_order import parse_special_order_invoice
from pdf_utils import extract_text


def safe_invoice_filename(inv: str) -> str:
    s = (inv or "UNKNOWN").strip()
    s = re.sub(r"[^\w\-]+", "_", s, flags=re.UNICODE)
    return s or "UNKNOWN"


def load_summary(pdf_path: str) -> tuple[InvoiceSummary, str]:
    raw = extract_text(pdf_path)
    kind = detect_invoice_type(raw)
    if kind == "SPECIAL_ORDER":
        summary = parse_special_order_invoice(pdf_path)
    else:
        summary = parse_lsc_wholesale_invoice(pdf_path, kind)
    if not summary.invoice_number:
        summary.invoice_number = Path(pdf_path).stem
    return summary, raw


def compute_sum_p(summary: InvoiceSummary) -> float:
    if summary.invoice_type == "SPECIAL_ORDER":
        return sum_p_so(summary)
    return sum_p_lw(summary)


def verify_invoice(pdf_path: str) -> tuple[InvoiceSummary, str, float, bool, float]:
    summary, raw = load_summary(pdf_path)
    sp = compute_sum_p(summary)
    ok, diff = verify_totals(summary, sp)
    return summary, raw, sp, ok, diff
