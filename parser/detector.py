from __future__ import annotations


def detect_invoice_type(text: str) -> str:
    upper = text.upper()
    if "SPECIAL ORDER INVOICE" in upper:
        return "SPECIAL_ORDER"
    if "LSC AR INVOICE" in upper:
        return "LSC_AR"
    if "PLCB WHOLESALE LICENSEE ORDER" in upper:
        return "WHOLESALE"
    raise ValueError("Tipo fattura non riconosciuto")
