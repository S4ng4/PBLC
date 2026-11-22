from __future__ import annotations

import re
from typing import Any, List, Optional, Sequence

from models import InvoiceItem, InvoiceSummary
from money import parse_money
from pdf_utils import extract_tables, extract_text

from ._text_parse import find_label_value, find_money_after_label, merge_wrapped_product_block


def _norm_cell(x: Any) -> str:
    if x is None:
        return ""
    return re.sub(r"\s+", " ", str(x).strip())


def _header_map(header: Sequence[Any]) -> dict[str, int]:
    mapping: dict[str, int] = {}
    for idx, raw in enumerate(header):
        key = re.sub(r"\s+", " ", str(raw or "")).strip().upper()
        if not key:
            continue
        if "PLCB" in key and "ITEM" in key:
            mapping["PLCB"] = idx
        if "DESCRIPTION" in key:
            mapping["DESC"] = idx
        if "UNIT" in key and "VOL" in key:
            mapping["VOL"] = idx
        if key == "UOM":
            mapping["UOM"] = idx
        if "BASE" in key and "PRICE" in key:
            mapping["BASE"] = idx
        if "BASE" in key and "FREIGHT" in key:
            mapping["BFREIGHT"] = idx
        if "LTMF" in key:
            mapping["LTMF"] = idx
        if "TOTAL" in key and "UNIT" in key and "PRICE" in key:
            mapping["TUNIT"] = idx
        if "UNIT" in key and "QTY" in key:
            mapping["QTY"] = idx
        if "LINE" in key and "AMOUNT" in key:
            mapping["LAMOUNT"] = idx
        if re.fullmatch(r"LINE(\s*#|\s*NO\.?)?", key):
            mapping.setdefault("LINE", idx)
    return mapping


def _cell(row: Sequence[Any], idx: int) -> str:
    if idx < 0 or idx >= len(row):
        return ""
    return _norm_cell(row[idx])


def _row_to_item_mapped(row: Sequence[Any], col: dict[str, int]) -> Optional[InvoiceItem]:
    line_s = _cell(row, col.get("LINE", 0))
    code = _cell(row, col.get("PLCB", 1))
    try:
        line_no = int(re.match(r"^(\d+)", line_s).group(1))  # type: ignore[union-attr]
    except (AttributeError, ValueError):
        return None
    if not re.fullmatch(r"\d{4,12}", code):
        return None

    desc = _cell(row, col["DESC"]) if "DESC" in col else ""
    vol = _cell(row, col["VOL"]) if "VOL" in col else ""
    uom = _cell(row, col["UOM"]) if "UOM" in col else ""
    base_price = parse_money(_cell(row, col["BASE"])) if "BASE" in col else 0.0
    base_freight = parse_money(_cell(row, col["BFREIGHT"])) if "BFREIGHT" in col else 0.0
    ltmf = parse_money(_cell(row, col["LTMF"])) if "LTMF" in col else 0.0
    total_unit = parse_money(_cell(row, col["TUNIT"])) if "TUNIT" in col else 0.0
    qty_f = parse_money(_cell(row, col["QTY"])) if "QTY" in col else 0.0
    line_amt = parse_money(_cell(row, col["LAMOUNT"])) if "LAMOUNT" in col else 0.0
    unit_qty = int(round(qty_f))

    if not desc:
        parts = [_norm_cell(c) for c in row if _norm_cell(c)]
        desc = " ".join(parts[2:])

    if total_unit <= 0 and base_price > 0:
        total_unit = base_price + base_freight + ltmf
    if not vol:
        mv = re.search(r"(\d+\s*(?:ML|L|oz|OZ))\b", desc, re.I)
        if mv:
            vol = mv.group(1)
    if not uom:
        mu = re.search(r"\b(EA|CS|CASE|BTL|BOTTLE)\b", desc, re.I)
        if mu:
            uom = mu.group(1).upper()
    uom = uom or "EA"

    if line_amt == 0 and total_unit and unit_qty:
        line_amt = total_unit * unit_qty

    return InvoiceItem(
        line=line_no,
        plcb_code=code,
        description=desc,
        unit_volume=vol,
        uom=uom,
        base_price=base_price,
        base_freight=base_freight,
        ltmf=ltmf,
        total_unit_price=total_unit,
        unit_qty=unit_qty,
        line_amount=line_amt,
    )


def _parse_items_from_tables(tables: List[List[Any]]) -> List[InvoiceItem]:
    items: List[InvoiceItem] = []
    for tbl in tables:
        if not tbl:
            continue
        header_idx = None
        col: dict[str, int] = {}
        for i, row in enumerate(tbl):
            joined = " ".join(_norm_cell(c) for c in row).upper()
            if "LINE" in joined and "PLCB" in joined and "DESCRIPTION" in joined:
                header_idx = i
                col = _header_map(row)
                break
        if header_idx is None or not col:
            continue
        for row in tbl[header_idx + 1 :]:
            if not any(_norm_cell(c) for c in row):
                continue
            it = _row_to_item_mapped(row, col)
            if it:
                items.append(it)
    dedup: dict[int, InvoiceItem] = {}
    for it in items:
        dedup[it.line] = it
    return [dedup[k] for k in sorted(dedup)]


def _parse_items_from_text(text: str) -> List[InvoiceItem]:
    lines = text.splitlines()
    start = None
    for i, ln in enumerate(lines):
        u = ln.upper()
        if "LINE" in u and "PLCB" in u and "DESCRIPTION" in u:
            start = i + 1
            break
    if start is None:
        return []
    block, _ = merge_wrapped_product_block(lines, start)
    out: List[InvoiceItem] = []
    for raw in block:
        cells = raw.split()
        if len(cells) < 6:
            continue
        try:
            line_no = int(cells[0])
        except ValueError:
            continue
        code = cells[1]
        if not re.fullmatch(r"\d{4,12}", code):
            continue
        tail_nums: List[float] = []
        j = len(cells) - 1
        while j > 1 and len(tail_nums) < 7:
            if re.fullmatch(r"\(?\$?-?\d[\d,]*(?:\.\d+)?\)?", cells[j].replace(",", "")) or re.fullmatch(
                r"\d+", cells[j]
            ):
                tail_nums.append(parse_money(cells[j]))
                j -= 1
            else:
                break
        tail_nums.reverse()
        if len(tail_nums) < 6:
            continue
        base_price, base_freight, ltmf, total_unit, unit_qty_f, line_amt = tail_nums[-6:]
        unit_qty = int(round(unit_qty_f))
        desc_tokens = cells[2 : j + 1]
        desc = " ".join(desc_tokens)
        vol = ""
        mv = re.search(r"(\d+\s*(?:ML|L|oz|OZ))\b", desc, re.I)
        if mv:
            vol = mv.group(1)
        uom = "EA"
        mu = re.search(r"\b(EA|CS|CASE|BTL|BOTTLE)\b", desc, re.I)
        if mu:
            uom = mu.group(1).upper()
        tu = total_unit if total_unit > 0 else base_price + base_freight + ltmf
        la = line_amt if line_amt else tu * unit_qty
        out.append(
            InvoiceItem(
                line=line_no,
                plcb_code=code,
                description=desc,
                unit_volume=vol,
                uom=uom,
                base_price=base_price,
                base_freight=base_freight,
                ltmf=ltmf,
                total_unit_price=tu,
                unit_qty=unit_qty,
                line_amount=la,
            )
        )
    dedup: dict[int, InvoiceItem] = {}
    for it in out:
        dedup[it.line] = it
    return [dedup[k] for k in sorted(dedup)]


def _parse_header_special(text: str) -> dict[str, str]:
    keys = [
        ("pickup_store", r"Pickup\s+Store"),
        ("customer_number", r"Customer\s+Number"),
        ("customer_name", r"Customer\s+Name"),
        ("invoice_number", r"Invoice\s+Number"),
        ("invoice_date", r"Invoice\s+Date"),
        ("order_number", r"Order\s+Number"),
        ("order_date", r"Order\s+Date"),
        ("vendor_number", r"Vendor\s+Number"),
        ("ship_date", r"Ship\s+Date"),
        ("payment_method", r"Payment\s+Method"),
        ("account_last_four", r"Account\s+Last\s+Four"),
        ("delivery_method", r"Delivery\s+Method"),
    ]
    out: dict[str, str] = {}
    for field, pat in keys:
        m = re.search(rf"{pat}\s*[:\s]\s*(.+)$", text, flags=re.IGNORECASE | re.MULTILINE)
        if m:
            out[field] = _norm_cell(m.group(1))
    return out


def parse_special_order_invoice(
    pdf_path: str,
    *,
    text: Optional[str] = None,
    tables: Optional[List[List[Any]]] = None,
) -> InvoiceSummary:
    if text is None:
        text = extract_text(pdf_path)
    if tables is None:
        tables = extract_tables(pdf_path)
    hdr = _parse_header_special(text)

    items = _parse_items_from_tables(tables)
    if not items:
        items = _parse_items_from_text(text)

    summary = InvoiceSummary(
        invoice_type="SPECIAL_ORDER",
        invoice_number=hdr.get("invoice_number", find_label_value(text, "Invoice Number")),
        invoice_date=hdr.get("invoice_date", find_label_value(text, "Invoice Date")),
        order_number=hdr.get("order_number", find_label_value(text, "Order Number")),
        customer_name=hdr.get("customer_name", find_label_value(text, "Customer Name")),
        customer_number=hdr.get("customer_number", find_label_value(text, "Customer Number")),
        pickup_store=hdr.get("pickup_store", ""),
        order_date=hdr.get("order_date", ""),
        vendor_number=hdr.get("vendor_number", ""),
        ship_date=hdr.get("ship_date", ""),
        payment_method=hdr.get("payment_method", ""),
        account_last_four=hdr.get("account_last_four", ""),
        delivery_method=hdr.get("delivery_method", ""),
        items=items,
    )

    summary.total_base_price = find_money_after_label(text, "Total Base Price")
    summary.delivery_fee = find_money_after_label(text, "Delivery Fee")
    summary.ltmf_fee = find_money_after_label(text, "LTMF Fee")
    summary.total_taxable = find_money_after_label(text, "Total Taxable Amount")
    summary.sales_tax_6 = find_money_after_label(text, "Sales Tax (6%)")
    summary.local_tax_2 = find_money_after_label(text, "Local Tax (2%)")
    if summary.local_tax_2 == 0.0:
        summary.local_tax_2 = find_money_after_label(text, "Local Tax")
    summary.base_freight = find_money_after_label(text, "Base Freight")
    summary.total_net_price = find_money_after_label(text, "Total Net Price")

    summary.total_units = sum(i.unit_qty for i in items)
    return summary
