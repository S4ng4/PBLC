from __future__ import annotations

import io
from pathlib import Path
from typing import Iterable, List, Sequence

from models import InvoiceItem, InvoiceSummary

_ASSETS_DIR = Path(__file__).resolve().parent / "assets"
_PLCB_LOGO = "plcb_logo.png"

_PLCB_NAVY = "#003B6F"
_PLCB_GOLD = "#C5A572"
_ZEBRA = "#F1F5F9"
_GRID = "#94A3B8"


def _title_for(summary: InvoiceSummary) -> str:
    if summary.invoice_type == "SPECIAL_ORDER":
        return "SPECIAL ORDER"
    if summary.invoice_type == "LSC_AR":
        return "LSC AR INVOICE"
    return "PLCB WHOLESALE LICENSEE ORDER"


def _units_total(items: Sequence[InvoiceItem]) -> int:
    return int(sum(i.unit_qty for i in items))


def _calc_o_p(summary: InvoiceSummary, it: InvoiceItem, total_units: int) -> tuple[float, float]:
    """
    Restituisce (O, P) coerenti con il foglio Excel `MarginEdge`:
    - Special Order: L=J6/Ktot, M=(J9+J10)*(F*J)/J5, N=M/J, O=I+L+N, P=O*J
    - LSC/Wholesale: L=J6/Ktot, M=(J8+J9)*(F*G)/J5, N=M/G, O=F+L+N, P=O*G
    """
    if total_units <= 0:
        return 0.0, 0.0

    if summary.invoice_type == "SPECIAL_ORDER":
        j5 = summary.total_base_price
        j6 = summary.delivery_fee
        j9 = summary.sales_tax_6
        j10 = summary.local_tax_2
        j = float(it.unit_qty)
        f = it.base_price
        iu = it.total_unit_price
        l_ = (j6 / float(total_units)) if total_units else 0.0
        m = ((j9 + j10) * (f * j) / j5) if j5 else 0.0
        n = (m / j) if j else 0.0
        o = iu + l_ + n
        p = o * j
        return o, p

    j5 = summary.total_gross
    j6 = summary.licensee_discount
    j8 = summary.sales_tax_6
    j9 = summary.local_tax_2
    g = float(it.unit_qty)
    f = it.base_price
    l_ = (j6 / float(total_units)) if total_units else 0.0
    m = ((j8 + j9) * (f * g) / j5) if j5 else 0.0
    n = (m / g) if g else 0.0
    o = f + l_ + n
    p = o * g
    return o, p


def _margin_edge_final_rows(summary: InvoiceSummary) -> List[List[str]]:
    total_units = _units_total(summary.items)
    rows: List[List[str]] = []
    for it in summary.items:
        o, p = _calc_o_p(summary, it, total_units)
        rows.append(
            [
                str(it.line),
                it.plcb_code,
                it.description,
                it.unit_volume,
                it.uom,
                str(int(it.unit_qty)),
                f"{o:.4f}",
                f"{p:.4f}",
            ]
        )
    return rows


def write_margin_edge_final_pdf(summary: InvoiceSummary, dest_pdf_path: str | Path) -> Path:
    """
    Genera un PDF che esporta il contenuto del foglio Excel `MarginEdgeFinal`:
    intestazione stile PLCB + blocco dati ordine + tabella con righe prodotto e colonne A–H.
    """
    try:
        from reportlab.lib import colors
        from reportlab.lib.pagesizes import letter
        from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
        from reportlab.lib.units import inch
        from reportlab.lib.utils import ImageReader
        from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
    except ImportError as exc:  # pragma: no cover
        raise RuntimeError("Installa reportlab: pip install reportlab") from exc

    dst = Path(dest_pdf_path)
    dst.parent.mkdir(parents=True, exist_ok=True)

    page_w, page_h = letter
    left_margin = right_margin = 0.7 * inch
    top_margin = bottom_margin = 0.75 * inch
    usable_w = page_w - left_margin - right_margin

    header_h = 0.78 * inch
    gold_h = 0.25 * inch
    content_top_gap = 0.22 * inch

    styles = getSampleStyleSheet()
    style_normal = styles["Normal"]
    style_label = ParagraphStyle(
        "label",
        parent=style_normal,
        fontName="Helvetica-Bold",
        fontSize=8.5,
        leading=10.0,
        textColor=colors.HexColor("#0f172a"),
    )
    style_small = ParagraphStyle(
        "small",
        parent=style_normal,
        fontName="Helvetica",
        fontSize=8.5,
        leading=10.0,
        textColor=colors.HexColor("#0f172a"),
    )
    style_small_center = ParagraphStyle(
        "small_center",
        parent=style_small,
        alignment=1,
    )

    def _money(v: float) -> str:
        try:
            return f"${float(v):,.2f}"
        except Exception:
            return "$0.00"

    from xml.sax.saxutils import escape as _xml_escape

    def _p(text: object, style) -> Paragraph:
        return Paragraph(_xml_escape(str(text or "")), style)

    hdr = [
        "Line",
        "PLCB Item Code",
        "Description",
        "Unit Vol.",
        "UOM",
        "Unit Qty",
        "Total Unit Price<br/>Up Charge Per Bottle",
        "Total",
    ]

    data: List[List[object]] = [
        [Paragraph(h, style_small_center) for h in hdr],
    ]
    for r in _margin_edge_final_rows(summary):
        row: List[object] = []
        # A, B, D, E, F: testo corto, centrato
        row.append(_p(r[0], style_small_center))
        row.append(_p(r[1], style_small_center))
        # C: descrizione con wrap
        row.append(_p(r[2], style_small))
        row.append(_p(r[3], style_small_center))
        row.append(_p(r[4], style_small_center))
        row.append(_p(r[5], style_small_center))
        # G, H: numeri a 4 decimali, centrati come in Excel
        row.append(_p(r[6], style_small_center))
        row.append(_p(r[7], style_small_center))
        data.append(row)

    # Proporzioni colonne (A..H): Line, Code, Description, Unit Vol., UOM, Qty, Upcharge, Total
    base = [7, 11, 28, 12, 8, 9, 16, 14]
    total = float(sum(base))
    col_widths = [usable_w * (b / total) for b in base]

    def _order_info_table() -> Table:
        """
        Blocco dati ordine (stile 'fattura' PLCB): due colonne di coppie label/valore.
        """
        # Alcuni campi possono essere vuoti a seconda del tipo fattura: mostriamo comunque le label principali.
        left_pairs = [
            ("Customer Name", summary.customer_name or ""),
            ("Customer Number", summary.customer_number or ""),
            ("Order Number", summary.order_number or ""),
            ("Order Date", getattr(summary, "order_date", "") or ""),
        ]
        if summary.invoice_type == "SPECIAL_ORDER":
            left_pairs.extend(
                [
                    ("Pickup Store", getattr(summary, "pickup_store", "") or ""),
                    ("Vendor Number", getattr(summary, "vendor_number", "") or ""),
                    ("Ship Date", getattr(summary, "ship_date", "") or ""),
                    ("Delivery Method", getattr(summary, "delivery_method", "") or ""),
                ]
            )
        else:
            left_pairs.extend(
                [
                    ("Delivery Date", getattr(summary, "delivery_date", "") or ""),
                ]
            )

        if summary.invoice_type == "SPECIAL_ORDER":
            right_pairs = [
                ("Invoice Number", summary.invoice_number or ""),
                ("Invoice Date", summary.invoice_date or ""),
                ("Payment Method", getattr(summary, "payment_method", "") or ""),
                ("Account Last Four", getattr(summary, "account_last_four", "") or ""),
                ("Total Base Price", _money(getattr(summary, "total_base_price", 0.0))),
                ("Delivery Fee", _money(getattr(summary, "delivery_fee", 0.0))),
                ("Sales Tax (6%)", _money(getattr(summary, "sales_tax_6", 0.0))),
                ("Local Tax (2%)", _money(getattr(summary, "local_tax_2", 0.0))),
                ("Total Net Price", _money(summary.total_net_price)),
            ]
        else:
            right_pairs = [
                ("Invoice Number", summary.invoice_number or ""),
                ("Invoice Date", summary.invoice_date or ""),
                ("Payment Method", getattr(summary, "payment_method", "") or ""),
                ("Account Last Four", getattr(summary, "account_last_four", "") or ""),
                ("Total Gross", _money(getattr(summary, "total_gross", 0.0))),
                ("Licensee Discount", _money(getattr(summary, "licensee_discount", 0.0))),
                ("Sales Tax (6%)", _money(getattr(summary, "sales_tax_6", 0.0))),
                ("Local Tax (2%)", _money(getattr(summary, "local_tax_2", 0.0))),
                ("Total Net Price", _money(summary.total_net_price)),
            ]

        # Normalizza le righe (stessa lunghezza per costruire tabella a 4 colonne)
        n = max(len(left_pairs), len(right_pairs))
        left_pairs = left_pairs + [("", "")] * (n - len(left_pairs))
        right_pairs = right_pairs + [("", "")] * (n - len(right_pairs))

        rows: List[List[object]] = []
        for (l_lab, l_val), (r_lab, r_val) in zip(left_pairs, right_pairs):
            rows.append(
                [
                    _p(l_lab, style_label),
                    _p(l_val, style_small),
                    _p(r_lab, style_label),
                    _p(r_val, style_small),
                ]
            )

        t = Table(
            rows,
            colWidths=[usable_w * 0.18, usable_w * 0.32, usable_w * 0.18, usable_w * 0.32],
            hAlign="LEFT",
        )
        t.setStyle(
            TableStyle(
                [
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("LINEBELOW", (0, 0), (-1, -1), 0.25, colors.HexColor("#CBD5E1")),
                    ("LEFTPADDING", (0, 0), (-1, -1), 2),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                    ("TOPPADDING", (0, 0), (-1, -1), 2),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
                ]
            )
        )
        return t

    def _draw_header(canvas, doc) -> None:
        canvas.saveState()
        navy = colors.HexColor(_PLCB_NAVY)
        gold = colors.HexColor(_PLCB_GOLD)

        x0 = left_margin
        y_top = page_h - top_margin

        # Navy band (3 rows in Excel) + right text block
        canvas.setFillColor(navy)
        canvas.rect(x0, y_top - header_h, usable_w, header_h, stroke=0, fill=1)

        # Logo box left side (approx A1:B3)
        logo_w = usable_w * 0.23
        logo_h = header_h
        logo_x = x0
        logo_y = y_top - header_h

        logo_path = _ASSETS_DIR / _PLCB_LOGO
        drew_logo = False
        if logo_path.is_file():
            try:
                img = ImageReader(str(logo_path))
                canvas.drawImage(
                    img,
                    logo_x + 0.10 * inch,
                    logo_y + 0.18 * inch,
                    width=1.55 * inch,
                    height=0.46 * inch,
                    mask="auto",
                    preserveAspectRatio=True,
                    anchor="sw",
                )
                drew_logo = True
            except Exception:
                drew_logo = False
        if not drew_logo:
            canvas.setFillColor(colors.white)
            canvas.setFont("Helvetica-Bold", 18)
            canvas.drawString(logo_x + 0.18 * inch, logo_y + 0.36 * inch, "PLCB")
            canvas.setFont("Helvetica", 9.5)
            canvas.setFillColor(colors.HexColor("#C8D7EB"))
            canvas.drawString(logo_x + 0.18 * inch, logo_y + 0.20 * inch, "MarginEdge Final")

        # Right header text (approx C1:H3)
        canvas.setFillColor(colors.white)
        canvas.setFont("Helvetica-Bold", 11.5)
        header_txt = (
            f"{_title_for(summary)}\n"
            "Gran Caffè L'Aquila\n"
            f"Invoice {summary.invoice_number}  ·  {summary.invoice_date}"
        )
        tx = canvas.beginText()
        tx.setTextOrigin(x0 + logo_w + 0.18 * inch, logo_y + header_h - 0.26 * inch)
        tx.setLeading(13)
        for line in header_txt.splitlines():
            tx.textLine(line)
        canvas.drawText(tx)

        # Gold bar row
        canvas.setFillColor(gold)
        canvas.rect(x0, y_top - header_h - gold_h, usable_w, gold_h, stroke=0, fill=1)
        canvas.setFillColor(colors.HexColor("#1e293b"))
        canvas.setFont("Helvetica-Bold", 9.5)
        canvas.drawCentredString(
            x0 + usable_w / 2.0,
            y_top - header_h - gold_h + 0.08 * inch,
            "Pennsylvania Liquor Control Board — dati per import MarginEdge",
        )

        canvas.restoreState()

    doc = SimpleDocTemplate(
        str(dst),
        pagesize=letter,
        leftMargin=left_margin,
        rightMargin=right_margin,
        topMargin=top_margin,
        bottomMargin=bottom_margin,
        title=f"{summary.invoice_number} MarginEdgeFinal",
        author="PLCB Invoice Processor",
    )

    # Spazio per header (disegnato via onPage) + contenuto in stile fattura
    story: List[object] = [Spacer(1, header_h + gold_h + content_top_gap)]
    story.append(_order_info_table())
    story.append(Spacer(1, 0.18 * inch))

    table = Table(data, colWidths=col_widths, repeatRows=1, hAlign="LEFT")
    navy = colors.HexColor(_PLCB_NAVY)
    zebra = colors.HexColor(_ZEBRA)
    grid = colors.HexColor(_GRID)

    ts = TableStyle(
        [
            ("BACKGROUND", (0, 0), (-1, 0), navy),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("ALIGN", (0, 0), (-1, 0), "CENTER"),
            ("LINEBELOW", (0, 0), (-1, 0), 0.8, grid),
            ("GRID", (0, 0), (-1, -1), 0.6, grid),
            ("LEFTPADDING", (0, 0), (-1, -1), 4),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ]
    )
    # Zebra striping (skip header row)
    for i in range(1, len(data)):
        if i % 2 == 0:
            ts.add("BACKGROUND", (0, i), (-1, i), zebra)
    table.setStyle(ts)

    story.append(table)
    doc.build(story, onFirstPage=_draw_header, onLaterPages=_draw_header)
    return dst

