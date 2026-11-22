from __future__ import annotations

import io
import re
from pathlib import Path
from typing import Optional

from models import InvoiceSummary
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins
from openpyxl.worksheet.properties import PageSetupProperties

_ASSETS_DIR = Path(__file__).resolve().parent.parent / "assets"
_LOGO_FILENAME = "plcb_logo.png"
_PLCB_NAVY = "003B6F"
_PLCB_GOLD = "C5A572"
_HEADER_WHITE_LG = Font(color="FFFFFF", bold=True, size=22)
_TABLE_HEADER_FONT = Font(color="FFFFFF", bold=True, size=11)
_THIN = Side(style="thin", color="94A3B8")
_BORDER_TABLE = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)


def _safe_filename_stem(s: str) -> str:
    t = (s or "UNKNOWN").strip()
    t = re.sub(r"[^\w\-]+", "_", t, flags=re.UNICODE)
    return t or "UNKNOWN"


MONEY_FMT = '"$"#,##0.00'
DEC4_FMT = "0.0000"


def _title_for(summary: InvoiceSummary) -> str:
    if summary.invoice_type == "SPECIAL_ORDER":
        return "SPECIAL ORDER"
    if summary.invoice_type == "LSC_AR":
        return "LSC AR INVOICE"
    return "PLCB WHOLESALE LICENSEE ORDER"


def _apply_money(cell) -> None:
    cell.number_format = MONEY_FMT


def _apply_dec4(cell) -> None:
    cell.number_format = DEC4_FMT


def _fill_gca(ws, text: str) -> None:
    for i, line in enumerate(text.splitlines(), start=1):
        ws.cell(row=i, column=1, value=line)


def _fill_margin_header(ws, summary: InvoiceSummary) -> tuple[int, int, int]:
    """Returns (first_product_row, last_product_row, total_row)."""
    ws["E2"] = _title_for(summary)
    ws["I2"] = summary.invoice_date

    labels = [
        ("Customer Name", summary.customer_name),
        ("Customer Number", summary.customer_number),
        ("Invoice Number", summary.invoice_number),
        ("Invoice Date", summary.invoice_date),
        ("Order Number", summary.order_number),
        ("Order Date", getattr(summary, "order_date", "") or ""),
    ]
    if summary.invoice_type == "SPECIAL_ORDER":
        labels.extend(
            [
                ("Pickup Store", summary.pickup_store),
                ("Vendor Number", summary.vendor_number),
                ("Ship Date", summary.ship_date),
                ("Payment Method", summary.payment_method),
                ("Account Last Four", summary.account_last_four),
                ("Delivery Method", summary.delivery_method),
            ]
        )
    else:
        labels.extend(
            [
                ("Delivery Date", summary.delivery_date),
                ("Payment Method", summary.payment_method),
                ("Account Last Four", summary.account_last_four),
            ]
        )

    r = 5
    for lab, val in labels:
        if r > 14:
            break
        ws.cell(row=r, column=1, value=lab)
        ws.cell(row=r, column=2, value=val)
        r += 1

    if summary.invoice_type == "SPECIAL_ORDER":
        ws["J5"] = summary.total_base_price
        ws["J6"] = summary.delivery_fee
        ws["J7"] = summary.ltmf_fee
        ws["J8"] = summary.total_taxable
        ws["J9"] = summary.sales_tax_6
        ws["J10"] = summary.local_tax_2
        ws["J11"] = summary.base_freight
        ws["J12"] = summary.total_net_price
        for rr in range(5, 13):
            _apply_money(ws[f"J{rr}"])
    else:
        ws["J5"] = summary.total_gross
        ws["J6"] = summary.licensee_discount
        ws["J7"] = summary.taxable_amount
        ws["J8"] = summary.sales_tax_6
        ws["J9"] = summary.local_tax_2
        ws["J10"] = summary.total_net_price
        for rr in range(5, 11):
            _apply_money(ws[f"J{rr}"])

    hdr = [
        "Line",
        "PLCB Item Code",
        "Description",
        "Unit Volume",
        "UOM" if summary.invoice_type == "SPECIAL_ORDER" else "Case Pack",
        "Base Price" if summary.invoice_type == "SPECIAL_ORDER" else "Unit Price",
        "Base Freight" if summary.invoice_type == "SPECIAL_ORDER" else "Invoice Qty",
        "LTMF" if summary.invoice_type == "SPECIAL_ORDER" else "UOM",
        "Total Unit Price",
        "Unit Qty",
        "Line Amount",
        "Delivery / Discount per bottle",
        "Tax alloc",
        "Tax per bottle",
        "Upcharge / bottle",
        "Line total",
    ]
    for c, val in enumerate(hdr, start=1):
        cell = ws.cell(row=15, column=c, value=val)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", wrap_text=True)

    first = 16
    last = 15 + len(summary.items)
    total_row = 16 + len(summary.items)

    for col in range(1, 17):
        ws.column_dimensions[get_column_letter(col)].width = 14

    return first, last, total_row


def _fill_products_margin_edge(ws, summary: InvoiceSummary, first: int, last: int, total_row: int) -> None:
    for idx, it in enumerate(summary.items):
        r = first + idx
        ws.cell(row=r, column=1, value=it.line)
        ws.cell(row=r, column=2, value=it.plcb_code)
        ws.cell(row=r, column=3, value=it.description)
        ws.cell(row=r, column=4, value=it.unit_volume)
        if summary.invoice_type == "SPECIAL_ORDER":
            ws.cell(row=r, column=5, value=it.uom)
            ws.cell(row=r, column=6, value=it.base_price)
            ws.cell(row=r, column=7, value=it.base_freight)
            ws.cell(row=r, column=8, value=it.ltmf)
            ws.cell(row=r, column=9, value=it.total_unit_price)
            ws.cell(row=r, column=10, value=it.unit_qty)
            ws.cell(row=r, column=11, value=f"=I{r}*J{r}")
            ws[f"L{r}"] = f"=$J$6/$K${total_row}"
            ws[f"M{r}"] = f"=($J$9+$J$10)*(F{r}*J{r})/$J$5"
            ws[f"N{r}"] = f"=M{r}/J{r}"
            ws[f"O{r}"] = f"=I{r}+L{r}+N{r}"
            ws[f"P{r}"] = f"=O{r}*J{r}"
            for c in (6, 7, 8, 9, 11):
                _apply_money(ws.cell(row=r, column=c))
            ws.cell(row=r, column=10).number_format = "0"
        else:
            pack = it.case_pack if it.case_pack is not None else ""
            ws.cell(row=r, column=5, value=pack)
            ws.cell(row=r, column=6, value=it.base_price)
            ws.cell(row=r, column=7, value=it.unit_qty)
            ws.cell(row=r, column=8, value=it.uom)
            ws.cell(row=r, column=9, value=f"=F{r}")
            ws[f"J{r}"] = f"=G{r}"
            ws.cell(row=r, column=11, value=f"=F{r}*G{r}")
            ws[f"L{r}"] = f"=$J$6/$K${total_row}"
            ws[f"M{r}"] = f"=($J$8+$J$9)*(F{r}*G{r})/$J$5"
            ws[f"N{r}"] = f"=M{r}/G{r}"
            ws[f"O{r}"] = f"=F{r}+L{r}+N{r}"
            ws[f"P{r}"] = f"=O{r}*G{r}"
            _apply_money(ws.cell(row=r, column=6))
            _apply_money(ws.cell(row=r, column=11))
            ws.cell(row=r, column=7).number_format = "0"
            ws.cell(row=r, column=9).number_format = MONEY_FMT
            ws.cell(row=r, column=10).number_format = "0"

        _apply_dec4(ws.cell(row=r, column=15))
        _apply_dec4(ws.cell(row=r, column=16))

    ws.cell(row=total_row, column=1, value="Total Units:")
    k_col = 11
    if last >= first:
        if summary.invoice_type == "SPECIAL_ORDER":
            ws.cell(row=total_row, column=k_col, value=f"=SUM(J{first}:J{last})")
        else:
            ws.cell(row=total_row, column=k_col, value=f"=SUM(G{first}:G{last})")
        ws.cell(row=total_row, column=16, value=f"=SUM(P{first}:P{last})")
    else:
        ws.cell(row=total_row, column=k_col, value=0)
        ws.cell(row=total_row, column=16, value=0.0)
    ws.cell(row=total_row, column=k_col).number_format = "0"
    _apply_dec4(ws.cell(row=total_row, column=16))


def _plcb_header_fill() -> PatternFill:
    return PatternFill(fill_type="solid", fgColor=_PLCB_NAVY)


def _logo_png_from_pillow() -> Optional[bytes]:
    try:
        from PIL import Image, ImageDraw, ImageFont
    except ImportError:
        return None
    w, h = 260, 54
    im = Image.new("RGBA", (w, h), (0, 59, 111, 255))
    dr = ImageDraw.Draw(im)
    try:
        font_big = ImageFont.truetype("/System/Library/Fonts/Supplemental/Arial Bold.ttf", 26)
        font_sm = ImageFont.truetype("/System/Library/Fonts/Supplemental/Arial.ttf", 11)
    except OSError:
        try:
            font_big = ImageFont.truetype("/Library/Fonts/Arial.ttf", 26)
            font_sm = ImageFont.truetype("/Library/Fonts/Arial.ttf", 11)
        except OSError:
            font_big = ImageFont.load_default()
            font_sm = font_big
    dr.text((14, 6), "PLCB", fill=(255, 255, 255, 255), font=font_big)
    dr.text((14, 36), "MarginEdge Final", fill=(200, 215, 235, 255), font=font_sm)
    buf = io.BytesIO()
    im.save(buf, format="PNG")
    return buf.getvalue()


def _logo_png_bytes() -> Optional[bytes]:
    path = _ASSETS_DIR / _LOGO_FILENAME
    if path.is_file():
        try:
            return path.read_bytes()
        except OSError:
            pass
    return _logo_png_from_pillow()


def _apply_margin_edge_final_plcb_layout(dst, summary: InvoiceSummary) -> None:
    """Righe 1–3: fascia istituzionale blu + logo; non sposta le righe dati (da riga 5 come MarginEdge)."""
    navy = _plcb_header_fill()
    gold = PatternFill(fill_type="solid", fgColor=_PLCB_GOLD)

    for r in range(1, 5):
        for c in range(1, 18):
            dst.cell(row=r, column=c).value = None

    dst.merge_cells("A1:B3")
    dst.merge_cells("C1:H3")

    for r in range(1, 4):
        for c in range(1, 9):
            dst.cell(row=r, column=c).fill = navy

    dst.row_dimensions[1].height = 20
    dst.row_dimensions[2].height = 20
    dst.row_dimensions[3].height = 20

    left = dst["A1"]
    left.value = None
    left.alignment = Alignment(horizontal="center", vertical="center")
    left.font = _HEADER_WHITE_LG

    right = dst["C1"]
    right.value = (
        f"{_title_for(summary)}\n"
        "Gran Caffè L'Aquila\n"
        f"Invoice {summary.invoice_number}  ·  {summary.invoice_date}"
    )
    right.font = Font(color="FFFFFF", bold=True, size=12)
    right.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

    png = _logo_png_bytes()
    if png:
        try:
            img = XLImage(io.BytesIO(png))
            img.width = 150
            img.height = 44
            dst.add_image(img, "A1")
        except Exception:
            left.value = "PLCB"
    else:
        left.value = "PLCB"

    dst.merge_cells("A4:H4")
    bar = dst["A4"]
    for c in range(1, 9):
        dst.cell(row=4, column=c).fill = gold
    bar.value = "Pennsylvania Liquor Control Board — dati per import MarginEdge"
    bar.font = Font(color="1e293b", bold=True, size=10)
    bar.alignment = Alignment(horizontal="center", vertical="center")
    dst.row_dimensions[4].height = 18

    dst.column_dimensions["A"].width = 7
    dst.column_dimensions["B"].width = 12
    for col, w in (("C", 32), ("D", 12), ("E", 8), ("F", 10), ("G", 18), ("H", 14)):
        dst.column_dimensions[col].width = w


def _fill_margin_edge_final(wb, summary: InvoiceSummary, first: int, last: int) -> None:
    src = wb["MarginEdge"]
    dst = wb["MarginEdgeFinal"]
    for r in range(1, 16):
        for c in range(1, 17):
            dst.cell(row=r, column=c, value=src.cell(row=r, column=c).value)

    _apply_margin_edge_final_plcb_layout(dst, summary)

    hdr = [
        "Line",
        "PLCB Item Code",
        "Description",
        "Unit Vol.",
        "UOM",
        "Unit Qty",
        "Total Unit Price\nUp Charge Per Bottle",
        "Total",
    ]
    navy = _plcb_header_fill()
    for c, val in enumerate(hdr, start=1):
        cell = dst.cell(row=15, column=c, value=val)
        cell.font = _TABLE_HEADER_FONT
        cell.fill = navy
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = _BORDER_TABLE

    uom_col = "E" if summary.invoice_type == "SPECIAL_ORDER" else "H"
    zebra = PatternFill(fill_type="solid", fgColor="F1F5F9")
    for idx in range(last - first + 1):
        r = 16 + idx
        src_r = 16 + idx
        dst.cell(row=r, column=1, value=f"=MarginEdge!A{src_r}")
        dst.cell(row=r, column=2, value=f"=MarginEdge!B{src_r}")
        dst.cell(row=r, column=3, value=f"=MarginEdge!C{src_r}")
        dst.cell(row=r, column=4, value=f"=MarginEdge!D{src_r}")
        dst.cell(row=r, column=5, value=f"=MarginEdge!{uom_col}{src_r}")
        dst.cell(row=r, column=6, value=f"=MarginEdge!J{src_r}")
        dst.cell(row=r, column=7, value=f"=MarginEdge!O{src_r}")
        dst.cell(row=r, column=8, value=f"=MarginEdge!P{src_r}")
        _apply_dec4(dst.cell(row=r, column=7))
        _apply_dec4(dst.cell(row=r, column=8))
        for c in range(1, 9):
            cell = dst.cell(row=r, column=c)
            cell.border = _BORDER_TABLE
            if idx % 2:
                cell.fill = zebra

    dst.freeze_panes = "A16"


def _apply_us_letter_portrait(ws, print_area: str | None = None) -> None:
    """
    Carta US Letter (8.5\"×11\") in verticale, margini tipici da stampante desktop.
    fitToWidth=1 adatta tutte le colonne dell’area di stampa alla larghezza utile di una pagina
    (come “adatta tutte le colonne su una pagina” in Excel).
    """
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True, autoPageBreaks=True)
    ws.page_setup.paperSize = 1  # 1 = US Letter (Excel)
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_margins = PageMargins(
        left=0.7,
        right=0.7,
        top=0.75,
        bottom=0.75,
        header=0.3,
        footer=0.3,
    )
    if print_area:
        ws.print_area = print_area


def build_and_save_excel(
    summary: InvoiceSummary, pdf_path: str, raw_text: str, output_path: Optional[str] = None
) -> Path:
    wb = Workbook()
    gca = wb.active
    gca.title = "GCA"
    _fill_gca(gca, raw_text)

    me = wb.create_sheet("MarginEdge")
    first, last, total_row = _fill_margin_header(me, summary)
    _fill_products_margin_edge(me, summary, first, last, total_row)

    wb.create_sheet("MarginEdgeFinal")
    _fill_margin_edge_final(wb, summary, first, last)

    n = len(summary.items)
    lines = max(1, len(raw_text.splitlines()))
    _apply_us_letter_portrait(gca, f"A1:A{lines}")
    _apply_us_letter_portrait(me, f"A1:P{total_row}")
    _apply_us_letter_portrait(wb["MarginEdgeFinal"], f"A1:H{15 + n}")

    inv = _safe_filename_stem(summary.invoice_number or "UNKNOWN")
    dest = Path(output_path) if output_path else Path(pdf_path).resolve().parent / f"{inv}_MarginEdge.xlsx"
    wb.save(dest)
    return dest
