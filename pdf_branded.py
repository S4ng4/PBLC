"""
PDF con lo stesso contenuto del PLCB originale, più logo Gran Caffè L'Aquila in sovrapposizione.

Usa PyMuPDF: non ri-layoutizza il documento (testo e grafica restano quelli del file sorgente).
"""
from __future__ import annotations

import io
from pathlib import Path
from typing import Optional

_ASSETS_DIR = Path(__file__).resolve().parent / "assets"
_GCA_LOGO_PNG = "gca_logo.png"
_GCA_LOGO_BUILTIN_JPEG = "gca_logo_builtin.jpg"

# Angolo in alto a destra: di solito il marchio PLCB è a sinistra; così si limita la sovrapposizione.
_LOGO_MAX_WIDTH_FRAC = 0.24
_MARGIN_PT = 20.0


def _gca_logo_png_from_pillow() -> Optional[bytes]:
    try:
        from PIL import Image, ImageDraw, ImageFont
    except ImportError:
        return None
    w, h = 300, 58
    im = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    dr = ImageDraw.Draw(im)
    try:
        font_big = ImageFont.truetype("/System/Library/Fonts/Supplemental/Arial Bold.ttf", 22)
        font_sm = ImageFont.truetype("/System/Library/Fonts/Supplemental/Arial.ttf", 14)
    except OSError:
        try:
            font_big = ImageFont.truetype("/Library/Fonts/Arial.ttf", 22)
            font_sm = ImageFont.truetype("/Library/Fonts/Arial.ttf", 14)
        except OSError:
            font_big = ImageFont.load_default()
            font_sm = font_big
    # Testo scuro su sfondo trasparente (leggibile sul bianco tipico del PDF PLCB)
    dr.text((12, 6), "Gran Caffè", fill=(15, 118, 110, 255), font=font_big)
    dr.text((12, 32), "L'Aquila", fill=(51, 65, 85, 255), font=font_sm)
    buf = io.BytesIO()
    im.save(buf, format="PNG")
    return buf.getvalue()


def gca_logo_png_bytes() -> Optional[bytes]:
    """
    Restituisce i byte di un'immagine logo (PNG o JPEG) per PyMuPDF ``insert_image``.

    Ordine: ``assets/gca_logo.png`` (override manuale), ``assets/gca_logo_builtin.jpg``
    (logo predefinito incluso nel progetto), altrimenti wordmark generato con Pillow.
    """
    for name in (_GCA_LOGO_PNG, _GCA_LOGO_BUILTIN_JPEG):
        path = _ASSETS_DIR / name
        if path.is_file():
            try:
                return path.read_bytes()
            except OSError:
                pass
    return _gca_logo_png_from_pillow()


def write_gca_branded_pdf(source_pdf_path: str | Path, dest_pdf_path: str | Path) -> Path:
    """
    Copia il PDF sorgente in `dest_pdf_path` inserendo il logo GCA su ogni pagina (angolo alto-destra).
    """
    try:
        import fitz
    except ImportError as exc:  # pragma: no cover
        raise RuntimeError("Installa pymupdf: pip install pymupdf") from exc

    logo = gca_logo_png_bytes()
    if not logo:
        raise RuntimeError(
            "Logo Gran Caffè L'Aquila non disponibile (manca assets/gca_logo_builtin.jpg e Pillow)."
        )

    src = Path(source_pdf_path)
    dst = Path(dest_pdf_path)
    try:
        from PIL import Image
    except ImportError as exc:
        raise RuntimeError("Pillow è necessario per calcolare le proporzioni del logo.") from exc

    im = Image.open(io.BytesIO(logo))
    if im.mode not in ("RGB", "RGBA"):
        im = im.convert("RGBA")
    iw, ih = im.size
    if iw <= 0 or ih <= 0:
        raise ValueError("Immagine logo non valida.")
    # PNG per insert_image: alcuni viewer/PDF hanno problemi con JPEG (DCT) in overlay su certe pagine.
    _png_buf = io.BytesIO()
    im.save(_png_buf, format="PNG")
    logo_png = _png_buf.getvalue()

    doc = fitz.open(str(src))
    try:
        for page in doc:
            pw = page.rect.width
            ph = page.rect.height
            max_w = pw * _LOGO_MAX_WIDTH_FRAC
            max_h = ph * 0.12
            scale = min(max_w / float(iw), max_h / float(ih))
            w = float(iw) * scale
            h = float(ih) * scale
            x0 = pw - _MARGIN_PT - w
            y0 = _MARGIN_PT
            if x0 < _MARGIN_PT:
                x0 = _MARGIN_PT
            rect = fitz.Rect(x0, y0, x0 + w, y0 + h)
            page.insert_image(rect, stream=logo_png, overlay=True, keep_proportion=True)
        dst.parent.mkdir(parents=True, exist_ok=True)
        doc.save(str(dst), garbage=2, deflate=True, clean=False)
    finally:
        doc.close()
    return dst
