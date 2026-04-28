from __future__ import annotations

from typing import Any, List


def extract_text(pdf_path: str) -> str:
    try:
        import pdfplumber
    except ImportError as exc:  # pragma: no cover
        raise RuntimeError("Install pdfplumber to read PDFs.") from exc

    with pdfplumber.open(pdf_path) as pdf:
        parts: List[str] = []
        for page in pdf.pages:
            t = page.extract_text() or ""
            parts.append(t)
        return "\n".join(parts)


def extract_tables(pdf_path: str) -> List[List[Any]]:
    try:
        import pdfplumber
    except ImportError as exc:  # pragma: no cover
        raise RuntimeError("Install pdfplumber to read PDFs.") from exc

    tables: List[List[Any]] = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            for tbl in page.extract_tables() or []:
                tables.append(tbl)
    return tables
