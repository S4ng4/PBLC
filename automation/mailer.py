from __future__ import annotations

import smtplib
import ssl
from email.message import EmailMessage
from pathlib import Path
from typing import Iterable, Optional, Sequence

from models import InvoiceSummary

from .config import AutomationConfig

AUTO_HEADER_NAME = "X-PLCB-Auto-Processed"
AUTO_HEADER_VALUE = "1"


def _guess_mime(path: Path) -> tuple[str, str]:
    suffix = path.suffix.lower()
    if suffix == ".pdf":
        return "application", "pdf"
    if suffix in (".xlsx", ".xlsm"):
        return "application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    return "application", "octet-stream"


def _attach_file(msg: EmailMessage, file_path: Path, filename: Optional[str] = None) -> None:
    maintype, subtype = _guess_mime(file_path)
    msg.add_attachment(
        file_path.read_bytes(),
        maintype=maintype,
        subtype=subtype,
        filename=filename or file_path.name,
    )


def _format_subject(template: str, **kwargs: object) -> str:
    try:
        return template.format(**kwargs)
    except Exception:
        inv = kwargs.get("invoice_number", "PLCB")
        return f"PLCB Invoice {inv}"


def _summary_context(summary: InvoiceSummary, sum_p: float) -> dict[str, object]:
    return {
        "invoice_number": summary.invoice_number or "UNKNOWN",
        "invoice_type": summary.invoice_type or "",
        "invoice_date": summary.invoice_date or "",
        "order_number": summary.order_number or "",
        "customer_name": summary.customer_name or "",
        "total_net": float(summary.total_net_price),
        "sum_p": float(sum_p),
        "n_items": len(summary.items),
        "units": sum(i.unit_qty for i in summary.items),
    }


def _create_base_message(
    cfg: AutomationConfig,
    recipients: Sequence[str],
    subject: str,
    body: str,
    attachments: Iterable[tuple[Path, Optional[str]]] = (),
) -> EmailMessage:
    msg = EmailMessage()
    msg["From"] = cfg.smtp_from or cfg.smtp_user or cfg.imap_user
    msg["To"] = ", ".join(recipients)
    msg["Subject"] = subject
    msg[AUTO_HEADER_NAME] = AUTO_HEADER_VALUE
    msg.set_content(body, subtype="plain", charset="utf-8")
    for path, custom_name in attachments:
        if path and path.is_file():
            _attach_file(msg, path, custom_name)
    return msg


def build_pdf_only_message(
    cfg: AutomationConfig,
    summary: InvoiceSummary,
    sum_p: float,
    pdf_gca_path: Path,
) -> EmailMessage:
    ctx = _summary_context(summary, sum_p)
    subject = _format_subject(cfg.pdf_only_subject_template, **ctx)
    body = (
        f"Fattura PLCB elaborata automaticamente (MarginEdgeFinal).\n\n"
        f"Invoice Number: {ctx['invoice_number']}\n"
        f"Tipo: {ctx['invoice_type']}\n"
        f"Data: {ctx['invoice_date']}\n"
        f"Order Number: {ctx['order_number']}\n"
        f"Prodotti: {ctx['n_items']} — Bottiglie: {ctx['units']}\n"
        f"Total Net Price: ${ctx['total_net']:,.2f}\n"
    )
    return _create_base_message(
        cfg,
        cfg.pdf_only_recipients,
        subject,
        body,
        attachments=[(pdf_gca_path, None)],
    )


def build_full_report_message(
    cfg: AutomationConfig,
    summary: InvoiceSummary,
    sum_p: float,
    pdf_gca_path: Path,
    excel_path: Path,
    original_pdf_path: Optional[Path] = None,
    original_pdf_name: Optional[str] = None,
) -> EmailMessage:
    ctx = _summary_context(summary, sum_p)
    subject = _format_subject(cfg.full_subject_template, **ctx)
    body = (
        f"Elaborazione automatica fattura PLCB completata con successo (Verifica OK).\n\n"
        f"Invoice Number: {ctx['invoice_number']}\n"
        f"Tipo: {ctx['invoice_type']}\n"
        f"Data fattura: {ctx['invoice_date']}\n"
        f"Order Number: {ctx['order_number']}\n"
        f"Cliente: {ctx['customer_name']}\n"
        f"Prodotti: {ctx['n_items']} — Bottiglie: {ctx['units']}\n"
        f"Total Net Price (PDF): ${ctx['total_net']:,.2f}\n"
        f"Totale calcolato MarginEdge: ${ctx['sum_p']:,.2f}\n\n"
        f"In allegato trovi il PDF MarginEdgeFinal e il file Excel completo.\n"
    )
    attachments: list[tuple[Path, Optional[str]]] = [
        (pdf_gca_path, None),
        (excel_path, None),
    ]
    if cfg.include_original_pdf_in_full and original_pdf_path is not None:
        attachments.append((original_pdf_path, original_pdf_name))
    return _create_base_message(
        cfg,
        cfg.full_recipients,
        subject,
        body,
        attachments=attachments,
    )


def build_mismatch_alert_message(
    cfg: AutomationConfig,
    summary: InvoiceSummary,
    sum_p: float,
    diff: float,
    source_subject: str,
    source_sender: str,
    original_pdf_path: Path,
    excel_path: Path,
    pdf_gca_path: Path,
    original_pdf_name: Optional[str] = None,
) -> EmailMessage:
    ctx = _summary_context(summary, sum_p)
    subject = _format_subject(cfg.alert_subject_template, **ctx)
    body = (
        "ATTENZIONE: invio automatico BLOCCATO per discrepanza nei totali (> $0.01).\n\n"
        f"Email originale da: {source_sender}\n"
        f"Oggetto originale: {source_subject}\n\n"
        f"Invoice Number: {ctx['invoice_number']}\n"
        f"Tipo: {ctx['invoice_type']}\n"
        f"Data fattura: {ctx['invoice_date']}\n"
        f"Prodotti: {ctx['n_items']} — Bottiglie: {ctx['units']}\n"
        f"Total Net Price (PDF): ${ctx['total_net']:,.2f}\n"
        f"Totale calcolato MarginEdge: ${ctx['sum_p']:,.2f}\n"
        f"Differenza (delta): ${diff:,.4f}\n\n"
        "Nessuna email è stata inviata ai destinatari automatici (PDF_ONLY / FULL).\n"
        "In allegato trovi il PDF originale, l'Excel e il PDF generato per il controllo manuale.\n"
    )
    return _create_base_message(
        cfg,
        cfg.alert_recipients,
        subject,
        body,
        attachments=[
            (original_pdf_path, original_pdf_name),
            (excel_path, None),
            (pdf_gca_path, None),
        ],
    )


def build_error_alert_message(
    cfg: AutomationConfig,
    filename: str,
    error_msg: str,
    source_subject: str,
    source_sender: str,
    original_pdf_path: Optional[Path] = None,
    original_pdf_name: Optional[str] = None,
) -> EmailMessage:
    subject = _format_subject(
        cfg.alert_subject_template,
        invoice_number=filename or "PDF",
        invoice_type="ERROR",
        invoice_date="",
        order_number="",
        customer_name="",
        total_net=0.0,
        sum_p=0.0,
        n_items=0,
        units=0,
    )
    body = (
        "ATTENZIONE: errore durante l'elaborazione automatica di una fattura PLCB.\n\n"
        f"Email originale da: {source_sender}\n"
        f"Oggetto originale: {source_subject}\n"
        f"File allegato: {filename}\n"
        f"Errore: {error_msg}\n\n"
        "Nessuna email è stata inviata ai destinatari automatici (PDF_ONLY / FULL).\n"
        "Verifica l'allegato originale qui incluso.\n"
    )
    attachments: list[tuple[Path, Optional[str]]] = []
    if original_pdf_path is not None:
        attachments.append((original_pdf_path, original_pdf_name or filename))
    return _create_base_message(
        cfg,
        cfg.alert_recipients,
        subject,
        body,
        attachments=attachments,
    )


def send_email_message(cfg: AutomationConfig, msg: EmailMessage) -> None:
    if cfg.dry_run:
        return
    context = ssl.create_default_context()
    if cfg.smtp_port == 465:
        with smtplib.SMTP_SSL(cfg.smtp_host, cfg.smtp_port, context=context, timeout=30) as server:
            server.login(cfg.smtp_user, cfg.smtp_password)
            server.send_message(msg)
    else:
        with smtplib.SMTP(cfg.smtp_host, cfg.smtp_port, timeout=30) as server:
            server.ehlo()
            server.starttls(context=context)
            server.ehlo()
            server.login(cfg.smtp_user, cfg.smtp_password)
            server.send_message(msg)
