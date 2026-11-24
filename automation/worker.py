from __future__ import annotations

import email
import imaplib
import logging
import shutil
import signal
import tempfile
import time
from dataclasses import dataclass
from email.header import decode_header
from email.message import EmailMessage, Message
from email.utils import getaddresses, parseaddr
from pathlib import Path
from typing import Callable, List, Optional, Tuple

from excel.builder import build_and_save_excel
from pdf_margin_edge_final import write_margin_edge_final_pdf
from processing import safe_invoice_filename, verify_invoice

from .config import AutomationConfig
from .mailer import (
    AUTO_HEADER_NAME,
    build_error_alert_message,
    build_full_report_message,
    build_mismatch_alert_message,
    build_pdf_only_message,
    send_email_message,
)

logger = logging.getLogger("plcb.automation")


@dataclass
class RunReport:
    checked_messages: int = 0
    skipped_messages: int = 0
    processed_messages: int = 0
    invoices_ok: int = 0
    invoices_mismatch: int = 0
    invoices_error: int = 0
    emails_sent: int = 0
    transient_failures: int = 0


def decode_mime_header(value: Optional[str]) -> str:
    if not value:
        return ""
    parts: list[str] = []
    for chunk, charset in decode_header(value):
        if isinstance(chunk, bytes):
            enc = charset or "utf-8"
            try:
                parts.append(chunk.decode(enc, errors="replace"))
            except LookupError:
                parts.append(chunk.decode("utf-8", errors="replace"))
        else:
            parts.append(str(chunk))
    return "".join(parts).strip()


def _quote_imap_mailbox(name: str) -> str:
    escaped = name.replace("\\", "\\\\").replace('"', '\\"')
    return f'"{escaped}"'


def _sender_matches(sender_email: str, allowed_patterns: Tuple[str, ...]) -> bool:
    if not allowed_patterns:
        return True
    s = sender_email.strip().lower()
    if not s:
        return False
    for raw_pat in allowed_patterns:
        pat = raw_pat.strip().lower()
        if not pat:
            continue
        if "@" in pat:
            if pat.startswith("@"):
                if s.endswith(pat):
                    return True
            elif s == pat:
                return True
        else:
            if s.endswith(f"@{pat}") or s == pat:
                return True
    return False


def _recipient_matches(msg: Message, to_filter: str) -> bool:
    target = (to_filter or "").strip().lower()
    if not target:
        return True
    header_vals = []
    for hdr in ("To", "Cc", "Delivered-To", "X-Original-To"):
        header_vals.extend(msg.get_all(hdr, []))
    parsed = [addr.strip().lower() for _, addr in getaddresses(header_vals) if addr]
    if any(target == addr for addr in parsed):
        return True
    return any(target in decode_mime_header(v).lower() for v in header_vals)


def should_process_message(msg: Message, cfg: AutomationConfig) -> Tuple[bool, str]:
    if msg.get(AUTO_HEADER_NAME):
        return False, "anti-loop header presente"

    sender_raw = decode_mime_header(msg.get("From"))
    _, sender_email = parseaddr(sender_raw)
    if not _sender_matches(sender_email, cfg.imap_allowed_senders):
        return False, f"mittente non in whitelist ({sender_email or sender_raw})"

    if not _recipient_matches(msg, cfg.imap_to_filter):
        return False, f"destinatario non corrisponde a {cfg.imap_to_filter}"

    if cfg.imap_subject_contains:
        subject = decode_mime_header(msg.get("Subject"))
        if cfg.imap_subject_contains.lower() not in subject.lower():
            return False, "subject non corrisponde al filtro"

    return True, "ok"


def extract_pdf_attachments(msg: Message) -> List[Tuple[str, bytes]]:
    attachments: List[Tuple[str, bytes]] = []
    idx = 0
    for part in msg.walk():
        if part.is_multipart():
            continue
        raw_name = part.get_filename()
        filename = decode_mime_header(raw_name) if raw_name else ""
        ctype = (part.get_content_type() or "").lower()
        is_pdf = filename.lower().endswith(".pdf") or ctype == "application/pdf"
        if not is_pdf:
            continue
        payload = part.get_payload(decode=True)
        if not payload:
            continue
        idx += 1
        clean_name = Path(filename).name if filename else f"attachment_{idx}.pdf"
        if not clean_name.lower().endswith(".pdf"):
            clean_name = f"{clean_name}.pdf"
        attachments.append((clean_name, payload))
    return attachments


def _ensure_mailbox_exists(imap: imaplib.IMAP4_SSL, folder: str) -> None:
    if not folder:
        return
    try:
        imap.create(_quote_imap_mailbox(folder))
    except Exception:
        # Molti server IMAP restituiscono errore se la cartella esiste già.
        pass


def _mark_and_move_uid(
    imap: imaplib.IMAP4_SSL,
    uid: bytes,
    target_folder: str,
) -> None:
    imap.uid("STORE", uid, "+FLAGS", r"(\Seen)")
    if not target_folder:
        return
    _ensure_mailbox_exists(imap, target_folder)
    status, _ = imap.uid("COPY", uid, _quote_imap_mailbox(target_folder))
    if status == "OK":
        imap.uid("STORE", uid, "+FLAGS", r"(\Deleted)")
        imap.expunge()
    else:
        logger.warning(
            "Impossibile spostare UID %s in %s (status=%s); messaggio lasciato marcato come letto.",
            uid.decode(errors="replace"),
            target_folder,
            status,
        )


def _archive_outputs(
    archive_dir: Optional[Path],
    inv_stem: str,
    excel_path: Path,
    pdf_gca_path: Path,
    original_pdf_path: Path,
) -> None:
    if archive_dir is None:
        return
    dest_dir = archive_dir / inv_stem
    dest_dir.mkdir(parents=True, exist_ok=True)
    shutil.copy2(excel_path, dest_dir / excel_path.name)
    shutil.copy2(pdf_gca_path, dest_dir / pdf_gca_path.name)
    shutil.copy2(original_pdf_path, dest_dir / f"{inv_stem}_fattura_PLCB.pdf")


def _dispatch_email(
    cfg: AutomationConfig,
    send_fn: Callable[[AutomationConfig, EmailMessage], None],
    kind: str,
    msg: EmailMessage,
) -> None:
    if cfg.dry_run:
        logger.info(
            "[DRY-RUN] Invio %s a %s | Oggetto: %s",
            kind,
            msg["To"],
            msg["Subject"],
        )
        return
    send_fn(cfg, msg)
    logger.info("Inviata email %s a %s | Oggetto: %s", kind, msg["To"], msg["Subject"])


def process_email_message(
    raw_bytes: bytes,
    cfg: AutomationConfig,
    *,
    send_fn: Callable[[AutomationConfig, EmailMessage], None] = send_email_message,
) -> Tuple[str, RunReport]:
    """
    Elabora una singola email grezza RFC822.

    Restituisce ``(outcome, delta_report)`` dove ``outcome`` è uno tra:
    - ``"skipped"``: email ignorata (non corrisponde ai filtri o nessun PDF PLCB da gestire)
    - ``"processed"``: tutte le fatture elaborate con Verifica OK e inviate
    - ``"error"``: almeno una discrepanza > $0.01 o errore di parsing (inviato alert)
    - ``"transient_failure"``: errore temporaneo di rete/SMTP (da ritentare al prossimo ciclo)
    """
    rep = RunReport(checked_messages=1)
    msg = email.message_from_bytes(raw_bytes)
    subject = decode_mime_header(msg.get("Subject"))
    sender = decode_mime_header(msg.get("From"))

    allowed, reason = should_process_message(msg, cfg)
    if not allowed:
        logger.info("Saltata email (%s) da %s — %s", subject or "(senza oggetto)", sender, reason)
        rep.skipped_messages = 1
        return "skipped", rep

    pdfs = extract_pdf_attachments(msg)
    if not pdfs:
        logger.info("Saltata email (%s) da %s — nessun allegato PDF", subject or "(senza oggetto)", sender)
        rep.skipped_messages = 1
        return "skipped", rep

    had_error_or_mismatch = False
    handled_any = False

    with tempfile.TemporaryDirectory(prefix="plcb_auto_") as tmp_dir:
        tmp_root = Path(tmp_dir)
        for idx, (orig_name, payload) in enumerate(pdfs, start=1):
            src_pdf = tmp_root / f"{idx}_{safe_invoice_filename(Path(orig_name).stem)}.pdf"
            src_pdf.write_bytes(payload)

            try:
                summary, raw_text, sum_p, ok, diff = verify_invoice(str(src_pdf))
            except ValueError as exc:
                is_unrecognized = "non riconosciuto" in str(exc).lower()
                if is_unrecognized and not cfg.alert_on_unrecognized_pdf:
                    logger.info("PDF non PLCB ignorato: %s (%s)", orig_name, exc)
                    continue
                handled_any = True
                had_error_or_mismatch = True
                rep.invoices_error += 1
                logger.warning("Errore riconoscimento PDF %s: %s", orig_name, exc)
                if cfg.alert_recipients:
                    alert_msg = build_error_alert_message(
                        cfg,
                        filename=orig_name,
                        error_msg=str(exc),
                        source_subject=subject,
                        source_sender=sender,
                        original_pdf_path=src_pdf,
                        original_pdf_name=orig_name,
                    )
                    try:
                        _dispatch_email(cfg, send_fn, "ERROR_ALERT", alert_msg)
                        rep.emails_sent += 1
                    except Exception as smtp_exc:
                        logger.error("Invio alert fallito per %s: %s", orig_name, smtp_exc)
                        rep.transient_failures += 1
                        return "transient_failure", rep
                continue
            except Exception as exc:
                handled_any = True
                had_error_or_mismatch = True
                rep.invoices_error += 1
                logger.exception("Errore elaborazione PDF %s", orig_name)
                if cfg.alert_recipients:
                    alert_msg = build_error_alert_message(
                        cfg,
                        filename=orig_name,
                        error_msg=str(exc),
                        source_subject=subject,
                        source_sender=sender,
                        original_pdf_path=src_pdf,
                        original_pdf_name=orig_name,
                    )
                    try:
                        _dispatch_email(cfg, send_fn, "ERROR_ALERT", alert_msg)
                        rep.emails_sent += 1
                    except Exception as smtp_exc:
                        logger.error("Invio alert fallito per %s: %s", orig_name, smtp_exc)
                        rep.transient_failures += 1
                        return "transient_failure", rep
                continue

            handled_any = True
            inv_stem = safe_invoice_filename(summary.invoice_number)
            excel_path = tmp_root / f"{inv_stem}_MarginEdge.xlsx"
            pdf_gca_path = tmp_root / f"{inv_stem}_fattura_GCA.pdf"

            try:
                build_and_save_excel(summary, str(src_pdf), raw_text, str(excel_path))
                write_margin_edge_final_pdf(summary, pdf_gca_path)
            except Exception as exc:
                had_error_or_mismatch = True
                rep.invoices_error += 1
                logger.exception("Errore generazione output per invoice %s", summary.invoice_number)
                if cfg.alert_recipients:
                    alert_msg = build_error_alert_message(
                        cfg,
                        filename=orig_name,
                        error_msg=f"Generazione Excel/PDF fallita: {exc}",
                        source_subject=subject,
                        source_sender=sender,
                        original_pdf_path=src_pdf,
                        original_pdf_name=orig_name,
                    )
                    try:
                        _dispatch_email(cfg, send_fn, "ERROR_ALERT", alert_msg)
                        rep.emails_sent += 1
                    except Exception as smtp_exc:
                        logger.error("Invio alert fallito per %s: %s", orig_name, smtp_exc)
                        rep.transient_failures += 1
                        return "transient_failure", rep
                continue

            if not ok:
                had_error_or_mismatch = True
                rep.invoices_mismatch += 1
                logger.warning(
                    "Fattura %s BLOCCATA: scarto $%.4f (PDF=$%.2f, calcolato=$%.2f)",
                    summary.invoice_number,
                    diff,
                    summary.total_net_price,
                    sum_p,
                )
                if cfg.alert_recipients:
                    alert_msg = build_mismatch_alert_message(
                        cfg,
                        summary=summary,
                        sum_p=sum_p,
                        diff=diff,
                        source_subject=subject,
                        source_sender=sender,
                        original_pdf_path=src_pdf,
                        excel_path=excel_path,
                        pdf_gca_path=pdf_gca_path,
                        original_pdf_name=orig_name,
                    )
                    try:
                        _dispatch_email(cfg, send_fn, "MISMATCH_ALERT", alert_msg)
                        rep.emails_sent += 1
                    except Exception as smtp_exc:
                        logger.error("Invio alert discrepanza fallito per %s: %s", summary.invoice_number, smtp_exc)
                        rep.transient_failures += 1
                        return "transient_failure", rep
                continue

            # Verifica OK: invio ai destinatari predisposti
            outgoing: list[Tuple[str, EmailMessage]] = []
            if cfg.pdf_only_recipients:
                outgoing.append(
                    ("PDF_ONLY", build_pdf_only_message(cfg, summary, sum_p, pdf_gca_path))
                )
            if cfg.full_recipients:
                outgoing.append(
                    (
                        "FULL",
                        build_full_report_message(
                            cfg,
                            summary=summary,
                            sum_p=sum_p,
                            pdf_gca_path=pdf_gca_path,
                            excel_path=excel_path,
                            original_pdf_path=src_pdf,
                            original_pdf_name=f"{inv_stem}_fattura_PLCB.pdf",
                        ),
                    )
                )

            try:
                for kind, out_msg in outgoing:
                    _dispatch_email(cfg, send_fn, kind, out_msg)
                    rep.emails_sent += 1
            except Exception as smtp_exc:
                logger.error("Errore SMTP per fattura %s: %s", summary.invoice_number, smtp_exc)
                rep.transient_failures += 1
                return "transient_failure", rep

            if not cfg.dry_run:
                _archive_outputs(cfg.archive_dir, inv_stem, excel_path, pdf_gca_path, src_pdf)
            rep.invoices_ok += 1

    if not handled_any:
        rep.skipped_messages = 1
        return "skipped", rep

    rep.processed_messages = 1
    return ("error" if had_error_or_mismatch else "processed"), rep


def _merge_reports(total: RunReport, delta: RunReport) -> None:
    total.checked_messages += delta.checked_messages
    total.skipped_messages += delta.skipped_messages
    total.processed_messages += delta.processed_messages
    total.invoices_ok += delta.invoices_ok
    total.invoices_mismatch += delta.invoices_mismatch
    total.invoices_error += delta.invoices_error
    total.emails_sent += delta.emails_sent
    total.transient_failures += delta.transient_failures


def _default_imap_factory(host: str, port: int) -> imaplib.IMAP4_SSL:
    return imaplib.IMAP4_SSL(host, port, timeout=30)


def run_once(
    cfg: AutomationConfig,
    *,
    imap_factory: Callable[[str, int], imaplib.IMAP4_SSL] = _default_imap_factory,
    send_fn: Callable[[AutomationConfig, EmailMessage], None] = send_email_message,
) -> RunReport:
    errors = cfg.validate()
    if errors:
        raise ValueError("Configurazione non valida:\n- " + "\n- ".join(errors))

    report = RunReport()
    imap = imap_factory(cfg.imap_host, cfg.imap_port)
    try:
        try:
            imap.login(cfg.imap_user, cfg.imap_password)
        except imaplib.IMAP4.error as exc:
            pw = cfg.imap_password or ""
            is_gmail_app_pw = len(pw) == 16 and pw.isascii() and pw.isalpha() and pw.islower()
            logger.error(
                "Login IMAP fallito su %s:%d (formato_app_password_16_lettere=%s, lunghezza=%d): %s",
                cfg.imap_host,
                cfg.imap_port,
                is_gmail_app_pw,
                len(pw),
                exc,
            )
            raise
        status, _ = imap.select(_quote_imap_mailbox(cfg.imap_mailbox))
        if status != "OK":
            raise RuntimeError(f"Impossibile aprire la cartella IMAP {cfg.imap_mailbox!r}.")

        status, data = imap.uid("SEARCH", None, "UNSEEN")
        if status != "OK":
            raise RuntimeError("Ricerca IMAP UNSEEN fallita.")

        uids = (data[0] or b"").split()
        logger.info("Trovate %d email non lette in %s.", len(uids), cfg.imap_mailbox)

        for uid in uids:
            fetch_status, msg_data = imap.uid("FETCH", uid, "(BODY.PEEK[])")
            if fetch_status != "OK" or not msg_data:
                report.transient_failures += 1
                continue

            raw_bytes = None
            for item in msg_data:
                if isinstance(item, tuple) and len(item) >= 2 and isinstance(item[1], (bytes, bytearray)):
                    raw_bytes = bytes(item[1])
                    break
            if raw_bytes is None:
                report.skipped_messages += 1
                continue

            outcome, delta = process_email_message(raw_bytes, cfg, send_fn=send_fn)
            _merge_reports(report, delta)

            if cfg.dry_run:
                continue

            if outcome == "processed":
                _mark_and_move_uid(imap, uid, cfg.imap_processed_folder)
            elif outcome == "error":
                _mark_and_move_uid(imap, uid, cfg.imap_error_folder)
            # "skipped" e "transient_failure" restano UNSEEN per sicurezza
    finally:
        try:
            imap.logout()
        except Exception:
            pass

    return report


def run_watch(cfg: AutomationConfig) -> None:
    errors = cfg.validate()
    if errors:
        raise ValueError("Configurazione non valida:\n- " + "\n- ".join(errors))

    stop = False

    def _handle_stop(_signum, _frame) -> None:
        nonlocal stop
        stop = True
        logger.info("Ricevuto segnale di arresto, chiusura in corso…")

    signal.signal(signal.SIGINT, _handle_stop)
    signal.signal(signal.SIGTERM, _handle_stop)

    logger.info(
        "Avvio PLCB email worker in modalità watch (casella=%s, intervallo=%ds, dry_run=%s)",
        cfg.imap_user,
        cfg.poll_interval_seconds,
        cfg.dry_run,
    )
    while not stop:
        try:
            rep = run_once(cfg)
            if rep.processed_messages or rep.transient_failures:
                logger.info(
                    "Ciclo completato: controllate=%d, ok=%d, mismatch=%d, errori=%d, email_inviate=%d",
                    rep.checked_messages,
                    rep.invoices_ok,
                    rep.invoices_mismatch,
                    rep.invoices_error,
                    rep.emails_sent,
                )
        except Exception:
            logger.exception("Errore durante il ciclo di controllo IMAP")

        for _ in range(cfg.poll_interval_seconds):
            if stop:
                break
            time.sleep(1)
