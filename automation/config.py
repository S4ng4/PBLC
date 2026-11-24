from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path
from typing import Mapping, Optional, Tuple


DEFAULT_PDF_ONLY_SUBJECT = "PLCB Invoice {invoice_number} ({invoice_type} - {invoice_date})"
DEFAULT_FULL_SUBJECT = "[GCA] Fattura PLCB {invoice_number} — MarginEdge (${total_net:.2f})"
DEFAULT_ALERT_SUBJECT = "[ATTENZIONE PLCB] Verifica fallita / Errore — {invoice_number}"


def load_dotenv_file(env_path: str | Path, *, override: bool = False) -> dict[str, str]:
    """
    Legge un file .env semplice (KEY=VALUE) senza dipendenze esterne.
    Con override=False (default) non sovrascrive variabili già presenti in os.environ.
    """
    path = Path(env_path)
    loaded: dict[str, str] = {}
    if not path.is_file():
        return loaded

    for raw_line in path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#"):
            continue
        if line.startswith("export "):
            line = line[len("export ") :].strip()
        if "=" not in line:
            continue
        key, val = line.split("=", 1)
        key = key.strip()
        val = val.strip()
        if not key:
            continue
        if len(val) >= 2 and val[0] == val[-1] and val[0] in ('"', "'"):
            val = val[1:-1]
        loaded[key] = val
        if override:
            os.environ[key] = val
        else:
            os.environ.setdefault(key, val)
    return loaded


def _format_env_value(val: str) -> str:
    v = (val or "").replace("\r", "").replace("\n", "").strip()
    if not v:
        return ""
    if any(ch in v for ch in (" ", "#", '"', "'")):
        escaped = v.replace("\\", "\\\\").replace('"', '\\"')
        return f'"{escaped}"'
    return v


def save_dotenv_file(updates: Mapping[str, str], env_path: str | Path) -> Path:
    """
    Crea o aggiorna un file .env preservando commenti e chiavi esistenti non modificate.
    Imposta i permessi a 0600 (solo proprietario) su sistemi POSIX.
    """
    path = Path(env_path)
    existing_lines: list[str] = []
    if path.is_file():
        existing_lines = path.read_text(encoding="utf-8").splitlines()

    remaining = {k: str(v) for k, v in updates.items()}
    new_lines: list[str] = []

    for line in existing_lines:
        stripped = line.strip()
        if not stripped or stripped.startswith("#") or "=" not in stripped:
            new_lines.append(line)
            continue
        raw_key = stripped.split("=", 1)[0].strip()
        if raw_key.startswith("export "):
            raw_key = raw_key[len("export ") :].strip()
        if raw_key in remaining:
            new_lines.append(f"{raw_key}={_format_env_value(remaining.pop(raw_key))}")
        else:
            new_lines.append(line)

    if remaining:
        if new_lines and new_lines[-1].strip():
            new_lines.append("")
        for key, val in remaining.items():
            new_lines.append(f"{key}={_format_env_value(val)}")

    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text("\n".join(new_lines) + "\n", encoding="utf-8")
    try:
        os.chmod(path, 0o600)
    except OSError:
        pass
    return path


def _parse_csv(raw: str) -> Tuple[str, ...]:
    items = [part.strip() for part in (raw or "").replace(";", ",").split(",")]
    return tuple(item for item in items if item)


def _parse_bool(raw: str, default: bool = False) -> bool:
    s = (raw or "").strip().lower()
    if not s:
        return default
    return s in ("1", "true", "yes", "y", "si", "sì", "on")


def _parse_int(raw: str, default: int) -> int:
    s = (raw or "").strip()
    if not s:
        return default
    try:
        return int(s)
    except ValueError as exc:
        raise ValueError(f"Valore intero non valido: {raw!r}") from exc


@dataclass(frozen=True)
class AutomationConfig:
    # IMAP (casella di ricezione predefinita)
    imap_host: str
    imap_port: int
    imap_user: str
    imap_password: str
    imap_mailbox: str = "INBOX"
    imap_processed_folder: str = "PLCB/Processed"
    imap_error_folder: str = "PLCB/Error"
    imap_allowed_senders: Tuple[str, ...] = ()
    imap_to_filter: str = ""
    imap_subject_contains: str = ""

    # SMTP (invio email predisposte)
    smtp_host: str = "smtp.gmail.com"
    smtp_port: int = 587
    smtp_user: str = ""
    smtp_password: str = ""
    smtp_from: str = ""

    # Destinatari predisposti (gruppi separati)
    pdf_only_recipients: Tuple[str, ...] = ()
    full_recipients: Tuple[str, ...] = ()
    alert_recipients: Tuple[str, ...] = ()

    # Opzioni contenuti e notifiche
    include_original_pdf_in_full: bool = True
    alert_on_unrecognized_pdf: bool = True
    pdf_only_subject_template: str = DEFAULT_PDF_ONLY_SUBJECT
    full_subject_template: str = DEFAULT_FULL_SUBJECT
    alert_subject_template: str = DEFAULT_ALERT_SUBJECT

    # Esecuzione
    poll_interval_seconds: int = 60
    archive_dir: Optional[Path] = None
    dry_run: bool = False

    def validate(self) -> list[str]:
        errors: list[str] = []
        if not self.imap_user:
            errors.append("IMAP_USER mancante (indirizzo email della casella predefinita).")
        if not self.imap_password:
            errors.append("IMAP_PASSWORD mancante (App Password della casella email).")
        if not self.imap_host:
            errors.append("IMAP_HOST mancante.")
        if not self.smtp_host:
            errors.append("SMTP_HOST mancante.")
        if not self.pdf_only_recipients and not self.full_recipients:
            errors.append(
                "Configura almeno uno tra PDF_ONLY_RECIPIENTS (es. MarginEdge) "
                "e FULL_RECIPIENTS (email interne PDF + Excel)."
            )
        if self.poll_interval_seconds < 5:
            errors.append("POLL_INTERVAL_SECONDS deve essere almeno 5 secondi.")
        return errors


def _normalize_password(host: str, password: str) -> str:
    pw = (password or "").strip()
    if "gmail" in (host or "").lower():
        compact = "".join(pw.split())
        if len(compact) == 16 and compact.isascii() and compact.isalpha():
            return compact.lower()
    return pw


def load_config(
    env: Optional[Mapping[str, str]] = None,
    *,
    env_file: Optional[str | Path] = None,
    override_env: bool = False,
    dry_run_override: Optional[bool] = None,
    poll_interval_override: Optional[int] = None,
) -> AutomationConfig:
    if env is None:
        if env_file is not None:
            load_dotenv_file(env_file, override=override_env)
        else:
            default_env = Path(__file__).resolve().parent.parent / ".env"
            if default_env.is_file():
                load_dotenv_file(default_env, override=override_env)
        source: Mapping[str, str] = os.environ
    else:
        source = env

    def get(key: str, default: str = "") -> str:
        return (source.get(key) or default).strip()

    imap_host = get("IMAP_HOST", "imap.gmail.com")
    smtp_host = get("SMTP_HOST", "smtp.gmail.com")
    imap_user = get("IMAP_USER")
    imap_password = _normalize_password(imap_host, get("IMAP_PASSWORD"))
    smtp_user = get("SMTP_USER", imap_user)
    smtp_password = _normalize_password(smtp_host, get("SMTP_PASSWORD", imap_password))
    smtp_from = get("SMTP_FROM", smtp_user or imap_user)

    alert_recipients = _parse_csv(get("ALERT_RECIPIENTS"))
    if not alert_recipients and imap_user:
        alert_recipients = (imap_user,)

    archive_raw = get("ARCHIVE_DIR")
    archive_dir = Path(archive_raw).expanduser() if archive_raw else None

    dry_run = _parse_bool(get("DRY_RUN"), default=False)
    if dry_run_override is not None:
        dry_run = dry_run_override

    poll_interval = _parse_int(get("POLL_INTERVAL_SECONDS"), default=60)
    if poll_interval_override is not None:
        poll_interval = poll_interval_override

    return AutomationConfig(
        imap_host=get("IMAP_HOST", "imap.gmail.com"),
        imap_port=_parse_int(get("IMAP_PORT"), default=993),
        imap_user=imap_user,
        imap_password=imap_password,
        imap_mailbox=get("IMAP_MAILBOX", "INBOX"),
        imap_processed_folder=get("IMAP_PROCESSED_FOLDER", "PLCB/Processed"),
        imap_error_folder=get("IMAP_ERROR_FOLDER", "PLCB/Error"),
        imap_allowed_senders=_parse_csv(get("IMAP_ALLOWED_SENDERS")),
        imap_to_filter=get("IMAP_TO_FILTER"),
        imap_subject_contains=get("IMAP_SUBJECT_CONTAINS"),
        smtp_host=get("SMTP_HOST", "smtp.gmail.com"),
        smtp_port=_parse_int(get("SMTP_PORT"), default=587),
        smtp_user=smtp_user,
        smtp_password=smtp_password,
        smtp_from=smtp_from,
        pdf_only_recipients=_parse_csv(get("PDF_ONLY_RECIPIENTS")),
        full_recipients=_parse_csv(get("FULL_RECIPIENTS")),
        alert_recipients=alert_recipients,
        include_original_pdf_in_full=_parse_bool(
            get("INCLUDE_ORIGINAL_PDF_IN_FULL"), default=True
        ),
        alert_on_unrecognized_pdf=_parse_bool(
            get("ALERT_ON_UNRECOGNIZED_PDF"), default=True
        ),
        pdf_only_subject_template=get("PDF_ONLY_SUBJECT_TEMPLATE", DEFAULT_PDF_ONLY_SUBJECT),
        full_subject_template=get("FULL_SUBJECT_TEMPLATE", DEFAULT_FULL_SUBJECT),
        alert_subject_template=get("ALERT_SUBJECT_TEMPLATE", DEFAULT_ALERT_SUBJECT),
        poll_interval_seconds=poll_interval,
        archive_dir=archive_dir,
        dry_run=dry_run,
    )
