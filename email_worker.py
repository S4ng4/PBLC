#!/usr/bin/env python3
"""
Worker email automatico per PLCB Invoice Processor.

Esempi:
  python3 email_worker.py --check-config
  python3 email_worker.py --once --dry-run
  python3 email_worker.py --once
  python3 email_worker.py --watch --interval 60
"""
from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path
from typing import Optional, Sequence

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from automation.config import load_config, load_dotenv_file, save_dotenv_file
from automation.worker import run_once, run_watch


def _prompt(label: str, default: str = "") -> str:
    suffix = f" [{default}]" if default else ""
    val = input(f"{label}{suffix}: ").strip()
    return val if val else default


def _run_setup_wizard(env_path: Path) -> int:
    import getpass

    existing = load_dotenv_file(env_path, override=False) if env_path.is_file() else {}
    print("\n=== Configurazione Automazione Email PLCB ===")
    print(f"File di configurazione: {env_path}")
    print("Premi Invio per mantenere il valore attuale tra [parentesi].\n")

    imap_user = _prompt("1. Email predefinita che riceve le fatture PLCB (IMAP_USER)", existing.get("IMAP_USER", ""))
    has_pw = bool(existing.get("IMAP_PASSWORD"))
    pw_hint = " (lascia vuoto per mantenere la password attuale)" if has_pw else ""
    imap_password = getpass.getpass(f"2. App Password Gmail / IMAP (IMAP_PASSWORD){pw_hint}: ").strip()
    if not imap_password:
        imap_password = existing.get("IMAP_PASSWORD", "")

    imap_to_filter = _prompt(
        "3. Filtro alias destinatario opzionale (es. tuamail+plcb@gmail.com, vuoto = tutta la casella)",
        existing.get("IMAP_TO_FILTER", ""),
    )
    pdf_only = _prompt(
        "4. Email destinatari SOLO PDF MarginEdgeFinal (es. email MarginEdge, separate da virgola)",
        existing.get("PDF_ONLY_RECIPIENTS", ""),
    )
    full_recips = _prompt(
        "5. Email destinatari PDF + Excel (es. contabilità / copia personale, separate da virgola)",
        existing.get("FULL_RECIPIENTS", ""),
    )
    alert_recips = _prompt(
        "6. Email destinatari ALERT in caso di scarto > $0.01 o errore",
        existing.get("ALERT_RECIPIENTS", "") or imap_user,
    )
    inc_orig = _prompt(
        "7. Allegare anche il PDF originale PLCB nelle email PDF + Excel? (true/false)",
        existing.get("INCLUDE_ORIGINAL_PDF_IN_FULL", "true"),
    )

    updates = {
        "IMAP_HOST": existing.get("IMAP_HOST", "imap.gmail.com"),
        "IMAP_PORT": existing.get("IMAP_PORT", "993"),
        "IMAP_USER": imap_user,
        "IMAP_PASSWORD": imap_password,
        "IMAP_TO_FILTER": imap_to_filter,
        "SMTP_HOST": existing.get("SMTP_HOST", "smtp.gmail.com"),
        "SMTP_PORT": existing.get("SMTP_PORT", "587"),
        "PDF_ONLY_RECIPIENTS": pdf_only,
        "FULL_RECIPIENTS": full_recips,
        "ALERT_RECIPIENTS": alert_recips,
        "INCLUDE_ORIGINAL_PDF_IN_FULL": inc_orig,
    }
    save_dotenv_file(updates, env_path)
    cfg = load_config(env_file=env_path, override_env=True)
    errors = cfg.validate()
    print(f"\nConfigurazione salvata in: {env_path}")
    if errors:
        print("Attenzione — mancano ancora alcuni dati:", file=sys.stderr)
        for err in errors:
            print(f"  - {err}", file=sys.stderr)
        return 2
    print("Pronto! Fai una prova sicura con:  python3 email_worker.py --once --dry-run")
    return 0


def _build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="email_worker.py",
        description=(
            "Controlla la casella email predefinita (IMAP), elabora automaticamente le fatture PLCB "
            "allegate e invia il PDF MarginEdgeFinal / Excel ai destinatari predisposti (SMTP)."
        ),
    )
    mode = p.add_mutually_exclusive_group()
    mode.add_argument(
        "--setup",
        action="store_true",
        help="Procedura guidata interattiva per inserire le email desiderate e salvare .env.",
    )
    mode.add_argument(
        "--once",
        action="store_true",
        help="Esegue un singolo controllo della casella ed esce (consigliato per cron e GitHub Actions).",
    )
    mode.add_argument(
        "--watch",
        action="store_true",
        help="Rimane in ascolto continuo controllando la casella a intervalli regolari.",
    )
    mode.add_argument(
        "--check-config",
        action="store_true",
        help="Verifica la configurazione (.env / variabili d'ambiente) senza collegarsi a IMAP/SMTP.",
    )
    p.add_argument(
        "--env",
        dest="env_file",
        default=None,
        help="Percorso di un file .env personalizzato (default: .env nella cartella del progetto).",
    )
    p.add_argument(
        "--interval",
        type=int,
        default=None,
        help="Intervallo di polling in secondi per --watch (sovrascrive POLL_INTERVAL_SECONDS).",
    )
    p.add_argument(
        "--dry-run",
        action="store_true",
        help="Simula l'elaborazione senza inviare email né spostare/marcare messaggi in IMAP.",
    )
    p.add_argument(
        "-v",
        "--verbose",
        action="store_true",
        help="Mostra log più dettagliati.",
    )
    return p


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = _build_parser().parse_args(argv)
    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    )

    if args.setup:
        target_env = Path(args.env_file) if args.env_file else ROOT / ".env"
        return _run_setup_wizard(target_env)

    cfg = load_config(
        env_file=args.env_file,
        dry_run_override=True if args.dry_run else None,
        poll_interval_override=args.interval,
    )

    errors = cfg.validate()
    if args.check_config:
        if errors:
            print("Configurazione NON valida:", file=sys.stderr)
            for err in errors:
                print(f"  - {err}", file=sys.stderr)
            return 2
        print("Configurazione OK:")
        print(f"  IMAP: {cfg.imap_user} @ {cfg.imap_host}:{cfg.imap_port} (cartella: {cfg.imap_mailbox})")
        print(f"  SMTP: {cfg.smtp_from} @ {cfg.smtp_host}:{cfg.smtp_port}")
        print(f"  Destinatari solo PDF (es. MarginEdge): {', '.join(cfg.pdf_only_recipients) or '(nessuno)'}")
        print(f"  Destinatari completi (PDF + Excel):    {', '.join(cfg.full_recipients) or '(nessuno)'}")
        print(f"  Destinatari alert errori/scarti:       {', '.join(cfg.alert_recipients) or '(nessuno)'}")
        print(f"  Cartella processate: {cfg.imap_processed_folder} | errori: {cfg.imap_error_folder}")
        print(f"  Dry-run: {cfg.dry_run}")
        return 0

    if errors:
        print("Configurazione NON valida:", file=sys.stderr)
        for err in errors:
            print(f"  - {err}", file=sys.stderr)
        return 2

    if args.watch:
        run_watch(cfg)
        return 0

    # Default: --once
    rep = run_once(cfg)
    print(
        f"Completato: controllate={rep.checked_messages}, "
        f"saltate={rep.skipped_messages}, "
        f"fatture_ok={rep.invoices_ok}, "
        f"discrepanze={rep.invoices_mismatch}, "
        f"errori={rep.invoices_error}, "
        f"email_inviate={rep.emails_sent}, "
        f"errori_temporanei={rep.transient_failures}"
    )
    return 1 if rep.transient_failures > 0 else 0


if __name__ == "__main__":
    raise SystemExit(main())
