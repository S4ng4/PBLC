from __future__ import annotations

import io
from email.message import EmailMessage
from pathlib import Path
from typing import Dict, List, Tuple

import pytest
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

from automation.config import AutomationConfig, load_config, load_dotenv_file
from automation.mailer import AUTO_HEADER_NAME, AUTO_HEADER_VALUE
from automation.worker import run_once
from email_worker import main as worker_cli_main
from models import InvoiceItem, InvoiceSummary
from pdf_margin_edge_final import write_margin_edge_final_pdf


def _make_pdf_bytes(lines: List[str]) -> bytes:
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=letter)
    y = 750
    c.setFont("Courier", 9)
    for line in lines:
        c.drawString(40, y, line)
        y -= 14
    c.save()
    return buf.getvalue()


def _sample_special_order_pdf(*, net_total: float = 114.40) -> bytes:
    # 2 bottles * base 50.00 = 100.00; delivery fee = 6.00; sales tax 6% = 6.36; local tax 2% = 2.04 -> total = 114.40
    lines = [
        "PENNSYLVANIA LIQUOR CONTROL BOARD",
        "SPECIAL ORDER INVOICE",
        "Invoice Number: SO-98765",
        "Invoice Date: 11/20/2026",
        "Order Number: ORD-111",
        "Customer Name: GRAN CAFFE & L'AQUILA",
        "Customer Number: 12345",
        "Line PLCB Item Code Description Unit Vol UOM Base Price Base Freight LTMF Total Unit Price Unit Qty Line Amount",
        "1 123456 MOET & CHANDON BRUT 750ML EA 50.00 0.00 0.00 50.00 2 100.00",
        "",
        "",
        "Total Base Price: $100.00",
        "Delivery Fee: $6.00",
        "LTMF Fee: $0.00",
        "Total Taxable Amount: $106.00",
        "Sales Tax (6%): $6.36",
        "Local Tax (2%): $2.04",
        "Base Freight: $0.00",
        f"Total Net Price: ${net_total:.2f}",
    ]
    return _make_pdf_bytes(lines)


def _sample_non_plcb_pdf() -> bytes:
    return _make_pdf_bytes(["HELLO WORLD", "This is a random non-PLCB PDF document."])


def _make_incoming_email(
    *,
    sender: str = "orders@pa.gov",
    to: str = "fatture.plcb@gmail.com",
    subject: str = "PLCB Special Order Invoice SO-98765",
    attachments: List[Tuple[str, bytes]] = (),
    extra_headers: Dict[str, str] | None = None,
) -> bytes:
    msg = EmailMessage()
    msg["From"] = sender
    msg["To"] = to
    msg["Subject"] = subject
    if extra_headers:
        for k, v in extra_headers.items():
            msg[k] = v
    msg.set_content("In allegato la fattura.")
    for filename, payload in attachments:
        msg.add_attachment(payload, maintype="application", subtype="pdf", filename=filename)
    return msg.as_bytes()


class FakeIMAP:
    def __init__(self, messages: Dict[bytes, bytes]) -> None:
        self.messages = dict(messages)
        self.logged_in = False
        self.selected_mailbox = ""
        self.created_folders: List[str] = []
        self.copied: List[Tuple[bytes, str]] = []
        self.flags: Dict[bytes, List[str]] = {uid: [] for uid in messages}
        self.expunged = False
        self.logged_out = False

    def login(self, user: str, password: str):
        self.logged_in = True
        return "OK", [b"Logged in"]

    def select(self, mailbox: str):
        self.selected_mailbox = mailbox
        return "OK", [b"1"]

    def create(self, folder: str):
        self.created_folders.append(folder)
        return "OK", [b"Created"]

    def expunge(self):
        self.expunged = True
        return "OK", [b"Expunged"]

    def logout(self):
        self.logged_out = True
        return "BYE", [b"Logged out"]

    def uid(self, command: str, *args):
        cmd = command.upper()
        if cmd == "SEARCH":
            unseen = [uid for uid in self.messages if r"(\Seen)" not in self.flags.get(uid, [])]
            return "OK", [b" ".join(unseen)]
        if cmd == "FETCH":
            uid = args[0]
            raw = self.messages[uid]
            return "OK", [(b"1 (RFC822)", raw)]
        if cmd == "STORE":
            uid = args[0]
            flag = args[2]
            self.flags.setdefault(uid, []).append(flag)
            return "OK", [b"Stored"]
        if cmd == "COPY":
            uid = args[0]
            folder = args[1]
            self.copied.append((uid, folder))
            return "OK", [b"Copied"]
        raise AssertionError(f"Unexpected IMAP UID command: {cmd} {args}")


def _base_config(tmp_path: Path, **overrides) -> AutomationConfig:
    data = dict(
        imap_host="imap.gmail.com",
        imap_port=993,
        imap_user="fatture.plcb@gmail.com",
        imap_password="app-password",
        smtp_host="smtp.gmail.com",
        smtp_port=587,
        smtp_user="fatture.plcb@gmail.com",
        smtp_password="app-password",
        smtp_from="fatture.plcb@gmail.com",
        pdf_only_recipients=("invoices@marginedge.com",),
        full_recipients=("accounting@grancaffelaquila.com", "owner@grancaffelaquila.com"),
        alert_recipients=("admin@grancaffelaquila.com",),
        archive_dir=tmp_path / "archive",
    )
    data.update(overrides)
    return AutomationConfig(**data)


def test_pdf_export_escapes_xml_characters(tmp_path: Path) -> None:
    summary = InvoiceSummary(
        invoice_type="SPECIAL_ORDER",
        invoice_number="SO-AMP-1",
        invoice_date="11/20/2026",
        order_number="ORD-1",
        customer_name="Gran Caffè & <L'Aquila>",
        customer_number="12345",
        total_base_price=100.0,
        delivery_fee=6.0,
        sales_tax_6=6.36,
        local_tax_2=2.04,
        total_net_price=114.40,
        items=[
            InvoiceItem(
                line=1,
                plcb_code="123456",
                description="MOET & CHANDON <BRUT> 750ML",
                unit_volume="750ML",
                uom="EA",
                base_price=50.0,
                base_freight=0.0,
                ltmf=0.0,
                total_unit_price=50.0,
                unit_qty=2,
                line_amount=100.0,
            )
        ],
    )
    out_pdf = tmp_path / "out.pdf"
    res = write_margin_edge_final_pdf(summary, out_pdf)
    assert res.is_file()
    assert res.stat().st_size > 1000


def test_end_to_end_happy_path_dispatches_pdf_only_and_full(tmp_path: Path) -> None:
    raw_email = _make_incoming_email(
        attachments=[("invoice_so.pdf", _sample_special_order_pdf(net_total=114.40))]
    )
    fake_imap = FakeIMAP({b"101": raw_email})
    sent: List[EmailMessage] = []

    cfg = _base_config(tmp_path)
    report = run_once(
        cfg,
        imap_factory=lambda _h, _p: fake_imap,
        send_fn=lambda _cfg, msg: sent.append(msg),
    )

    assert report.checked_messages == 1
    assert report.processed_messages == 1
    assert report.invoices_ok == 1
    assert report.invoices_mismatch == 0
    assert report.invoices_error == 0
    assert report.emails_sent == 2
    assert len(sent) == 2

    # 1) PDF-only email (MarginEdge)
    pdf_only_msg = sent[0]
    assert pdf_only_msg["To"] == "invoices@marginedge.com"
    assert pdf_only_msg[AUTO_HEADER_NAME] == AUTO_HEADER_VALUE
    pdf_only_names = [part.get_filename() for part in pdf_only_msg.iter_attachments()]
    assert pdf_only_names == ["SO-98765_fattura_GCA.pdf"]

    # 2) Full email (internal recipients)
    full_msg = sent[1]
    assert "accounting@grancaffelaquila.com" in full_msg["To"]
    assert "owner@grancaffelaquila.com" in full_msg["To"]
    full_names = [part.get_filename() for part in full_msg.iter_attachments()]
    assert full_names == [
        "SO-98765_fattura_GCA.pdf",
        "SO-98765_MarginEdge.xlsx",
        "SO-98765_fattura_PLCB.pdf",
    ]

    # Moved to Processed folder
    assert fake_imap.copied == [(b"101", '"PLCB/Processed"')]
    assert r"(\Seen)" in fake_imap.flags[b"101"]
    assert r"(\Deleted)" in fake_imap.flags[b"101"]

    # Archived locally
    arch_dir = tmp_path / "archive" / "SO-98765"
    assert (arch_dir / "SO-98765_fattura_GCA.pdf").is_file()
    assert (arch_dir / "SO-98765_MarginEdge.xlsx").is_file()
    assert (arch_dir / "SO-98765_fattura_PLCB.pdf").is_file()


def test_mismatch_blocks_auto_dispatch_and_alerts_admin(tmp_path: Path) -> None:
    # Wrong net total (125.00 instead of 114.40 -> diff > $0.01)
    raw_email = _make_incoming_email(
        attachments=[("bad_total.pdf", _sample_special_order_pdf(net_total=125.00))]
    )
    fake_imap = FakeIMAP({b"202": raw_email})
    sent: List[EmailMessage] = []

    cfg = _base_config(tmp_path)
    report = run_once(
        cfg,
        imap_factory=lambda _h, _p: fake_imap,
        send_fn=lambda _cfg, msg: sent.append(msg),
    )

    assert report.invoices_ok == 0
    assert report.invoices_mismatch == 1
    assert report.emails_sent == 1
    assert len(sent) == 1

    alert = sent[0]
    assert alert["To"] == "admin@grancaffelaquila.com"
    assert "ATTENZIONE" in alert["Subject"]
    att_names = [part.get_filename() for part in alert.iter_attachments()]
    assert att_names == [
        "bad_total.pdf",
        "SO-98765_MarginEdge.xlsx",
        "SO-98765_fattura_GCA.pdf",
    ]
    assert fake_imap.copied == [(b"202", '"PLCB/Error"')]


def test_anti_loop_and_filters_skip_untouched(tmp_path: Path) -> None:
    loop_email = _make_incoming_email(
        attachments=[("inv.pdf", _sample_special_order_pdf())],
        extra_headers={AUTO_HEADER_NAME: AUTO_HEADER_VALUE},
    )
    wrong_to_email = _make_incoming_email(
        to="other@gmail.com",
        attachments=[("inv.pdf", _sample_special_order_pdf())],
    )
    fake_imap = FakeIMAP({b"301": loop_email, b"302": wrong_to_email})
    sent: List[EmailMessage] = []

    cfg = _base_config(tmp_path, imap_to_filter="fatture.plcb+auto@gmail.com")
    report = run_once(
        cfg,
        imap_factory=lambda _h, _p: fake_imap,
        send_fn=lambda _cfg, msg: sent.append(msg),
    )

    assert report.checked_messages == 2
    assert report.skipped_messages == 2
    assert report.emails_sent == 0
    assert sent == []
    assert fake_imap.copied == []
    assert fake_imap.flags[b"301"] == []
    assert fake_imap.flags[b"302"] == []


def test_unrecognized_pdf_respects_alert_setting(tmp_path: Path) -> None:
    raw_email = _make_incoming_email(attachments=[("random.pdf", _sample_non_plcb_pdf())])

    # With alert_on_unrecognized_pdf=True -> sends alert and moves to Error
    fake_imap_alert = FakeIMAP({b"401": raw_email})
    sent_alert: List[EmailMessage] = []
    cfg_alert = _base_config(tmp_path, alert_on_unrecognized_pdf=True)
    rep1 = run_once(
        cfg_alert,
        imap_factory=lambda _h, _p: fake_imap_alert,
        send_fn=lambda _cfg, msg: sent_alert.append(msg),
    )
    assert rep1.invoices_error == 1
    assert len(sent_alert) == 1
    assert fake_imap_alert.copied == [(b"401", '"PLCB/Error"')]

    # With alert_on_unrecognized_pdf=False -> skips silently without moving
    fake_imap_skip = FakeIMAP({b"402": raw_email})
    sent_skip: List[EmailMessage] = []
    cfg_skip = _base_config(tmp_path, alert_on_unrecognized_pdf=False)
    rep2 = run_once(
        cfg_skip,
        imap_factory=lambda _h, _p: fake_imap_skip,
        send_fn=lambda _cfg, msg: sent_skip.append(msg),
    )
    assert rep2.skipped_messages == 1
    assert sent_skip == []
    assert fake_imap_skip.copied == []


def test_dry_run_does_not_send_or_move(tmp_path: Path) -> None:
    raw_email = _make_incoming_email(
        attachments=[("invoice_so.pdf", _sample_special_order_pdf())]
    )
    fake_imap = FakeIMAP({b"501": raw_email})
    sent: List[EmailMessage] = []

    cfg = _base_config(tmp_path, dry_run=True)
    report = run_once(
        cfg,
        imap_factory=lambda _h, _p: fake_imap,
        send_fn=lambda _cfg, msg: sent.append(msg),
    )

    assert report.invoices_ok == 1
    assert report.emails_sent == 2  # Simulated count
    assert sent == []  # Never actually called send_fn
    assert fake_imap.copied == []
    assert fake_imap.flags[b"501"] == []


def test_smtp_transient_failure_keeps_message_unseen(tmp_path: Path) -> None:
    raw_email = _make_incoming_email(
        attachments=[("invoice_so.pdf", _sample_special_order_pdf())]
    )
    fake_imap = FakeIMAP({b"601": raw_email})

    def failing_smtp(_cfg, _msg):
        raise RuntimeError("SMTP connection refused")

    cfg = _base_config(tmp_path)
    report = run_once(
        cfg,
        imap_factory=lambda _h, _p: fake_imap,
        send_fn=failing_smtp,
    )

    assert report.transient_failures == 1
    assert fake_imap.copied == []
    assert fake_imap.flags[b"601"] == []


def test_dotenv_and_cli_check_config(tmp_path: Path, monkeypatch) -> None:
    env_file = tmp_path / "test.env"
    env_file.write_text(
        "\n".join(
            [
                "# comment",
                "IMAP_USER=plcb@example.com",
                'IMAP_PASSWORD="secret-pass"',
                "PDF_ONLY_RECIPIENTS=me@marginedge.com",
                "FULL_RECIPIENTS=acc@example.com, boss@example.com",
            ]
        ),
        encoding="utf-8",
    )
    for k in (
        "IMAP_USER",
        "IMAP_PASSWORD",
        "PDF_ONLY_RECIPIENTS",
        "FULL_RECIPIENTS",
        "ALERT_RECIPIENTS",
    ):
        monkeypatch.delenv(k, raising=False)

    loaded = load_dotenv_file(env_file)
    assert loaded["IMAP_PASSWORD"] == "secret-pass"

    cfg = load_config(env_file=env_file)
    assert cfg.validate() == []
    assert cfg.pdf_only_recipients == ("me@marginedge.com",)
    assert cfg.full_recipients == ("acc@example.com", "boss@example.com")
    assert cfg.alert_recipients == ("plcb@example.com",)

    rc = worker_cli_main(["--check-config", "--env", str(env_file)])
    assert rc == 0


def test_web_ui_automation_config_save_and_mask_password(tmp_path: Path) -> None:
    from web_ui import _read_automation_ui_config, _save_automation_ui_config

    env_file = tmp_path / ".env"
    code, saved = _save_automation_ui_config(
        {
            "imap_user": "fatture.plcb@gmail.com",
            "imap_password": "abcd efgh ijkl mnop",
            "imap_to_filter": "fatture.plcb+auto@gmail.com",
            "pdf_only_recipients": "invoices@marginedge.com",
            "full_recipients": "admin@gca.com, owner@gca.com",
            "alert_recipients": "alerts@gca.com",
            "include_original_pdf_in_full": True,
        },
        env_file,
    )
    assert code == 200
    assert saved["has_password"] is True
    assert "imap_password" not in saved
    assert saved["validation_errors"] == []
    assert saved["pdf_only_recipients"] == "invoices@marginedge.com"
    assert saved["full_recipients"] == "admin@gca.com, owner@gca.com"

    # Updating recipients with empty password preserves existing password
    code2, updated = _save_automation_ui_config(
        {
            "imap_user": "fatture.plcb@gmail.com",
            "imap_password": "",
            "pdf_only_recipients": "new-me@marginedge.com",
            "full_recipients": "contabilita@gca.com",
            "alert_recipients": "alerts@gca.com",
            "include_original_pdf_in_full": False,
        },
        env_file,
    )
    assert code2 == 200
    assert updated["has_password"] is True
    assert updated["pdf_only_recipients"] == "new-me@marginedge.com"
    assert updated["include_original_pdf_in_full"] is False

    reloaded = _read_automation_ui_config(env_file)
    assert reloaded["has_password"] is True
    assert reloaded["full_recipients"] == "contabilita@gca.com"


def test_cli_setup_wizard_writes_env(tmp_path: Path, monkeypatch) -> None:
    import getpass

    env_file = tmp_path / "wizard.env"
    answers = iter(
        [
            "plcb.bot@gmail.com",       # IMAP_USER
            "",                          # IMAP_TO_FILTER
            "invoices@marginedge.com",   # PDF_ONLY_RECIPIENTS
            "contabilita@gca.com",       # FULL_RECIPIENTS
            "admin@gca.com",             # ALERT_RECIPIENTS
            "true",                      # INCLUDE_ORIGINAL_PDF_IN_FULL
        ]
    )
    monkeypatch.setattr("builtins.input", lambda _prompt="": next(answers))
    monkeypatch.setattr(getpass, "getpass", lambda _prompt="": "my-app-pass-1234")

    rc = worker_cli_main(["--setup", "--env", str(env_file)])
    assert rc == 0
    cfg = load_config(env_file=env_file, override_env=True)
    assert cfg.imap_user == "plcb.bot@gmail.com"
    assert cfg.imap_password == "my-app-pass-1234"
    assert cfg.pdf_only_recipients == ("invoices@marginedge.com",)
    assert cfg.full_recipients == ("contabilita@gca.com",)
    assert cfg.alert_recipients == ("admin@gca.com",)

