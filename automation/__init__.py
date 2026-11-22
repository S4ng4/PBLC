"""Automazione email IMAP/SMTP per il processamento automatico delle fatture PLCB."""
from .config import AutomationConfig, load_config
from .worker import RunReport, run_once, run_watch

__all__ = [
    "AutomationConfig",
    "RunReport",
    "load_config",
    "run_once",
    "run_watch",
]
