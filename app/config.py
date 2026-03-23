"""Shared path constants for the MailMerge application."""

from pathlib import Path

PROJECT_DIR: Path = Path(__file__).resolve().parent.parent
DATA_DIR: Path = PROJECT_DIR / "data"
DEFAULT_CSV: Path = DATA_DIR / "contacts.csv"
TEMPLATES_DIR: Path = DATA_DIR / "templates"
