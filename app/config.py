"""Shared path constants for the MailMerge application."""

from pathlib import Path

PROJECT_DIR: Path = Path(__file__).resolve().parent.parent
DEFAULT_CSV: Path = PROJECT_DIR / "contacts.csv"
TEMPLATES_DIR: Path = PROJECT_DIR / "templates"
