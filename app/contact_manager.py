"""Load, save, and validate the CSV contact list."""

import csv
import re
from pathlib import Path

_EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


def load_csv(path: Path) -> list[dict[str, str]]:
    """Read a CSV file and return a list of row dicts.

    Uses ``utf-8-sig`` encoding to handle an optional BOM from Excel.
    Raises ``FileNotFoundError`` if *path* does not exist.
    """
    with path.open(encoding="utf-8-sig", newline="") as fh:
        reader = csv.DictReader(fh, skipinitialspace=True)
        return list(reader)


def save_csv(path: Path, rows: list[dict[str, str]]) -> None:
    """Write *rows* to a CSV file, preserving the key order of the first row."""
    if not rows:
        path.write_text("", encoding="utf-8")
        return

    headers = list(rows[0].keys())
    with path.open("w", encoding="utf-8", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=headers)
        writer.writeheader()
        writer.writerows(rows)


def validate_row(row: dict[str, str], available_languages: list[str]) -> list[str]:
    """Return a list of human-readable error strings for *row*.

    An empty list means the row is valid.
    """
    errors: list[str] = []

    email = row.get("email", "").strip()
    if not email:
        errors.append("E-Mail fehlt")
    elif not _EMAIL_RE.match(email):
        errors.append(f"E-Mail ungueltig: {email}")

    language = row.get("language", "").strip()
    if not language:
        errors.append("Sprache fehlt")
    elif language not in available_languages:
        errors.append(
            f"Keine Vorlage fuer Sprache '{language}' "
            f"(verfuegbar: {', '.join(available_languages)})"
        )

    return errors
