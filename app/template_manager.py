"""Load, save, and inspect email templates.

Each template is a .txt file in the templates/ directory. The first line is
the email subject, followed by a blank line, followed by the body. Both
subject and body may contain {placeholder} fields that match CSV column names.
"""

import string
from pathlib import Path

TEMPLATES_DIR: Path = Path(__file__).resolve().parent.parent / "templates"


def list_languages(templates_dir: Path = TEMPLATES_DIR) -> list[str]:
    """Return sorted language codes available in *templates_dir*."""
    return sorted(p.stem for p in templates_dir.glob("*.txt"))


def load_template(lang: str, templates_dir: Path = TEMPLATES_DIR) -> dict[str, str]:
    """Load a template and return ``{'subject': ..., 'body': ...}``.

    Raises ``FileNotFoundError`` if the template file does not exist.
    """
    path = templates_dir / f"{lang}.txt"
    text = path.read_text(encoding="utf-8")
    lines = text.split("\n")

    subject = lines[0] if lines else ""
    # Find the first blank line separating subject from body
    body_start = 1
    for i, line in enumerate(lines[1:], start=1):
        if line.strip() == "":
            body_start = i + 1
            break

    body = "\n".join(lines[body_start:]).strip()
    return {"subject": subject, "body": body}


def save_template(
    lang: str,
    subject: str,
    body: str,
    templates_dir: Path = TEMPLATES_DIR,
) -> Path:
    """Write a template file and return its path."""
    templates_dir.mkdir(parents=True, exist_ok=True)
    path = templates_dir / f"{lang}.txt"
    path.write_text(f"{subject}\n\n{body}\n", encoding="utf-8")
    return path


def extract_placeholders(text: str) -> list[str]:
    """Return sorted unique ``{placeholder}`` names found in *text*."""
    formatter = string.Formatter()
    names: set[str] = set()
    for _, field_name, _, _ in formatter.parse(text):
        if field_name is not None:
            # Only keep the top-level name (before any '.' or '[')
            root = field_name.split(".")[0].split("[")[0]
            if root:
                names.add(root)
    return sorted(names)
