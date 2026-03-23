"""Load, save, and inspect email templates.

Templates live in topic sub-folders of the templates/ directory. Each topic
folder contains one .txt file per language (e.g. ``templates/partnership/en.txt``).

Each .txt file has the format:
  Line 1:    Subject (may contain {placeholders})
  Line 2:    blank
  Line 3+:   Body (may contain {placeholders})
"""

import re
import string
from pathlib import Path

import markdown as _md
from jinja2 import BaseLoader, Environment

from app.config import TEMPLATES_DIR


def list_topics(templates_dir: Path = TEMPLATES_DIR) -> list[str]:
    """Return sorted topic names (sub-folder names) in *templates_dir*."""
    if not templates_dir.is_dir():
        return []
    return sorted(
        p.name
        for p in templates_dir.iterdir()
        if p.is_dir() and not p.name.startswith(".")
    )


def list_languages(
    topic: str | None = None, templates_dir: Path = TEMPLATES_DIR
) -> list[str]:
    """Return sorted language codes available for a *topic*.

    If *topic* is ``None``, return the union of languages across all topics.
    """
    if topic is not None:
        topic_dir = templates_dir / topic
        if not topic_dir.is_dir():
            return []
        return sorted(p.stem for p in topic_dir.glob("*.txt"))

    langs: set[str] = set()
    for topic_name in list_topics(templates_dir):
        langs.update(list_languages(topic_name, templates_dir))
    return sorted(langs)


def load_template(
    topic: str, lang: str, templates_dir: Path = TEMPLATES_DIR
) -> dict[str, str]:
    """Load a template and return ``{'subject': ..., 'body': ...}``.

    Raises ``FileNotFoundError`` if the template file does not exist.
    """
    path = templates_dir / topic / f"{lang}.txt"
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
    topic: str,
    lang: str,
    subject: str,
    body: str,
    templates_dir: Path = TEMPLATES_DIR,
) -> Path:
    """Write a template file and return its path."""
    topic_dir = templates_dir / topic
    topic_dir.mkdir(parents=True, exist_ok=True)
    path = topic_dir / f"{lang}.txt"
    path.write_text(f"{subject}\n\n{body}\n", encoding="utf-8")
    return path


_jinja_env = Environment(loader=BaseLoader(), keep_trailing_newline=True)


def resolve_template(text: str, context: dict[str, str]) -> str:
    """Resolve placeholders in *text* using Jinja2.

    Supports both ``{name}`` shorthand (converted to ``{{ name }}``) and
    native Jinja2 syntax (``{% if formal? %}...{% endif %}``).

    Boolean columns (names ending with ``?``) are converted to ``True``/``False``
    in the context automatically.
    """
    # Expand boolean shorthand: {col?:yes_val:no_val} → {% if col? %}yes_val{% else %}no_val{% endif %}
    text = re.sub(
        r"\{([a-zA-Z_][a-zA-Z0-9_]*\?):([^}:]*):([^}]*)\}",
        r"{% if \1 %}\2{% else %}\3{% endif %}",
        text,
    )
    # Convert {name} shorthand to {{ name }}, but leave {% %} and {{ }} untouched
    converted = re.sub(
        r"(?<!\{)\{([a-zA-Z_][a-zA-Z0-9_?]*)\}(?!\})",
        r"{{ \1 }}",
        text,
    )
    # Jinja2 doesn't allow ? in variable names, so replace col? with col_bool
    # in both the template and context
    bool_context: dict[str, str | bool] = {}
    for key, value in context.items():
        if key.endswith("?"):
            clean_key = key[:-1] + "_bool"
            bool_context[clean_key] = value.strip().lower() in (
                "true",
                "1",
                "ja",
                "yes",
                "x",
            )
        else:
            bool_context[key] = value
    converted = re.sub(
        r"(\{\{|\{%)(.*?)\b([a-zA-Z_][a-zA-Z0-9_]*)\?(.*?)(%\}|\}\})",
        r"\1\2\3_bool\4\5",
        converted,
    )
    tpl = _jinja_env.from_string(converted)
    return tpl.render(**bool_context)


def render_html(
    text: str,
    topic: str | None = None,
    templates_dir: Path = TEMPLATES_DIR,
    use_cid: bool = False,
    font_family: str = "'Verdana', sans-serif",
    font_size: str = "10pt",
) -> str:
    """Convert markdown-formatted *text* to an HTML email body.

    Supports links, bold, italic, paragraphs, line breaks, and images.

    ``image:filename.png`` references are resolved to either ``file://``
    paths (for local preview) or ``cid:`` references (for sending via Outlook).
    """
    html_body = _md.markdown(text, extensions=["nl2br"])

    # Resolve image:filename references
    if topic:
        images_dir = templates_dir / topic / "images"

        def _replace_image(match: re.Match) -> str:
            filename = match.group(1)
            if use_cid:
                return f'src="cid:{filename}"'
            local_path = images_dir / filename
            return f'src="file:///{local_path.as_posix()}"'

        html_body = re.sub(r'src="image:([^"]+)"', _replace_image, html_body)

    return (
        f'<div style="font-family: {font_family}; font-size: {font_size};">'
        f"{html_body}"
        "</div>"
    )


def list_images(topic: str, templates_dir: Path = TEMPLATES_DIR) -> list[str]:
    """Return image filenames referenced by ``image:`` in a topic's images folder."""
    images_dir = templates_dir / topic / "images"
    if not images_dir.is_dir():
        return []
    return sorted(p.name for p in images_dir.iterdir() if p.is_file())


def extract_placeholders(text: str) -> list[str]:
    """Return sorted unique placeholder names found in *text*.

    Detects both ``{name}`` shorthand and Jinja2 ``{{ name }}`` /
    ``{% if name %}`` references.
    """
    names: set[str] = set()

    # {name} shorthand
    formatter = string.Formatter()
    try:
        parsed = formatter.parse(text)
        for _, field_name, _, _ in parsed:
            if field_name is not None:
                root = field_name.split(".")[0].split("[")[0]
                if root:
                    names.add(root)
    except ValueError:
        pass

    # Boolean shorthand {col?:yes:no}
    names.update(re.findall(r"\{([a-zA-Z_][a-zA-Z0-9_]*\?):[^}]*\}", text))

    # Jinja2 {{ name }} and {% if name %} references
    names.update(re.findall(r"\{\{[\s]*([a-zA-Z_][a-zA-Z0-9_?]*)[\s]*\}\}", text))
    names.update(re.findall(r"\{%.*?\b([a-zA-Z_][a-zA-Z0-9_?]*)\b.*?%\}", text))
    # Filter out Jinja2 keywords
    keywords = {"if", "else", "endif", "for", "endfor", "not", "and", "or", "in"}
    names -= keywords

    return sorted(names)
