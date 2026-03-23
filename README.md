# MailMerge

A PyQt6 desktop application that sends personalized bulk emails through Microsoft Outlook on Windows.

## Features

- **Contact management** — CSV-based with language sub-tabs, search/filter, validation, and inline editing (copy/paste, multi-row delete)
- **Template system** — per-topic, per-language email templates with Markdown formatting, image embedding, and font control
- **Placeholders** — `{name}`, `{company}`, etc. are resolved per contact from CSV columns
- **Boolean columns** — columns ending with `?` (e.g. `formal?`) render as checkboxes and support conditional template syntax: `{formal?:Sehr geehrte:Liebe}`
- **Jinja2 support** — full Jinja2 syntax (`{% if %}`, `{{ }}`) available for advanced templates alongside the simple shorthand
- **Column management** — add, rename, delete, and reorder columns via the header context menu
- **Topic & language management** — add, rename, and delete templates from a management dialog
- **Send modes** — dry-run (preview), draft (save to Outlook drafts), and send, with per-email progress and logging
- **Outlook preview** — open a single email in Outlook for inspection before bulk sending

## Template Format

Templates are stored in `templates/<topic>/<language>.txt`:

```
Subject line with {placeholders}

Body with {name} and **Markdown** formatting.

Conditional: {formal?:Sehr geehrte:Liebe} {name},

Images: ![description](image:photo.png)
```

## Development (Linux)

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python main.py
```

The GUI launches fully on Linux. Send functionality is disabled since Outlook is unavailable.

## Build (Windows)

```
build.bat
```

Produces `dist/MailMerge.exe` via PyInstaller.

## License

This project is licensed under the GNU General Public License v3.0 — see [LICENSE](LICENSE) for details.
