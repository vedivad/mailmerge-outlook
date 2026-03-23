# MailMerge

A PyQt6 desktop app for personalized bulk email from CSV contacts and per-language templates.

## Features

- CSV contact management with validation and per-language tabs
- Topic/language template editor with placeholders and Markdown/Jinja2 support
- Inline image support in templates (`image:filename`)
- Send modes: Dry run, Send, and Draft (provider-dependent)
- Delivery backends: Outlook (Windows) and SMTP (cross-platform)
- UI localization: English and German

## Template Format

Templates live in `templates/<topic>/<language>.txt`:

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

The GUI runs fully on Linux. Outlook sending is unavailable on Linux, but SMTP works when configured.

## Delivery Setup

Open `Settings` -> `Email delivery...` and choose:

- `Outlook`: Windows + Outlook + pywin32
- `SMTP`: host, port, sender, credentials, TLS options

After saving delivery settings, restart the app when prompted.

Provider capabilities:

- Outlook: Dry run, Send, Draft, Preview
- SMTP: Dry run, Send

## Language

Use `Language` -> `English` / `German` in the menu bar. Changing language requires restart.

Compile translations:

```bash
lrelease translations/de.ts -qm translations/de.qm
```

If `translations/de.qm` is missing, German falls back to source strings.

## Build (Windows)

```bash
build.bat
```

Produces `dist/MailMerge.exe` via PyInstaller.

## License

This project is licensed under the GNU General Public License v3.0 — see [LICENSE](LICENSE) for details.
