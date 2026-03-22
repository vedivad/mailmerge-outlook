# MailMerge

A PyQt6 desktop application that sends personalized bulk emails through Microsoft Outlook on Windows.

## Features

- CSV-based contact management with validation
- Per-language email templates with `{placeholder}` support
- Dry-run mode for previewing emails without sending
- Progress tracking and per-email status logging

## Development (Linux)

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt   # pywin32 will fail — that's expected
python main.py
```

The GUI launches fully on Linux. Send functionality is disabled since Outlook is unavailable.

## Build (Windows)

```
build.bat
```

Produces `dist/MailMerge.exe` via PyInstaller.
