# MailMerge — Claude Code Instructions

## What this project is

A PyQt6 desktop application that sends personalized bulk emails through Microsoft
Outlook on Windows. The user maintains a CSV contact list and per-language email
templates with {placeholder} fields. The app resolves placeholders per contact
and sends via Outlook COM automation (win32com).

## Platform situation — read this carefully

- **Target platform: Windows** (Outlook COM via win32com is Windows-only)
- **Development platform: Linux**
- pywin32 may not be installable on Linux — this is expected and fine
- mailer.py guards the import and degrades gracefully when Outlook is unavailable
- The GUI must be fully launchable and testable on Linux; only the send
  functionality is Windows-only
- Never attempt to mock or stub win32com in a way that hides real errors on Windows
- build.bat is for Windows only — do not try to run it on Linux

## Project structure

mailmerge/
  main.py
  app/
    __init__.py
    gui.py             # all PyQt6 code lives here
    mailer.py          # win32com isolation — no GUI code here
    template_manager.py
    contact_manager.py
  templates/           # one .txt file per language, e.g. en.txt, de.txt
  contacts.csv
  requirements.txt
  build.bat
  CLAUDE.md
  .gitignore
  README.md

## Module responsibilities — enforce separation

- gui.py: all PyQt6 widgets, layouts, signals/slots. No business logic.
- mailer.py: only Outlook COM interaction. No GUI imports.
- template_manager.py: only file I/O and parsing for templates. No GUI, no COM.
- contact_manager.py: only CSV loading, saving, validation. No GUI, no COM.
- main.py: entry point only, minimal code, just launches the QApplication.

If a task bleeds across these boundaries, restructure rather than cutting corners.

## Template format

Line 1:    Subject line (may contain {placeholders})
Line 2:    blank
Line 3+:   Body (may contain {placeholders})

Placeholders use Python str.format() syntax. Column names in the CSV must match
placeholder names exactly.

## Code rules

- Use pathlib.Path throughout — no os.path, no hardcoded strings
- No global state — pass dependencies explicitly via function arguments or constructors
- All public functions and classes must have docstrings
- Validate early — catch bad CSV rows and missing templates before attempting to send
- When Outlook is unavailable, raise RuntimeError with a clear human-readable message
- Never silently swallow exceptions — log them to the Send tab log panel or re-raise

## How to run during development (Linux)
```bash
source .venv/bin/activate
python main.py
```

The GUI should launch fully. Send buttons will be disabled with an explanatory
notice since Outlook is unavailable on Linux.

## How to build for Windows

On a Windows machine:
```
build.bat
```
This installs dependencies and produces dist/MailMerge.exe via PyInstaller.

## When making changes

- After any change to template_manager.py or contact_manager.py, verify the
  corresponding GUI tab still reflects the change correctly
- After adding a new placeholder to sample templates, update contacts.csv to
  include a matching column
- Do not modify build.bat unless explicitly asked — it is intentionally minimal
- Keep requirements.txt in sync with any new imports added to the codebase

## What good looks like

- The GUI launches on Linux with no errors or warnings in the terminal
- All three tabs render correctly and are interactive
- Loading contacts.csv populates the Contacts tab table
- Switching languages in the Templates tab loads the correct template
- Dry-run mode produces realistic formatted output in the log panel
- Validation correctly highlights bad rows in red before sending

## Linting and formatting

This project uses [ruff](https://docs.astral.sh/ruff/) for both linting and formatting.
It is the only linter/formatter used — do not introduce black, flake8, isort, or pylint.

Install it into the venv if not already present:
```bash
pip install ruff
```

Run after every non-trivial change:
```bash
ruff check .        # lint
ruff format .       # format
```

Fix all lint errors before considering a task done. If a rule produces a false
positive that genuinely cannot be resolved, suppress it inline with a comment
and briefly explain why:
```python
result = some_call()  # noqa: ERA001 — kept for reference during development
```

Add ruff to requirements.txt alongside the other dependencies.

A minimal ruff config should be added to pyproject.toml:
```toml
[tool.ruff]
target-version = "py311"
line-length = 100

[tool.ruff.lint]
select = ["E", "F", "I", "UP", "B", "SIM"]
```

This enables:
- E/F — standard pycodestyle and pyflakes rules
- I   — import sorting (replaces isort)
- UP  — pyupgrade (modernise syntax automatically)
- B   — flake8-bugbear (common bug patterns)
- SIM — flake8-simplify (unnecessary complexity)