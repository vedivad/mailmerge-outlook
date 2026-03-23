"""MailMerge — entry point."""

import sys
from pathlib import Path

from PyQt6.QtCore import QLocale, QProcess, QSettings, QTranslator
from PyQt6.QtGui import QIcon
from PyQt6.QtWidgets import QApplication

from app.gui import MainWindow

_ICON_PATH = Path(__file__).resolve().parent / "favicon.ico"
_TRANSLATIONS_DIR = Path(__file__).resolve().parent / "translations"


def _resolve_language(settings: QSettings) -> str:
    """Resolve the language code from settings or system locale."""
    preferred = settings.value("ui/language", "auto", type=str)
    if preferred and preferred != "auto":
        return preferred

    return "de" if QLocale.system().name().lower().startswith("de") else "en"


def _install_translator(application: QApplication, language_code: str) -> QTranslator | None:
    """Load and install a QM translator for *language_code* if available."""
    if language_code == "en":
        return None

    qm_path = _TRANSLATIONS_DIR / f"{language_code}.qm"
    translator = QTranslator()
    if not qm_path.exists():
        print(f"[i18n] Translation file not found: {qm_path}")
        return None
    if not translator.load(str(qm_path)):
        print(f"[i18n] Failed to load translation file: {qm_path}")
        return None

    application.installTranslator(translator)
    return translator


def main() -> None:
    """Launch the MailMerge application."""
    application = QApplication(sys.argv)
    application.setApplicationName("MailMerge")
    settings = QSettings("MailMerge", "MailMerge")

    active_language = _resolve_language(settings)
    translator = _install_translator(application, active_language)
    if _ICON_PATH.exists():
        application.setWindowIcon(QIcon(str(_ICON_PATH)))

    def on_language_change(language_code: str) -> None:
        """Persist a language change and restart the app process."""
        settings.setValue("ui/language", language_code)
        QProcess.startDetached(sys.executable, sys.argv)
        application.quit()

    def on_restart_requested() -> None:
        """Restart the app process after settings changes."""
        QProcess.startDetached(sys.executable, sys.argv)
        application.quit()

    window = MainWindow(
        current_language=active_language,
        on_language_change=on_language_change,
        on_restart_requested=on_restart_requested,
    )
    # Keep a strong reference for translator lifetime.
    application._mailmerge_translator = translator  # type: ignore[attr-defined]
    window.show()
    sys.exit(application.exec())


if __name__ == "__main__":
    main()
