"""MailMerge — entry point."""

import sys
from pathlib import Path

from PyQt6.QtGui import QIcon
from PyQt6.QtWidgets import QApplication

from app.gui import MainWindow

_ICON_PATH = Path(__file__).resolve().parent / "favicon.ico"


def main() -> None:
    """Launch the MailMerge application."""
    application = QApplication(sys.argv)
    application.setApplicationName("MailMerge")
    if _ICON_PATH.exists():
        application.setWindowIcon(QIcon(str(_ICON_PATH)))
    window = MainWindow()
    window.show()
    sys.exit(application.exec())


if __name__ == "__main__":
    main()
