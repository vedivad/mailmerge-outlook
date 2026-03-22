"""MailMerge — entry point."""

import sys

from PyQt6.QtWidgets import QApplication

from app.gui import MainWindow


def main() -> None:
    """Launch the MailMerge application."""
    application = QApplication(sys.argv)
    application.setApplicationName("MailMerge")
    window = MainWindow()
    window.show()
    sys.exit(application.exec())


if __name__ == "__main__":
    main()
