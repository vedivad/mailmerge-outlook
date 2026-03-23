"""Main application window — slim coordinator for the three tab widgets."""

from PyQt6.QtWidgets import QMainWindow, QTabWidget

from app.config import DEFAULT_CSV
from app.tabs import ContactsTab, SendTab, TemplatesTab


class MainWindow(QMainWindow):
    """MailMerge main window with Contacts, Templates, and Send tabs."""

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("MailMerge")
        self.resize(900, 620)

        # Create tab widgets
        self._contacts_tab = ContactsTab()
        self._templates_tab = TemplatesTab(get_headers=self._contacts_tab.headers)
        self._send_tab = SendTab(
            get_all_contacts=self._contacts_tab.all_contacts,
            get_font_kwargs=self._templates_tab.font_kwargs,
        )

        # Assemble tabs
        tabs = QTabWidget()
        tabs.addTab(self._contacts_tab, "Kontakte")
        tabs.addTab(self._templates_tab, "Vorlagen")
        tabs.addTab(self._send_tab, "Senden")
        self.setCentralWidget(tabs)

        # Cross-tab coordination: when templates change, refresh all tabs
        self._templates_tab.templates_changed.connect(self._refresh_templates)

        # Load initial data
        self._contacts_tab.load_csv(DEFAULT_CSV)
        self._refresh_templates()

    def _refresh_templates(self) -> None:
        """Reload topics/languages across all tabs."""
        topics = self._templates_tab.refresh()
        self._send_tab.refresh_topics(topics)
        self._contacts_tab.rebuild_lang_tabs()
