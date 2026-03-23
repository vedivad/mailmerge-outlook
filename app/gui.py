"""Main application window — slim coordinator for the three tab widgets."""

from collections.abc import Callable

from PyQt6.QtWidgets import QMainWindow, QMessageBox, QTabWidget

from app.config import DEFAULT_CSV
from app.tabs import ContactsTab, SendTab, TemplatesTab


class MainWindow(QMainWindow):
    """MailMerge main window with Contacts, Templates, and Send tabs."""

    def __init__(
        self,
        current_language: str = "en",
        on_language_change: Callable[[str], None] | None = None,
    ) -> None:
        super().__init__()
        self.setWindowTitle("MailMerge")
        self.resize(900, 620)
        self._current_language = current_language
        self._on_language_change = on_language_change

        # Create tab widgets
        self._contacts_tab = ContactsTab()
        self._templates_tab = TemplatesTab(get_headers=self._contacts_tab.headers)
        self._send_tab = SendTab(
            get_all_contacts=self._contacts_tab.all_contacts,
            get_font_kwargs=self._templates_tab.font_kwargs,
        )

        # Assemble tabs
        tabs = QTabWidget()
        tabs.addTab(self._contacts_tab, self.tr("Contacts"))
        tabs.addTab(self._templates_tab, self.tr("Templates"))
        tabs.addTab(self._send_tab, self.tr("Send"))
        self.setCentralWidget(tabs)
        self._build_language_menu()

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

    def _build_language_menu(self) -> None:
        """Build the language menu in the menu bar."""
        language_menu = self.menuBar().addMenu(self.tr("Language"))

        self._lang_actions: dict[str, object] = {}
        for code, label in (("en", self.tr("English")), ("de", self.tr("German"))):
            action = language_menu.addAction(label)
            action.setCheckable(True)
            action.setChecked(code == self._current_language)
            action.triggered.connect(
                lambda checked, lang_code=code: self._handle_language_action(
                    checked, lang_code
                )
            )
            self._lang_actions[code] = action

    def _handle_language_action(self, checked: bool, language_code: str) -> None:
        """Handle a language action and request app restart when needed."""
        if not checked:
            return
        if language_code == self._current_language:
            return

        answer = QMessageBox.question(
            self,
            self.tr("Restart required"),
            self.tr("Language will change after restart. Restart now?"),
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.Yes,
        )
        if answer == QMessageBox.StandardButton.Yes:
            if self._on_language_change is not None:
                self._on_language_change(language_code)
            return

        # Revert check state if user cancels restart.
        for code, action in self._lang_actions.items():
            action.setChecked(code == self._current_language)
