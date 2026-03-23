"""Main application window — slim coordinator for the three tab widgets."""

from collections.abc import Callable

from PyQt6.QtCore import QSettings
from PyQt6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFormLayout,
    QHBoxLayout,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QSpinBox,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from app.config import DEFAULT_CSV
from app.tabs import ContactsTab, SendTab, TemplatesTab


class MainWindow(QMainWindow):
    """MailMerge main window with Contacts, Templates, and Send tabs."""

    def __init__(
        self,
        current_language: str = "en",
        on_language_change: Callable[[str], None] | None = None,
        on_restart_requested: Callable[[], None] | None = None,
    ) -> None:
        super().__init__()
        self.setWindowTitle("MailMerge")
        self.resize(900, 620)
        self._current_language = current_language
        self._on_language_change = on_language_change
        self._on_restart_requested = on_restart_requested
        self._settings = QSettings("MailMerge", "MailMerge")

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
        self._build_settings_menu()

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

    def _build_settings_menu(self) -> None:
        """Build the settings menu with email backend configuration."""
        settings_menu = self.menuBar().addMenu(self.tr("Settings"))
        email_action = settings_menu.addAction(self.tr("Email delivery..."))
        email_action.triggered.connect(self._open_email_settings)

    def _open_email_settings(self) -> None:
        """Open and process the email delivery settings dialog."""
        dialog = _EmailSettingsDialog(self)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        if not dialog.was_changed():
            return

        answer = QMessageBox.question(
            self,
            self.tr("Restart required"),
            self.tr("Email settings were updated. Restart now?"),
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.Yes,
        )
        if answer == QMessageBox.StandardButton.Yes and self._on_restart_requested:
            self._on_restart_requested()


class _EmailSettingsDialog(QDialog):
    """Dialog for selecting provider and configuring SMTP settings."""

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._settings = QSettings("MailMerge", "MailMerge")
        self._changed = False

        self.setWindowTitle(self.tr("Email delivery settings"))
        self.resize(520, 320)

        root = QVBoxLayout(self)
        form = QFormLayout()
        root.addLayout(form)

        self._provider_combo = QComboBox()
        self._provider_combo.addItem(self.tr("Outlook"), "outlook")
        self._provider_combo.addItem(self.tr("SMTP"), "smtp")
        form.addRow(self.tr("Provider:"), self._provider_combo)

        self._host_edit = QLineEdit()
        form.addRow(self.tr("SMTP host:"), self._host_edit)

        self._port_spin = QSpinBox()
        self._port_spin.setRange(1, 65535)
        form.addRow(self.tr("SMTP port:"), self._port_spin)

        self._from_edit = QLineEdit()
        form.addRow(self.tr("From address:"), self._from_edit)

        self._user_edit = QLineEdit()
        form.addRow(self.tr("SMTP username:"), self._user_edit)

        self._password_edit = QLineEdit()
        self._password_edit.setEchoMode(QLineEdit.EchoMode.Password)
        form.addRow(self.tr("SMTP password:"), self._password_edit)

        tls_row = QHBoxLayout()
        self._starttls_cb = QCheckBox(self.tr("Use STARTTLS"))
        self._ssl_cb = QCheckBox(self.tr("Use SSL/TLS"))
        tls_row.addWidget(self._starttls_cb)
        tls_row.addWidget(self._ssl_cb)
        tls_row.addStretch()
        form.addRow(self.tr("Security:"), tls_row)

        self._ssl_cb.toggled.connect(self._on_ssl_toggled)
        self._provider_combo.currentIndexChanged.connect(self._sync_enabled_fields)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self._save_and_accept)
        buttons.rejected.connect(self.reject)
        root.addWidget(buttons)

        self._load_from_settings()
        self._sync_enabled_fields()

    def was_changed(self) -> bool:
        """Return True if settings were modified and saved."""
        return self._changed

    def _load_from_settings(self) -> None:
        """Populate widgets from QSettings values."""
        provider = self._settings.value("delivery/provider", "outlook", type=str)
        index = self._provider_combo.findData(provider)
        self._provider_combo.setCurrentIndex(0 if index < 0 else index)

        self._host_edit.setText(self._settings.value("delivery/smtp/host", "", type=str))
        self._port_spin.setValue(self._settings.value("delivery/smtp/port", 587, type=int))
        self._from_edit.setText(self._settings.value("delivery/smtp/from", "", type=str))
        self._user_edit.setText(self._settings.value("delivery/smtp/user", "", type=str))
        self._password_edit.setText(
            self._settings.value("delivery/smtp/password", "", type=str)
        )
        self._starttls_cb.setChecked(
            self._settings.value("delivery/smtp/use_starttls", True, type=bool)
        )
        self._ssl_cb.setChecked(
            self._settings.value("delivery/smtp/use_ssl", False, type=bool)
        )

    def _sync_enabled_fields(self) -> None:
        """Enable SMTP fields only when SMTP is selected."""
        is_smtp = self._provider_combo.currentData() == "smtp"
        for widget in (
            self._host_edit,
            self._port_spin,
            self._from_edit,
            self._user_edit,
            self._password_edit,
            self._starttls_cb,
            self._ssl_cb,
        ):
            widget.setEnabled(is_smtp)

    def _on_ssl_toggled(self, checked: bool) -> None:
        """Adjust related fields for SSL selection."""
        if checked:
            self._starttls_cb.setChecked(False)
            if self._port_spin.value() == 587:
                self._port_spin.setValue(465)
        elif self._port_spin.value() == 465:
            self._port_spin.setValue(587)

    def _save_and_accept(self) -> None:
        """Persist changes and close the dialog."""
        provider = self._provider_combo.currentData()
        if provider == "smtp" and (
            not self._host_edit.text().strip() or not self._from_edit.text().strip()
        ):
            QMessageBox.warning(
                self,
                self.tr("Email delivery settings"),
                self.tr("SMTP host and From address are required for SMTP provider."),
            )
            return

        values = {
            "delivery/provider": provider,
            "delivery/smtp/host": self._host_edit.text().strip(),
            "delivery/smtp/port": self._port_spin.value(),
            "delivery/smtp/from": self._from_edit.text().strip(),
            "delivery/smtp/user": self._user_edit.text().strip(),
            "delivery/smtp/password": self._password_edit.text(),
            "delivery/smtp/use_starttls": self._starttls_cb.isChecked(),
            "delivery/smtp/use_ssl": self._ssl_cb.isChecked(),
        }

        changed = False
        for key, value in values.items():
            previous = self._settings.value(key)
            if str(previous) != str(value):
                changed = True
                self._settings.setValue(key, value)

        self._changed = changed
        self.accept()
