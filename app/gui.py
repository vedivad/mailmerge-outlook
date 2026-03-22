"""Main application window with Contacts, Templates, and Send tabs."""

from datetime import datetime
from pathlib import Path

from PyQt6.QtCore import Qt
from PyQt6.QtGui import QColor
from PyQt6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QDialog,
    QFileDialog,
    QHBoxLayout,
    QHeaderView,
    QInputDialog,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QSizePolicy,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from app import contact_manager, mailer, template_manager

# Paths relative to project root
_PROJECT_DIR = Path(__file__).resolve().parent.parent
_DEFAULT_CSV = _PROJECT_DIR / "contacts.csv"
_TEMPLATES_DIR = _PROJECT_DIR / "templates"


class MainWindow(QMainWindow):
    """MailMerge main window with Contacts, Templates, and Send tabs."""

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("MailMerge")
        self.resize(900, 620)

        self._contacts_path: Path = _DEFAULT_CSV
        self._rows: list[dict[str, str]] = []
        self._headers: list[str] = []

        tabs = QTabWidget()
        tabs.addTab(self._build_contacts_tab(), "Contacts")
        tabs.addTab(self._build_templates_tab(), "Templates")
        tabs.addTab(self._build_send_tab(), "Send")
        self.setCentralWidget(tabs)

        # Load initial data
        self._load_csv(self._contacts_path)
        self._refresh_languages()

    # ------------------------------------------------------------------
    # Contacts tab
    # ------------------------------------------------------------------

    def _build_contacts_tab(self) -> QWidget:
        """Build the Contacts tab with an editable table and action buttons."""
        widget = QWidget()
        layout = QVBoxLayout(widget)

        # Buttons
        btn_layout = QHBoxLayout()
        btn_load = QPushButton("Load CSV")
        btn_save = QPushButton("Save CSV")
        btn_add = QPushButton("Add Row")
        btn_del = QPushButton("Delete Row")
        btn_add_col = QPushButton("Add Column")
        btn_load.clicked.connect(self._on_load_csv)
        btn_save.clicked.connect(self._on_save_csv)
        btn_add.clicked.connect(self._on_add_row)
        btn_del.clicked.connect(self._on_delete_row)
        btn_add_col.clicked.connect(self._on_add_column)
        for btn in (btn_load, btn_save, btn_add, btn_del, btn_add_col):
            btn_layout.addWidget(btn)
        btn_layout.addStretch()
        layout.addLayout(btn_layout)

        # Table
        self._contacts_table = QTableWidget()
        self._contacts_table.horizontalHeader().setStretchLastSection(True)
        self._contacts_table.horizontalHeader().setSectionResizeMode(
            QHeaderView.ResizeMode.Stretch
        )
        self._contacts_table.cellChanged.connect(self._on_cell_changed)
        layout.addWidget(self._contacts_table)

        return widget

    def _load_csv(self, path: Path) -> None:
        """Load a CSV file into the contacts table."""
        try:
            self._rows = contact_manager.load_csv(path)
        except FileNotFoundError:
            QMessageBox.warning(self, "File not found", f"Could not find {path}")
            return

        self._contacts_path = path
        if self._rows:
            self._headers = list(self._rows[0].keys())
        else:
            self._headers = []

        self._populate_table()

    def _populate_table(self) -> None:
        """Fill the QTableWidget from ``self._rows``."""
        table = self._contacts_table
        table.blockSignals(True)
        table.setRowCount(len(self._rows))
        table.setColumnCount(len(self._headers))
        table.setHorizontalHeaderLabels(self._headers)

        for r, row in enumerate(self._rows):
            for c, header in enumerate(self._headers):
                item = QTableWidgetItem(row.get(header, ""))
                table.setItem(r, c, item)

        table.blockSignals(False)
        self._validate_all_rows()

    def _validate_all_rows(self) -> None:
        """Highlight invalid rows in the contacts table."""
        languages = template_manager.list_languages(_TEMPLATES_DIR)
        table = self._contacts_table
        err_bg = QColor(255, 80, 80, 60)

        for r in range(table.rowCount()):
            row_dict = self._table_row_to_dict(r)
            errors = contact_manager.validate_row(row_dict, languages)
            tooltip = "; ".join(errors) if errors else ""
            for c in range(table.columnCount()):
                item = table.item(r, c)
                if item:
                    if errors:
                        item.setBackground(err_bg)
                    else:
                        item.setData(Qt.ItemDataRole.BackgroundRole, None)
                    item.setToolTip(tooltip)

    def _table_row_to_dict(self, row_index: int) -> dict[str, str]:
        """Convert a table row back to a dict keyed by column headers."""
        table = self._contacts_table
        result: dict[str, str] = {}
        for c, header in enumerate(self._headers):
            item = table.item(row_index, c)
            result[header] = item.text() if item else ""
        return result

    def _table_to_rows(self) -> list[dict[str, str]]:
        """Convert the entire table to a list of dicts."""
        return [
            self._table_row_to_dict(r) for r in range(self._contacts_table.rowCount())
        ]

    # -- Contacts slots --

    def _on_load_csv(self) -> None:
        """Open a file dialog and load a CSV."""
        path, _ = QFileDialog.getOpenFileName(
            self, "Open CSV", str(_PROJECT_DIR), "CSV Files (*.csv)"
        )
        if path:
            self._load_csv(Path(path))

    def _on_save_csv(self) -> None:
        """Save the current table back to CSV."""
        self._rows = self._table_to_rows()
        contact_manager.save_csv(self._contacts_path, self._rows)

    def _on_add_row(self) -> None:
        """Append an empty row to the table."""
        table = self._contacts_table
        table.blockSignals(True)
        r = table.rowCount()
        table.insertRow(r)
        for c, header in enumerate(self._headers):
            table.setItem(r, c, QTableWidgetItem(""))
        table.blockSignals(False)
        self._validate_all_rows()

    def _on_delete_row(self) -> None:
        """Delete the currently selected row."""
        row = self._contacts_table.currentRow()
        if row >= 0:
            self._contacts_table.removeRow(row)
            self._validate_all_rows()

    def _on_add_column(self) -> None:
        """Add a new column (placeholder) to the contacts table."""
        name, ok = QInputDialog.getText(
            self, "Add Column", "Column name (e.g. title, department):"
        )
        if not ok or not name.strip():
            return
        name = name.strip()
        if name in self._headers:
            QMessageBox.information(
                self, "Add Column", f"Column '{name}' already exists."
            )
            return
        self._headers.append(name)
        table = self._contacts_table
        table.blockSignals(True)
        col = table.columnCount()
        table.setColumnCount(col + 1)
        table.setHorizontalHeaderLabels(self._headers)
        for r in range(table.rowCount()):
            table.setItem(r, col, QTableWidgetItem(""))
        table.blockSignals(False)
        self._validate_all_rows()

    def _on_cell_changed(self, row: int, _col: int) -> None:
        """Re-validate the edited row."""
        languages = template_manager.list_languages(_TEMPLATES_DIR)
        row_dict = self._table_row_to_dict(row)
        errors = contact_manager.validate_row(row_dict, languages)
        err_bg = QColor(255, 80, 80, 60)
        tooltip = "; ".join(errors) if errors else ""
        for c in range(self._contacts_table.columnCount()):
            item = self._contacts_table.item(row, c)
            if item:
                if errors:
                    item.setBackground(err_bg)
                else:
                    item.setData(Qt.ItemDataRole.BackgroundRole, None)
                item.setToolTip(tooltip)

    # ------------------------------------------------------------------
    # Templates tab
    # ------------------------------------------------------------------

    def _build_templates_tab(self) -> QWidget:
        """Build the Templates tab with language selector and editor fields."""
        widget = QWidget()
        layout = QVBoxLayout(widget)

        # Language selector row
        sel_layout = QHBoxLayout()
        sel_layout.addWidget(QLabel("Language:"))
        self._lang_combo = QComboBox()
        self._lang_combo.currentTextChanged.connect(self._on_language_changed)
        sel_layout.addWidget(self._lang_combo)
        btn_new_lang = QPushButton("New Language")
        btn_new_lang.clicked.connect(self._on_new_language)
        sel_layout.addWidget(btn_new_lang)
        sel_layout.addStretch()
        layout.addLayout(sel_layout)

        # Subject
        layout.addWidget(QLabel("Subject:"))
        self._subject_edit = QLineEdit()
        layout.addWidget(self._subject_edit)

        # Body
        layout.addWidget(QLabel("Body:"))
        self._body_edit = QTextEdit()
        layout.addWidget(self._body_edit)

        # Save button
        btn_save_tpl = QPushButton("Save Template")
        btn_save_tpl.clicked.connect(self._on_save_template)
        layout.addWidget(btn_save_tpl)

        # Placeholder info
        self._placeholder_label = QLabel()
        self._placeholder_label.setWordWrap(True)
        layout.addWidget(self._placeholder_label)

        # Update placeholders when text changes
        self._subject_edit.textChanged.connect(self._update_placeholder_label)
        self._body_edit.textChanged.connect(self._update_placeholder_label)

        return widget

    def _refresh_languages(self) -> None:
        """Reload the language combo box from disk."""
        self._lang_combo.blockSignals(True)
        self._lang_combo.clear()
        langs = template_manager.list_languages(_TEMPLATES_DIR)
        self._lang_combo.addItems(langs)
        self._lang_combo.blockSignals(False)
        if langs:
            self._lang_combo.setCurrentIndex(0)
            self._on_language_changed(langs[0])

    def _on_language_changed(self, lang: str) -> None:
        """Load the selected language template into the editor fields."""
        if not lang:
            return
        try:
            tpl = template_manager.load_template(lang, _TEMPLATES_DIR)
        except FileNotFoundError:
            return
        self._subject_edit.setText(tpl["subject"])
        self._body_edit.setPlainText(tpl["body"])

    def _on_save_template(self) -> None:
        """Save the current subject/body to the selected language file."""
        lang = self._lang_combo.currentText()
        if not lang:
            return
        template_manager.save_template(
            lang,
            self._subject_edit.text(),
            self._body_edit.toPlainText(),
            _TEMPLATES_DIR,
        )
        self._refresh_languages()
        self._lang_combo.setCurrentText(lang)

    def _on_new_language(self) -> None:
        """Prompt for a new language code and clear the editor."""
        code, ok = QInputDialog.getText(
            self, "New Language", "Language code (e.g. fr, es):"
        )
        if ok and code.strip():
            code = code.strip().lower()
            template_manager.save_template(code, "", "", _TEMPLATES_DIR)
            self._refresh_languages()
            self._lang_combo.setCurrentText(code)

    def _update_placeholder_label(self) -> None:
        """Show placeholders found in the current subject + body."""
        text = self._subject_edit.text() + "\n" + self._body_edit.toPlainText()
        names = template_manager.extract_placeholders(text)
        if names:
            self._placeholder_label.setText(
                "Placeholders: " + ", ".join(f"{{{n}}}" for n in names)
            )
        else:
            self._placeholder_label.setText("No placeholders found.")

    # ------------------------------------------------------------------
    # Send tab
    # ------------------------------------------------------------------

    def _build_send_tab(self) -> QWidget:
        """Build the Send tab with dry-run toggle, send buttons, and log."""
        widget = QWidget()
        layout = QVBoxLayout(widget)

        # Outlook availability notice
        if not mailer.OUTLOOK_AVAILABLE:
            notice = QLabel(
                "⚠ Outlook is not available on this platform. "
                "Only dry-run mode is functional."
            )
            notice.setStyleSheet(
                "color: #b45309; background: #fef3c7; padding: 6px; border-radius: 4px;"
            )
            notice.setWordWrap(True)
            layout.addWidget(notice)

        # Controls row
        ctrl_layout = QHBoxLayout()

        self._dry_run_cb = QCheckBox("Dry run")
        self._dry_run_cb.setChecked(True)
        ctrl_layout.addWidget(self._dry_run_cb)

        btn_preview = QPushButton("Preview")
        btn_preview.clicked.connect(self._on_preview)
        ctrl_layout.addWidget(btn_preview)

        self._btn_send_all = QPushButton("Send All")
        self._btn_send_all.clicked.connect(self._on_send_all)
        ctrl_layout.addWidget(self._btn_send_all)

        self._btn_send_sel = QPushButton("Send Selected")
        self._btn_send_sel.clicked.connect(self._on_send_selected)
        ctrl_layout.addWidget(self._btn_send_sel)

        ctrl_layout.addStretch()
        layout.addLayout(ctrl_layout)

        # Disable real-send buttons when Outlook is unavailable
        if not mailer.OUTLOOK_AVAILABLE:
            for btn in (self._btn_send_all, self._btn_send_sel):
                btn.setEnabled(False)
                btn.setToolTip("Outlook is not available on this platform")

        # Progress bar
        self._progress = QProgressBar()
        self._progress.setValue(0)
        layout.addWidget(self._progress)

        # Log panel
        self._log = QTextEdit()
        self._log.setReadOnly(True)
        self._log.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding
        )
        layout.addWidget(self._log)

        # Summary
        self._summary_label = QLabel()
        layout.addWidget(self._summary_label)

        return widget

    def _log_msg(self, message: str) -> None:
        """Append a timestamped message to the send log."""
        ts = datetime.now().strftime("%H:%M:%S")
        self._log.append(f"[{ts}] {message}")

    def _on_preview(self) -> None:
        """Show a preview of the resolved email for the selected contact."""
        row_idx = self._contacts_table.currentRow()
        if row_idx < 0:
            QMessageBox.information(self, "Preview", "Select a contact row first.")
            return

        row = self._table_row_to_dict(row_idx)
        lang = row.get("language", "")
        try:
            tpl = template_manager.load_template(lang, _TEMPLATES_DIR)
        except FileNotFoundError:
            QMessageBox.warning(
                self,
                "Preview",
                f"No template for language '{lang}'.",
            )
            return

        try:
            subject = tpl["subject"].format(**row)
            body = tpl["body"].format(**row)
        except KeyError as exc:
            QMessageBox.warning(
                self,
                "Preview",
                f"Missing placeholder value: {exc}",
            )
            return

        dlg = QDialog(self)
        dlg.setWindowTitle("Email Preview")
        dlg.resize(500, 400)
        dlg_layout = QVBoxLayout(dlg)
        dlg_layout.addWidget(QLabel(f"<b>To:</b> {row.get('email', '')}"))
        dlg_layout.addWidget(QLabel(f"<b>Subject:</b> {subject}"))
        body_view = QTextEdit()
        body_view.setReadOnly(True)
        body_view.setPlainText(body)
        dlg_layout.addWidget(body_view)
        btn_close = QPushButton("Close")
        btn_close.clicked.connect(dlg.accept)
        dlg_layout.addWidget(btn_close)
        dlg.exec()

    def _on_send_all(self) -> None:
        """Send emails to all contacts."""
        rows = self._table_to_rows()
        self._send_emails(rows)

    def _on_send_selected(self) -> None:
        """Send emails to selected contacts only."""
        selected_rows = sorted(
            {idx.row() for idx in self._contacts_table.selectedIndexes()}
        )
        if not selected_rows:
            QMessageBox.information(
                self, "Send", "Select one or more contact rows first."
            )
            return
        rows = [self._table_row_to_dict(r) for r in selected_rows]
        self._send_emails(rows)

    def _send_emails(self, rows: list[dict[str, str]]) -> None:
        """Execute the send loop for *rows* (dry-run or real)."""
        dry_run = self._dry_run_cb.isChecked()
        languages = template_manager.list_languages(_TEMPLATES_DIR)

        # Validate first
        invalid_rows: list[tuple[int, list[str]]] = []
        for i, row in enumerate(rows):
            errors = contact_manager.validate_row(row, languages)
            if errors:
                invalid_rows.append((i + 1, errors))

        if invalid_rows:
            msg_lines = [f"Row {idx}: {'; '.join(errs)}" for idx, errs in invalid_rows]
            QMessageBox.warning(
                self,
                "Validation errors",
                "Fix these issues before sending:\n\n" + "\n".join(msg_lines),
            )
            return

        # Prepare COM object once for real sends
        outlook_app = None
        if not dry_run:
            try:
                import win32com.client  # type: ignore[import-untyped]

                outlook_app = win32com.client.Dispatch("Outlook.Application")
            except Exception as exc:
                QMessageBox.critical(
                    self,
                    "Outlook error",
                    f"Could not connect to Outlook:\n{exc}",
                )
                return

        sent = 0
        skipped = 0
        errors = 0
        total = len(rows)

        self._progress.setMaximum(total)
        self._progress.setValue(0)
        self._log.clear()
        self._summary_label.clear()

        from PyQt6.QtWidgets import QApplication

        for i, row in enumerate(rows):
            email = row.get("email", "")
            lang = row.get("language", "")

            # Load template
            try:
                tpl = template_manager.load_template(lang, _TEMPLATES_DIR)
            except FileNotFoundError:
                self._log_msg(f"SKIPPED {email} — no template for '{lang}'")
                skipped += 1
                self._progress.setValue(i + 1)
                QApplication.processEvents()
                continue

            # Resolve placeholders
            try:
                subject = tpl["subject"].format(**row)
                body = tpl["body"].format(**row)
            except KeyError as exc:
                self._log_msg(f"SKIPPED {email} — missing placeholder {exc}")
                skipped += 1
                self._progress.setValue(i + 1)
                QApplication.processEvents()
                continue

            # Send or dry-run
            if dry_run:
                result = mailer.dry_run_email(email, subject, body)
                self._log_msg(f"DRY RUN {email}\n{result}")
                sent += 1
            else:
                try:
                    mailer.send_email(email, subject, body, outlook_app=outlook_app)
                    self._log_msg(f"SENT {email}")
                    sent += 1
                except Exception as exc:
                    self._log_msg(f"ERROR {email} — {exc}")
                    errors += 1

            self._progress.setValue(i + 1)
            QApplication.processEvents()

        mode = "dry-run" if dry_run else "send"
        self._summary_label.setText(
            f"Done ({mode}): {sent} sent, {skipped} skipped, {errors} errors"
        )
