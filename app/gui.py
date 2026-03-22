"""Main application window with Contacts, Templates, and Send tabs."""

from datetime import datetime
from pathlib import Path

from PyQt6.QtCore import QTimer, Qt
from PyQt6.QtGui import QColor
from PyQt6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QDialog,
    QDialogButtonBox,
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


class _ExcelTable(QTableWidget):
    """QTableWidget that moves down on Enter, like Excel."""

    def keyPressEvent(self, event) -> None:  # noqa: N802
        """Commit the current edit and move down on Enter/Return."""
        if event.key() in (Qt.Key.Key_Return, Qt.Key.Key_Enter):
            row = self.currentRow()
            col = self.currentColumn()
            super().keyPressEvent(event)
            next_row = row + 1
            if next_row < self.rowCount():
                QTimer.singleShot(0, lambda: self.setCurrentCell(next_row, col))
        else:
            super().keyPressEvent(event)


class _ContactPickerDialog(QDialog):
    """Dialog for picking one or more contacts with search filtering."""

    def __init__(
        self,
        contacts: list[dict[str, str]],
        multi_select: bool = True,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle("Select Contacts" if multi_select else "Select Contact")
        self.resize(600, 450)
        self._contacts = contacts
        self._multi_select = multi_select
        self._checkboxes: list[QCheckBox] = []
        self._selected: list[dict[str, str]] = []

        layout = QVBoxLayout(self)

        # Search bar
        self._search = QLineEdit()
        self._search.setPlaceholderText("Filter contacts...")
        self._search.textChanged.connect(self._apply_filter)
        layout.addWidget(self._search)

        # Select all (multi-select only)
        if multi_select:
            self._select_all_cb = QCheckBox("Select all")
            self._select_all_cb.stateChanged.connect(self._on_select_all)
            layout.addWidget(self._select_all_cb)

        # Contact list as table with checkboxes / radio-style selection
        headers = list(contacts[0].keys()) if contacts else []
        self._table = QTableWidget(
            len(contacts), len(headers) + (1 if multi_select else 0)
        )

        if multi_select:
            self._table.setHorizontalHeaderLabels([""] + headers)
        else:
            self._table.setHorizontalHeaderLabels(headers)
            self._table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
            self._table.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)

        for r, contact in enumerate(contacts):
            col_offset = 0
            if multi_select:
                cb = QCheckBox()
                container = QWidget()
                lay = QHBoxLayout(container)
                lay.addWidget(cb)
                lay.setAlignment(Qt.AlignmentFlag.AlignCenter)
                lay.setContentsMargins(0, 0, 0, 0)
                self._table.setCellWidget(r, 0, container)
                self._checkboxes.append(cb)
                col_offset = 1

            for c, header in enumerate(headers):
                item = QTableWidgetItem(contact.get(header, ""))
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                self._table.setItem(r, c + col_offset, item)

        if multi_select:
            self._table.horizontalHeader().setSectionResizeMode(
                0, QHeaderView.ResizeMode.ResizeToContents
            )
            for c in range(1, self._table.columnCount()):
                self._table.horizontalHeader().setSectionResizeMode(
                    c, QHeaderView.ResizeMode.Stretch
                )
        else:
            self._table.horizontalHeader().setSectionResizeMode(
                QHeaderView.ResizeMode.Stretch
            )

        layout.addWidget(self._table)

        # Buttons
        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self._on_accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _apply_filter(self, text: str) -> None:
        """Show/hide rows based on the search text."""
        text = text.lower()
        for r, contact in enumerate(self._contacts):
            values = " ".join(contact.values()).lower()
            visible = text in values
            self._table.setRowHidden(r, not visible)

    def _on_select_all(self, state: int) -> None:
        """Toggle all visible checkboxes."""
        checked = state == Qt.CheckState.Checked.value
        for r, cb in enumerate(self._checkboxes):
            if not self._table.isRowHidden(r):
                cb.setChecked(checked)

    def _on_accept(self) -> None:
        """Collect selected contacts and accept the dialog."""
        if self._multi_select:
            self._selected = [
                self._contacts[r]
                for r, cb in enumerate(self._checkboxes)
                if cb.isChecked()
            ]
        else:
            rows = self._table.selectionModel().selectedRows()
            if rows:
                self._selected = [self._contacts[rows[0].row()]]
        self.accept()

    def selected_contacts(self) -> list[dict[str, str]]:
        """Return the contacts chosen by the user."""
        return self._selected


class MainWindow(QMainWindow):
    """MailMerge main window with Contacts, Templates, and Send tabs."""

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("MailMerge")
        self.resize(900, 620)

        self._contacts_path: Path = _DEFAULT_CSV
        self._rows: list[dict[str, str]] = []
        # Headers excluding "language" — language is determined by sub-tab
        self._headers: list[str] = []
        self._loading_template: bool = False

        tabs = QTabWidget()
        tabs.addTab(self._build_contacts_tab(), "Contacts")
        tabs.addTab(self._build_templates_tab(), "Templates")
        tabs.addTab(self._build_send_tab(), "Send")
        self.setCentralWidget(tabs)

        # Load initial data
        self._load_csv(self._contacts_path)
        self._refresh_templates()

    # ------------------------------------------------------------------
    # Contacts tab
    # ------------------------------------------------------------------

    def _build_contacts_tab(self) -> QWidget:
        """Build the Contacts tab with language sub-tabs and action buttons."""
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

        # Language sub-tabs — one table per language
        self._lang_tabs = QTabWidget()
        self._lang_tables: dict[str, _ExcelTable] = {}
        layout.addWidget(self._lang_tabs)

        return widget

    def _current_lang(self) -> str:
        """Return the language code of the currently selected contacts sub-tab."""
        return self._lang_tabs.tabText(self._lang_tabs.currentIndex())

    def _current_table(self) -> _ExcelTable | None:
        """Return the table for the currently selected language tab."""
        lang = self._current_lang()
        return self._lang_tables.get(lang)

    def _make_table(self, lang: str) -> _ExcelTable:
        """Create a new table widget for a language tab."""
        table = _ExcelTable()
        table.cellChanged.connect(
            lambda row, col: self._on_cell_changed(lang, row, col)
        )
        self._lang_tables[lang] = table
        return table

    def _load_csv(self, path: Path) -> None:
        """Load a CSV file into the contacts tables, split by language."""
        try:
            self._rows = contact_manager.load_csv(path)
        except FileNotFoundError:
            QMessageBox.warning(self, "File not found", f"Could not find {path}")
            return

        self._contacts_path = path
        if self._rows:
            self._headers = [h for h in self._rows[0].keys() if h != "language"]
        else:
            self._headers = []

        self._rebuild_lang_tabs()

    def _rebuild_lang_tabs(self) -> None:
        """Rebuild the language sub-tabs from current data."""
        self._lang_tabs.blockSignals(True)
        self._lang_tabs.clear()
        self._lang_tables.clear()

        # Union of languages across all topics
        languages = template_manager.list_languages(templates_dir=_TEMPLATES_DIR)

        # Group rows by language
        rows_by_lang: dict[str, list[dict[str, str]]] = {lang: [] for lang in languages}
        for row in self._rows:
            lang = row.get("language", "")
            if lang not in rows_by_lang:
                rows_by_lang[lang] = []
            rows_by_lang[lang].append(row)

        # Create a tab for each language
        for lang in languages:
            table = self._make_table(lang)
            self._populate_table(table, rows_by_lang.get(lang, []))
            self._lang_tabs.addTab(table, lang)

        # If there are rows with unknown languages, put them in an "other" tab
        other_rows = []
        for lang, rows in rows_by_lang.items():
            if lang not in languages:
                other_rows.extend(rows)
        if other_rows:
            table = self._make_table("other")
            self._populate_table(table, other_rows)
            self._lang_tabs.addTab(table, "other")

        self._lang_tabs.blockSignals(False)

    def _populate_table(self, table: _ExcelTable, rows: list[dict[str, str]]) -> None:
        """Fill a table from a list of row dicts (without the language column)."""
        table.blockSignals(True)
        table.setRowCount(len(rows))
        table.setColumnCount(len(self._headers))
        table.setHorizontalHeaderLabels(self._headers)

        for r, row in enumerate(rows):
            for c, header in enumerate(self._headers):
                item = QTableWidgetItem(row.get(header, ""))
                table.setItem(r, c, item)

        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        table.blockSignals(False)
        self._validate_table(table)

    def _validate_table(self, table: _ExcelTable) -> None:
        """Highlight invalid rows in a contacts table."""
        err_bg = QColor(255, 80, 80, 60)

        for r in range(table.rowCount()):
            row_dict = self._table_row_to_dict(table, r)
            errors = self._validate_contact(row_dict)
            tooltip = "; ".join(errors) if errors else ""
            for c in range(table.columnCount()):
                item = table.item(r, c)
                if item:
                    if errors:
                        item.setBackground(err_bg)
                    else:
                        item.setData(Qt.ItemDataRole.BackgroundRole, None)
                    item.setToolTip(tooltip)

    @staticmethod
    def _validate_contact(row: dict[str, str]) -> list[str]:
        """Validate a contact row (language is always valid from the tab)."""
        import re

        errors: list[str] = []
        email = row.get("email", "").strip()
        if not email:
            errors.append("Email is missing")
        elif not re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", email):
            errors.append(f"Email is malformed: {email}")
        return errors

    def _table_row_to_dict(self, table: _ExcelTable, row_index: int) -> dict[str, str]:
        """Convert a table row to a dict keyed by column headers."""
        result: dict[str, str] = {}
        for c, header in enumerate(self._headers):
            item = table.item(row_index, c)
            result[header] = item.text() if item else ""
        return result

    def _all_contacts(self) -> list[dict[str, str]]:
        """Gather all contacts from all language tabs, with language added."""
        rows: list[dict[str, str]] = []
        for i in range(self._lang_tabs.count()):
            lang = self._lang_tabs.tabText(i)
            table = self._lang_tables[lang]
            for r in range(table.rowCount()):
                row = self._table_row_to_dict(table, r)
                row["language"] = lang
                rows.append(row)
        return rows

    # -- Contacts slots --

    def _on_load_csv(self) -> None:
        """Open a file dialog and load a CSV."""
        path, _ = QFileDialog.getOpenFileName(
            self, "Open CSV", str(_PROJECT_DIR), "CSV Files (*.csv)"
        )
        if path:
            self._load_csv(Path(path))

    def _on_save_csv(self) -> None:
        """Save all contacts (from all language tabs) back to CSV."""
        rows = self._all_contacts()
        contact_manager.save_csv(self._contacts_path, rows)

    def _on_add_row(self) -> None:
        """Append an empty row to the current language tab's table."""
        table = self._current_table()
        if table is None:
            return
        table.blockSignals(True)
        r = table.rowCount()
        table.insertRow(r)
        for c in range(len(self._headers)):
            table.setItem(r, c, QTableWidgetItem(""))
        table.blockSignals(False)
        self._validate_table(table)

    def _on_delete_row(self) -> None:
        """Delete the currently selected row in the active language tab."""
        table = self._current_table()
        if table is None:
            return
        row = table.currentRow()
        if row >= 0:
            table.removeRow(row)
            self._validate_table(table)

    def _on_add_column(self) -> None:
        """Add a new column (placeholder) to all language tables."""
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
        for table in self._lang_tables.values():
            table.blockSignals(True)
            col = table.columnCount()
            table.setColumnCount(col + 1)
            table.setHorizontalHeaderLabels(self._headers)
            for r in range(table.rowCount()):
                table.setItem(r, col, QTableWidgetItem(""))
            table.blockSignals(False)

    def _on_cell_changed(self, lang: str, row: int, _col: int) -> None:
        """Re-validate the edited row."""
        table = self._lang_tables.get(lang)
        if table is None:
            return
        row_dict = self._table_row_to_dict(table, row)
        errors = self._validate_contact(row_dict)
        err_bg = QColor(255, 80, 80, 60)
        tooltip = "; ".join(errors) if errors else ""
        for c in range(table.columnCount()):
            item = table.item(row, c)
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
        """Build the Templates tab with topic/language selectors and editor fields."""
        widget = QWidget()
        layout = QVBoxLayout(widget)

        # Topic + Language selector row
        sel_layout = QHBoxLayout()

        sel_layout.addWidget(QLabel("Topic:"))
        self._topic_combo = QComboBox()
        self._topic_combo.currentTextChanged.connect(self._on_topic_changed)
        sel_layout.addWidget(self._topic_combo)

        btn_new_topic = QPushButton("New Topic")
        btn_new_topic.clicked.connect(self._on_new_topic)
        sel_layout.addWidget(btn_new_topic)

        sel_layout.addWidget(QLabel("Language:"))
        self._lang_combo = QComboBox()
        self._lang_combo.currentTextChanged.connect(self._on_template_selection_changed)
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

        # Formatting toolbar
        fmt_layout = QHBoxLayout()
        btn_bold = QPushButton("B")
        btn_bold.setFixedWidth(32)
        btn_bold.setStyleSheet("font-weight: bold;")
        btn_bold.setToolTip("Bold")
        btn_bold.clicked.connect(lambda: self._insert_format("**", "**", "bold text"))

        btn_italic = QPushButton("I")
        btn_italic.setFixedWidth(32)
        btn_italic.setStyleSheet("font-style: italic;")
        btn_italic.setToolTip("Italic")
        btn_italic.clicked.connect(lambda: self._insert_format("*", "*", "italic text"))

        btn_link = QPushButton("Link")
        btn_link.setToolTip("Insert link")
        btn_link.clicked.connect(self._insert_link)

        btn_image = QPushButton("Image")
        btn_image.setToolTip("Insert image")
        btn_image.clicked.connect(self._insert_image)

        for btn in (btn_bold, btn_italic, btn_link, btn_image):
            fmt_layout.addWidget(btn)
        fmt_layout.addStretch()
        layout.addLayout(fmt_layout)

        self._body_edit = QTextEdit()
        layout.addWidget(self._body_edit)

        # Bottom row: status + preview button + placeholder info
        bottom_layout = QHBoxLayout()

        self._save_status_label = QLabel()
        bottom_layout.addWidget(self._save_status_label)

        bottom_layout.addStretch()

        self._placeholder_label = QLabel()
        self._placeholder_label.setWordWrap(True)
        bottom_layout.addWidget(self._placeholder_label)

        bottom_layout.addStretch()

        btn_preview_tpl = QPushButton("Preview")
        btn_preview_tpl.clicked.connect(self._on_preview_template)
        bottom_layout.addWidget(btn_preview_tpl)

        layout.addLayout(bottom_layout)

        # Auto-save timer (debounce 500ms)
        self._save_timer = QTimer()
        self._save_timer.setSingleShot(True)
        self._save_timer.setInterval(500)
        self._save_timer.timeout.connect(self._auto_save_template)

        # Connect edits to debounce and placeholder update
        self._subject_edit.textChanged.connect(self._on_template_edited)
        self._body_edit.textChanged.connect(self._on_template_edited)

        return widget

    def _refresh_templates(self) -> None:
        """Reload topic/language combos and contacts sub-tabs from disk."""
        topics = template_manager.list_topics(_TEMPLATES_DIR)

        # Topic combo
        prev_topic = self._topic_combo.currentText()
        self._topic_combo.blockSignals(True)
        self._topic_combo.clear()
        self._topic_combo.addItems(topics)
        self._topic_combo.blockSignals(False)

        if prev_topic and prev_topic in topics:
            self._topic_combo.setCurrentText(prev_topic)
        elif topics:
            self._topic_combo.setCurrentIndex(0)
        self._on_topic_changed(self._topic_combo.currentText())

        # Send tab topic combo
        prev_send_topic = self._send_topic_combo.currentText()
        self._send_topic_combo.blockSignals(True)
        self._send_topic_combo.clear()
        self._send_topic_combo.addItems(topics)
        self._send_topic_combo.blockSignals(False)
        if prev_send_topic and prev_send_topic in topics:
            self._send_topic_combo.setCurrentText(prev_send_topic)
        elif topics:
            self._send_topic_combo.setCurrentIndex(0)

        # Refresh contacts sub-tabs (union of all languages)
        self._rebuild_lang_tabs()

    def _on_topic_changed(self, topic: str) -> None:
        """Update the language combo when the topic changes."""
        if not topic:
            return
        langs = template_manager.list_languages(topic, _TEMPLATES_DIR)
        prev_lang = self._lang_combo.currentText()
        self._lang_combo.blockSignals(True)
        self._lang_combo.clear()
        self._lang_combo.addItems(langs)
        self._lang_combo.blockSignals(False)
        if prev_lang and prev_lang in langs:
            self._lang_combo.setCurrentText(prev_lang)
        elif langs:
            self._lang_combo.setCurrentIndex(0)
        self._on_template_selection_changed(self._lang_combo.currentText())

    def _on_template_selection_changed(self, lang: str) -> None:
        """Load the selected topic/language template into the editor."""
        topic = self._topic_combo.currentText()
        if not topic or not lang:
            return
        # Suppress auto-save while loading
        self._loading_template = True
        try:
            tpl = template_manager.load_template(topic, lang, _TEMPLATES_DIR)
        except FileNotFoundError:
            self._subject_edit.clear()
            self._body_edit.clear()
            self._save_status_label.clear()
            self._loading_template = False
            return
        self._subject_edit.setText(tpl["subject"])
        self._body_edit.setPlainText(tpl["body"])
        self._save_status_label.setText("Saved")
        self._save_status_label.setStyleSheet("color: gray;")
        self._loading_template = False

    def _on_template_edited(self) -> None:
        """Mark as unsaved and restart the debounce timer."""
        self._update_placeholder_label()
        if getattr(self, "_loading_template", False):
            return
        self._save_status_label.setText("Unsaved changes...")
        self._save_status_label.setStyleSheet("color: orange;")
        self._save_timer.start()

    def _auto_save_template(self) -> None:
        """Save the current template to disk (called by debounce timer)."""
        topic = self._topic_combo.currentText()
        lang = self._lang_combo.currentText()
        if not topic or not lang:
            return
        template_manager.save_template(
            topic,
            lang,
            self._subject_edit.text(),
            self._body_edit.toPlainText(),
            _TEMPLATES_DIR,
        )
        self._save_status_label.setText("Saved")
        self._save_status_label.setStyleSheet("color: gray;")

    def _on_preview_template(self) -> None:
        """Show the current template body rendered as HTML."""
        body = self._body_edit.toPlainText()
        if not body.strip():
            return
        topic = self._topic_combo.currentText()
        html = template_manager.render_html(
            body, topic=topic, templates_dir=_TEMPLATES_DIR
        )
        dlg = QDialog(self)
        dlg.setWindowTitle("Template Preview")
        dlg.resize(500, 400)
        dlg_layout = QVBoxLayout(dlg)
        subject = self._subject_edit.text()
        if subject:
            dlg_layout.addWidget(QLabel(f"<b>Subject:</b> {subject}"))
        view = QTextEdit()
        view.setReadOnly(True)
        view.setHtml(html)
        dlg_layout.addWidget(view)
        btn_close = QPushButton("Close")
        btn_close.clicked.connect(dlg.accept)
        dlg_layout.addWidget(btn_close)
        dlg.exec()

    def _insert_format(self, prefix: str, suffix: str, placeholder: str) -> None:
        """Wrap the selected text (or insert placeholder) with markdown formatting."""
        cursor = self._body_edit.textCursor()
        selected = cursor.selectedText()
        if selected:
            cursor.insertText(f"{prefix}{selected}{suffix}")
        else:
            cursor.insertText(f"{prefix}{placeholder}{suffix}")
            # Select the placeholder so the user can type over it
            cursor.movePosition(
                cursor.MoveOperation.Left, cursor.MoveMode.MoveAnchor, len(suffix)
            )
            cursor.movePosition(
                cursor.MoveOperation.Left, cursor.MoveMode.KeepAnchor, len(placeholder)
            )
            self._body_edit.setTextCursor(cursor)
        self._body_edit.setFocus()

    def _insert_link(self) -> None:
        """Prompt for link text and URL, then insert a markdown link."""
        cursor = self._body_edit.textCursor()
        selected = cursor.selectedText()

        dlg = QDialog(self)
        dlg.setWindowTitle("Insert Link")
        dlg_layout = QVBoxLayout(dlg)

        dlg_layout.addWidget(QLabel("Display text:"))
        text_edit = QLineEdit(selected if selected else "")
        text_edit.setPlaceholderText("e.g. Click here")
        dlg_layout.addWidget(text_edit)

        dlg_layout.addWidget(QLabel("URL:"))
        url_edit = QLineEdit()
        url_edit.setPlaceholderText("e.g. https://example.com")
        dlg_layout.addWidget(url_edit)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(dlg.accept)
        buttons.rejected.connect(dlg.reject)
        dlg_layout.addWidget(buttons)

        if dlg.exec() != QDialog.DialogCode.Accepted:
            return

        link_text = text_edit.text().strip() or "link"
        url = url_edit.text().strip()
        if not url:
            return

        cursor.insertText(f"[{link_text}]({url})")
        self._body_edit.setFocus()

    def _insert_image(self) -> None:
        """Pick an image file, copy it to the topic's images folder, and insert a reference."""
        topic = self._topic_combo.currentText()
        if not topic:
            QMessageBox.information(self, "Image", "Select a topic first.")
            return

        path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Image",
            "",
            "Images (*.png *.jpg *.jpeg *.gif *.bmp)",
        )
        if not path:
            return

        src = Path(path)
        images_dir = _TEMPLATES_DIR / topic / "images"
        images_dir.mkdir(parents=True, exist_ok=True)

        import shutil

        dest = images_dir / src.name
        if dest.exists():
            reply = QMessageBox.question(
                self,
                "Image exists",
                f"'{src.name}' already exists. Overwrite?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            if reply != QMessageBox.StandardButton.Yes:
                return
        shutil.copy2(src, dest)

        desc, ok = QInputDialog.getText(
            self, "Image Description", "Description (for accessibility):"
        )
        desc = desc.strip() if ok and desc.strip() else src.stem

        cursor = self._body_edit.textCursor()
        cursor.insertText(f"![{desc}](image:{src.name})")
        self._body_edit.setFocus()

    def _on_new_topic(self) -> None:
        """Prompt for a new topic name."""
        name, ok = QInputDialog.getText(
            self, "New Topic", "Topic name (e.g. partnership, follow-up):"
        )
        if ok and name.strip():
            name = name.strip().lower().replace(" ", "-")
            langs = template_manager.list_languages(templates_dir=_TEMPLATES_DIR)
            first_lang = langs[0] if langs else "en"
            template_manager.save_template(name, first_lang, "", "", _TEMPLATES_DIR)
            self._refresh_templates()
            self._topic_combo.setCurrentText(name)

    def _on_new_language(self) -> None:
        """Prompt for a new language code within the current topic."""
        topic = self._topic_combo.currentText()
        if not topic:
            return
        code, ok = QInputDialog.getText(
            self, "New Language", "Language code (e.g. fr, es):"
        )
        if ok and code.strip():
            code = code.strip().lower()
            template_manager.save_template(topic, code, "", "", _TEMPLATES_DIR)
            self._refresh_templates()
            self._topic_combo.setCurrentText(topic)
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
        """Build the Send tab with topic selector, dry-run toggle, and log."""
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

        # Topic + Signature selector row
        topic_layout = QHBoxLayout()
        topic_layout.addWidget(QLabel("Topic:"))
        self._send_topic_combo = QComboBox()
        topic_layout.addWidget(self._send_topic_combo)

        topic_layout.addWidget(QLabel("Signature:"))
        self._signature_combo = QComboBox()
        self._signature_combo.addItem("Default", None)
        for sig in mailer.list_signatures():
            self._signature_combo.addItem(sig, sig)
        topic_layout.addWidget(self._signature_combo)

        topic_layout.addStretch()
        layout.addLayout(topic_layout)

        # Controls row
        ctrl_layout = QHBoxLayout()

        self._dry_run_cb = QCheckBox("Dry run")
        self._dry_run_cb.setChecked(True)
        self._dry_run_cb.toggled.connect(self._on_dry_run_toggled)
        ctrl_layout.addWidget(self._dry_run_cb)

        btn_preview = QPushButton("Preview")
        btn_preview.clicked.connect(self._on_preview)
        ctrl_layout.addWidget(btn_preview)

        self._btn_send_sel = QPushButton("Send")
        self._btn_send_sel.clicked.connect(self._on_send_selected)
        ctrl_layout.addWidget(self._btn_send_sel)

        ctrl_layout.addStretch()
        layout.addLayout(ctrl_layout)

        # Update send button state based on Outlook availability and dry-run
        self._on_dry_run_toggled(self._dry_run_cb.isChecked())

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

    def _send_topic(self) -> str:
        """Return the topic selected in the Send tab."""
        return self._send_topic_combo.currentText()

    def _on_dry_run_toggled(self, checked: bool) -> None:
        """Enable or disable send buttons based on dry-run state and Outlook availability."""
        enabled = checked or mailer.OUTLOOK_AVAILABLE
        for btn in (self._btn_send_sel,):
            btn.setEnabled(enabled)
            if not enabled:
                btn.setToolTip("Outlook is not available — enable dry run to test")
            else:
                btn.setToolTip("")

    def _log_msg(self, message: str) -> None:
        """Append a timestamped message to the send log."""
        ts = datetime.now().strftime("%H:%M:%S")
        self._log.append(f"[{ts}] {message}")

    def _on_preview(self) -> None:
        """Open a contact picker, then show a preview for the chosen contact."""
        topic = self._send_topic()
        if not topic:
            QMessageBox.information(self, "Preview", "Select a topic first.")
            return

        contacts = self._all_contacts()
        if not contacts:
            QMessageBox.information(self, "Preview", "No contacts loaded.")
            return

        dlg = _ContactPickerDialog(contacts, multi_select=False, parent=self)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        selected = dlg.selected_contacts()
        if not selected:
            return

        row = selected[0]
        lang = row.get("language", "")

        try:
            tpl = template_manager.load_template(topic, lang, _TEMPLATES_DIR)
        except FileNotFoundError:
            QMessageBox.warning(
                self,
                "Preview",
                f"No template for topic '{topic}', language '{lang}'.",
            )
            return

        try:
            subject = tpl["subject"].format(**row)
            body = tpl["body"].format(**row)
        except KeyError as exc:
            QMessageBox.warning(self, "Preview", f"Missing placeholder value: {exc}")
            return

        html_body = template_manager.render_html(
            body, topic=topic, templates_dir=_TEMPLATES_DIR
        )

        preview = QDialog(self)
        preview.setWindowTitle("Email Preview")
        preview.resize(500, 400)
        preview_layout = QVBoxLayout(preview)
        preview_layout.addWidget(QLabel(f"<b>To:</b> {row.get('email', '')}"))
        preview_layout.addWidget(QLabel(f"<b>Subject:</b> {subject}"))
        body_view = QTextEdit()
        body_view.setReadOnly(True)
        body_view.setHtml(html_body)
        preview_layout.addWidget(body_view)
        btn_close = QPushButton("Close")
        btn_close.clicked.connect(preview.accept)
        preview_layout.addWidget(btn_close)
        preview.exec()

    def _on_send_selected(self) -> None:
        """Open a contact picker and send to the chosen contacts."""
        contacts = self._all_contacts()
        if not contacts:
            QMessageBox.information(self, "Send", "No contacts loaded.")
            return

        dlg = _ContactPickerDialog(contacts, multi_select=True, parent=self)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        selected = dlg.selected_contacts()
        if not selected:
            QMessageBox.information(self, "Send", "No contacts selected.")
            return
        self._send_emails(selected)

    def _send_emails(self, rows: list[dict[str, str]]) -> None:
        """Execute the send loop for *rows* using the topic selected in the Send tab."""
        topic = self._send_topic()
        if not topic:
            QMessageBox.warning(self, "Send", "Select a topic first.")
            return

        dry_run = self._dry_run_cb.isChecked()

        # Validate first
        invalid_rows: list[tuple[int, list[str]]] = []
        for i, row in enumerate(rows):
            errors = self._validate_contact(row)
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

            # Load template for this topic + language
            try:
                tpl = template_manager.load_template(topic, lang, _TEMPLATES_DIR)
            except FileNotFoundError:
                self._log_msg(f"SKIPPED {email} — no '{topic}' template for '{lang}'")
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

            html_body = template_manager.render_html(
                body, topic=topic, templates_dir=_TEMPLATES_DIR, use_cid=True
            )

            # Collect image paths for embedding
            import re as _re

            images_dir = _TEMPLATES_DIR / topic / "images"
            image_filenames = _re.findall(r'src="cid:([^"]+)"', html_body)
            image_paths = [
                images_dir / fn for fn in image_filenames if (images_dir / fn).exists()
            ]

            # Send or dry-run
            if dry_run:
                result = mailer.dry_run_email(email, subject, body)
                self._log_msg(f"DRY RUN {email}\n{result}")
                sent += 1
            else:
                try:
                    mailer.send_email(
                        email,
                        subject,
                        html_body,
                        outlook_app=outlook_app,
                        image_paths=image_paths,
                        signature=self._signature_combo.currentData(),
                    )
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
