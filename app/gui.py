"""Main application window with Contacts, Templates, and Send tabs."""

import re
import shutil
from datetime import datetime
from pathlib import Path

from PyQt6.QtCore import QTimer, Qt
from PyQt6.QtGui import QAction, QColor
from PyQt6.QtWidgets import (
    QApplication,
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
    QMenu,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QSizePolicy,
    QTabWidget,
    QTableWidgetItem,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from app import contact_manager, mailer, template_manager
from app.config import DEFAULT_CSV, PROJECT_DIR, TEMPLATES_DIR
from app.widgets import ContactPickerDialog, ExcelTable


class MainWindow(QMainWindow):
    """MailMerge main window with Contacts, Templates, and Send tabs."""

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("MailMerge")
        self.resize(900, 620)

        self._contacts_path: Path = DEFAULT_CSV
        self._rows: list[dict[str, str]] = []
        # Headers excluding "language" — language is determined by sub-tab
        self._headers: list[str] = []
        self._loading_template: bool = False
        self._loading_contacts: bool = False

        tabs = QTabWidget()
        tabs.addTab(self._build_contacts_tab(), "Kontakte")
        tabs.addTab(self._build_templates_tab(), "Vorlagen")
        tabs.addTab(self._build_send_tab(), "Senden")
        self.setCentralWidget(tabs)

        # Contacts auto-save timer (debounce 500ms)
        self._contacts_save_timer = QTimer()
        self._contacts_save_timer.setSingleShot(True)
        self._contacts_save_timer.setInterval(500)
        self._contacts_save_timer.timeout.connect(self._auto_save_contacts)

        # Load initial data
        self._load_csv(self._contacts_path)
        self._refresh_templates()

    # ------------------------------------------------------------------
    # Contacts tab
    # ------------------------------------------------------------------

    def _build_contacts_tab(self) -> QWidget:
        """Build the Contacts tab with search, language sub-tabs, and action buttons."""
        widget = QWidget()
        layout = QVBoxLayout(widget)

        # Top row: search + buttons
        top_layout = QHBoxLayout()

        self._contacts_search = QLineEdit()
        self._contacts_search.setPlaceholderText("Kontakte filtern...")
        self._contacts_search.setClearButtonEnabled(True)
        self._contacts_search.textChanged.connect(self._on_contacts_filter_changed)
        top_layout.addWidget(self._contacts_search)

        btn_import = QPushButton("CSV importieren")
        btn_export = QPushButton("CSV exportieren")
        btn_import.clicked.connect(self._on_import_csv)
        btn_export.clicked.connect(self._on_export_csv)
        for btn in (btn_import, btn_export):
            top_layout.addWidget(btn)

        self._contacts_save_status = QLabel()
        top_layout.addWidget(self._contacts_save_status)

        layout.addLayout(top_layout)

        # Language sub-tabs — one table per language
        self._lang_tabs = QTabWidget()
        self._lang_tables: dict[str, ExcelTable] = {}
        # Track sort state per language tab: {lang: (col, ascending)}
        self._sort_state: dict[str, tuple[int, bool]] = {}
        layout.addWidget(self._lang_tabs)

        return widget

    def _current_lang(self) -> str:
        """Return the language code of the currently selected contacts sub-tab."""
        return self._lang_tabs.tabText(self._lang_tabs.currentIndex())

    def _current_table(self) -> ExcelTable | None:
        """Return the table for the currently selected language tab."""
        lang = self._current_lang()
        return self._lang_tables.get(lang)

    def _make_table(self, lang: str) -> ExcelTable:
        """Create a new table widget for a language tab."""
        table = ExcelTable()
        table.cellChanged.connect(
            lambda row, col: self._on_cell_changed(lang, row, col)
        )
        table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        table.customContextMenuRequested.connect(
            lambda pos, t=table: self._on_table_context_menu(t, pos)
        )
        table.horizontalHeader().sectionClicked.connect(
            lambda col, la=lang: self._on_sort_column(la, col)
        )
        header = table.horizontalHeader()
        header.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        header.customContextMenuRequested.connect(
            lambda pos, t=table: self._on_header_context_menu(t, pos)
        )
        self._lang_tables[lang] = table
        return table

    def _load_csv(self, path: Path) -> None:
        """Load a CSV file into the contacts tables, split by language."""
        try:
            self._rows = contact_manager.load_csv(path)
        except FileNotFoundError:
            QMessageBox.warning(
                self, "Datei nicht gefunden", f"Konnte {path} nicht finden"
            )
            return

        self._contacts_path = path
        if self._rows:
            self._headers = [h for h in self._rows[0].keys() if h != "language"]
        else:
            self._headers = []

        self._loading_contacts = True
        self._rebuild_lang_tabs()
        self._loading_contacts = False

    def _rebuild_lang_tabs(self) -> None:
        """Rebuild the language sub-tabs from current data."""
        self._lang_tabs.blockSignals(True)
        self._lang_tabs.clear()
        self._lang_tables.clear()

        # Union of languages across all topics
        languages = template_manager.list_languages(templates_dir=TEMPLATES_DIR)

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

    def _populate_table(self, table: ExcelTable, rows: list[dict[str, str]]) -> None:
        """Fill a table from a list of row dicts, plus an empty sentinel row at the end."""
        table.blockSignals(True)
        table.setRowCount(len(rows) + 1)  # +1 for sentinel
        table.setColumnCount(len(self._headers))
        table.setHorizontalHeaderLabels(self._headers)

        for r, row in enumerate(rows):
            for c, header in enumerate(self._headers):
                item = QTableWidgetItem(row.get(header, ""))
                table.setItem(r, c, item)

        # Sentinel row — empty placeholder cells
        self._init_sentinel_row(table, len(rows))

        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        table.blockSignals(False)
        self._validate_table(table)

    def _init_sentinel_row(self, table: ExcelTable, row: int) -> None:
        """Set up the empty sentinel row with greyed-out placeholder style."""
        for c in range(len(self._headers)):
            item = QTableWidgetItem("")
            item.setForeground(QColor(180, 180, 180))
            table.setItem(row, c, item)

    def _data_row_count(self, table: ExcelTable) -> int:
        """Return the number of data rows (excluding the sentinel row)."""
        return max(0, table.rowCount() - 1)

    def _lang_for_table(self, table: ExcelTable) -> str:
        """Return the language code associated with a table widget."""
        for lang, t in self._lang_tables.items():
            if t is table:
                return lang
        return ""

    def _validate_row_dict(self, row_dict: dict[str, str], lang: str) -> list[str]:
        """Validate a contact row, injecting the language from the tab."""
        row_with_lang = {**row_dict, "language": lang}
        languages = template_manager.list_languages(templates_dir=TEMPLATES_DIR)
        return contact_manager.validate_row(row_with_lang, languages)

    def _validate_table(self, table: ExcelTable) -> None:
        """Highlight invalid rows in a contacts table (skip sentinel)."""
        err_bg = QColor(255, 80, 80, 60)
        lang = self._lang_for_table(table)

        for r in range(self._data_row_count(table)):
            row_dict = self._table_row_to_dict(table, r)
            errors = self._validate_row_dict(row_dict, lang)
            tooltip = "; ".join(errors) if errors else ""
            for c in range(table.columnCount()):
                item = table.item(r, c)
                if item:
                    if errors:
                        item.setBackground(err_bg)
                    else:
                        item.setData(Qt.ItemDataRole.BackgroundRole, None)
                    item.setToolTip(tooltip)

    def _table_row_to_dict(self, table: ExcelTable, row_index: int) -> dict[str, str]:
        """Convert a table row to a dict keyed by column headers."""
        result: dict[str, str] = {}
        for c, header in enumerate(self._headers):
            item = table.item(row_index, c)
            result[header] = item.text() if item else ""
        return result

    def _all_contacts(self) -> list[dict[str, str]]:
        """Gather all contacts from all language tabs, with language added (skip sentinel)."""
        rows: list[dict[str, str]] = []
        for i in range(self._lang_tabs.count()):
            lang = self._lang_tabs.tabText(i)
            table = self._lang_tables[lang]
            for r in range(self._data_row_count(table)):
                row = self._table_row_to_dict(table, r)
                row["language"] = lang
                rows.append(row)
        return rows

    # -- Contacts slots --

    def _on_import_csv(self) -> None:
        """Open a file dialog and import contacts from a CSV."""
        path, _ = QFileDialog.getOpenFileName(
            self, "CSV importieren", str(PROJECT_DIR), "CSV-Dateien (*.csv)"
        )
        if path:
            self._load_csv(Path(path))
            self._auto_save_contacts()

    def _on_export_csv(self) -> None:
        """Export all contacts to a user-chosen CSV file."""
        path, _ = QFileDialog.getSaveFileName(
            self, "CSV exportieren", str(PROJECT_DIR), "CSV-Dateien (*.csv)"
        )
        if path:
            rows = self._all_contacts()
            contact_manager.save_csv(Path(path), rows)

    def _schedule_contacts_save(self) -> None:
        """Mark contacts as unsaved and restart the debounce timer."""
        if self._loading_contacts:
            return
        self._contacts_save_status.setText("Ungespeichert...")
        self._contacts_save_status.setStyleSheet("color: orange;")
        self._contacts_save_timer.start()

    def _auto_save_contacts(self) -> None:
        """Save all contacts to the default CSV (called by debounce timer)."""
        rows = self._all_contacts()
        contact_manager.save_csv(self._contacts_path, rows)
        self._contacts_save_status.setText("Gespeichert")
        self._contacts_save_status.setStyleSheet("color: gray;")

    def _on_table_context_menu(self, table: ExcelTable, pos) -> None:
        """Show a right-click context menu for the contacts table."""
        row = table.rowAt(pos.y())
        if row < 0 or row >= self._data_row_count(table):
            return

        menu = QMenu(table)
        delete_action = QAction("Zeile loeschen", table)
        delete_action.triggered.connect(lambda: self._delete_row(table, row))
        menu.addAction(delete_action)
        menu.exec(table.viewport().mapToGlobal(pos))

    def _delete_row(self, table: ExcelTable, row: int) -> None:
        """Delete a specific row from a contacts table."""
        if 0 <= row < self._data_row_count(table):
            table.removeRow(row)
            self._validate_table(table)
            self._schedule_contacts_save()

    def _on_add_column(self) -> None:
        """Add a new column (placeholder) to all language tables."""
        name, ok = QInputDialog.getText(
            self, "Spalte hinzufuegen", "Spaltenname (z.B. titel, abteilung):"
        )
        if not ok or not name.strip():
            return
        name = name.strip()
        if name in self._headers:
            QMessageBox.information(
                self, "Spalte hinzufuegen", f"Spalte '{name}' existiert bereits."
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
        self._schedule_contacts_save()

    def _on_header_context_menu(self, table: ExcelTable, pos) -> None:
        """Show a context menu on the column header for add/remove column."""
        header = table.horizontalHeader()
        col = header.logicalIndexAt(pos)
        menu = QMenu(self)

        add_action = QAction("Spalte hinzufuegen", self)
        add_action.triggered.connect(self._on_add_column)
        menu.addAction(add_action)

        if col >= 0:
            col_name = self._headers[col] if col < len(self._headers) else ""
            if col_name.lower() != "email":
                remove_action = QAction(f"Spalte '{col_name}' entfernen", self)
                remove_action.triggered.connect(lambda: self._remove_column(col))
                menu.addAction(remove_action)

        menu.exec(header.mapToGlobal(pos))

    def _remove_column(self, col: int) -> None:
        """Remove a column from all language tables and headers."""
        if col < 0 or col >= len(self._headers):
            return
        col_name = self._headers[col]
        reply = QMessageBox.question(
            self,
            "Spalte entfernen",
            f"Spalte '{col_name}' wirklich aus allen Sprach-Tabs entfernen?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return
        self._headers.pop(col)
        for table in self._lang_tables.values():
            table.blockSignals(True)
            table.removeColumn(col)
            table.setHorizontalHeaderLabels(self._headers)
            table.blockSignals(False)
        self._schedule_contacts_save()

    def _on_cell_changed(self, lang: str, row: int, _col: int) -> None:
        """Re-validate the edited row, promote sentinel if typed into, and auto-save."""
        table = self._lang_tables.get(lang)
        if table is None:
            return

        # If the user typed into the sentinel (last) row, promote it to a data row
        sentinel_row = table.rowCount() - 1
        if row == sentinel_row:
            row_dict = self._table_row_to_dict(table, row)
            if any(v.strip() for v in row_dict.values()):
                table.blockSignals(True)
                # Reset foreground color on the promoted row
                for c in range(table.columnCount()):
                    item = table.item(row, c)
                    if item:
                        item.setData(Qt.ItemDataRole.ForegroundRole, None)
                # Append a new sentinel row
                new_sentinel = table.rowCount()
                table.insertRow(new_sentinel)
                self._init_sentinel_row(table, new_sentinel)
                table.blockSignals(False)
                # Move cursor down to the new sentinel
                col = table.currentColumn()
                QTimer.singleShot(0, lambda: table.setCurrentCell(new_sentinel, col))

        # Skip validation for the sentinel row
        if row >= self._data_row_count(table):
            return

        row_dict = self._table_row_to_dict(table, row)
        errors = self._validate_row_dict(row_dict, lang)
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
        self._schedule_contacts_save()

    def _on_sort_column(self, lang: str, col: int) -> None:
        """Sort a language table by column, keeping the sentinel row pinned at the bottom."""
        table = self._lang_tables.get(lang)
        if table is None:
            return

        # Toggle sort direction
        prev_col, prev_asc = self._sort_state.get(lang, (None, True))
        ascending = not prev_asc if prev_col == col else True
        self._sort_state[lang] = (col, ascending)

        # Collect data rows (exclude sentinel)
        data_count = self._data_row_count(table)
        rows: list[list[str]] = []
        for r in range(data_count):
            row_data = [
                (table.item(r, c).text() if table.item(r, c) else "")
                for c in range(table.columnCount())
            ]
            rows.append(row_data)

        rows.sort(key=lambda row: row[col].lower(), reverse=not ascending)

        # Repopulate data rows
        table.blockSignals(True)
        for r, row_data in enumerate(rows):
            for c, text in enumerate(row_data):
                item = table.item(r, c)
                if item:
                    item.setText(text)
                else:
                    table.setItem(r, c, QTableWidgetItem(text))
        table.blockSignals(False)

        # Update header sort indicator
        table.horizontalHeader().setSortIndicator(
            col,
            Qt.SortOrder.AscendingOrder if ascending else Qt.SortOrder.DescendingOrder,
        )
        table.horizontalHeader().setSortIndicatorShown(True)

        self._validate_table(table)
        self._schedule_contacts_save()

    def _on_contacts_filter_changed(self, text: str) -> None:
        """Show/hide rows across all language tables based on search text."""
        text = text.lower()
        for _lang, table in self._lang_tables.items():
            data_count = self._data_row_count(table)
            for r in range(data_count):
                row_text = " ".join(
                    (table.item(r, c).text() if table.item(r, c) else "")
                    for c in range(table.columnCount())
                ).lower()
                table.setRowHidden(r, text not in row_text)
            # Always show the sentinel row
            sentinel = table.rowCount() - 1
            table.setRowHidden(sentinel, False)

    # ------------------------------------------------------------------
    # Templates tab
    # ------------------------------------------------------------------

    def _build_templates_tab(self) -> QWidget:
        """Build the Templates tab with topic/language selectors and editor fields."""
        widget = QWidget()
        layout = QVBoxLayout(widget)

        # Topic + Language selector row
        sel_layout = QHBoxLayout()

        sel_layout.addWidget(QLabel("Thema:"))
        self._topic_combo = QComboBox()
        self._topic_combo.currentTextChanged.connect(self._on_topic_changed)
        sel_layout.addWidget(self._topic_combo)

        btn_new_topic = QPushButton("Neues Thema")
        btn_new_topic.clicked.connect(self._on_new_topic)
        sel_layout.addWidget(btn_new_topic)

        sel_layout.addWidget(QLabel("Sprache:"))
        self._lang_combo = QComboBox()
        self._lang_combo.currentTextChanged.connect(self._on_template_selection_changed)
        sel_layout.addWidget(self._lang_combo)

        btn_new_lang = QPushButton("Neue Sprache")
        btn_new_lang.clicked.connect(self._on_new_language)
        sel_layout.addWidget(btn_new_lang)

        sel_layout.addStretch()
        layout.addLayout(sel_layout)

        # Subject
        layout.addWidget(QLabel("Betreff:"))
        self._subject_edit = QLineEdit()
        layout.addWidget(self._subject_edit)

        # Body
        layout.addWidget(QLabel("Inhalt:"))

        # Formatting toolbar
        fmt_layout = QHBoxLayout()
        btn_bold = QPushButton("B")
        btn_bold.setFixedWidth(32)
        btn_bold.setStyleSheet("font-weight: bold;")
        btn_bold.setToolTip("Fett")
        btn_bold.clicked.connect(lambda: self._insert_format("**", "**", "Fettschrift"))

        btn_italic = QPushButton("I")
        btn_italic.setFixedWidth(32)
        btn_italic.setStyleSheet("font-style: italic;")
        btn_italic.setToolTip("Kursiv")
        btn_italic.clicked.connect(
            lambda: self._insert_format("*", "*", "Kursivschrift")
        )

        btn_link = QPushButton("Link")
        btn_link.setToolTip("Link einfuegen")
        btn_link.clicked.connect(self._insert_link)

        btn_image = QPushButton("Bild")
        btn_image.setToolTip("Bild einfuegen")
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

        btn_preview_tpl = QPushButton("Vorschau")
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
        topics = template_manager.list_topics(TEMPLATES_DIR)

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
        langs = template_manager.list_languages(topic, TEMPLATES_DIR)
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
            tpl = template_manager.load_template(topic, lang, TEMPLATES_DIR)
        except FileNotFoundError:
            self._subject_edit.clear()
            self._body_edit.clear()
            self._save_status_label.clear()
            self._loading_template = False
            return
        self._subject_edit.setText(tpl["subject"])
        self._body_edit.setPlainText(tpl["body"])
        self._save_status_label.setText("Gespeichert")
        self._save_status_label.setStyleSheet("color: gray;")
        self._loading_template = False

    def _on_template_edited(self) -> None:
        """Mark as unsaved and restart the debounce timer."""
        self._update_placeholder_label()
        if self._loading_template:
            return
        self._save_status_label.setText("Ungespeicherte Aenderungen...")
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
            TEMPLATES_DIR,
        )
        self._save_status_label.setText("Gespeichert")
        self._save_status_label.setStyleSheet("color: gray;")

    def _on_preview_template(self) -> None:
        """Show the current template body rendered as HTML."""
        body = self._body_edit.toPlainText()
        if not body.strip():
            return
        topic = self._topic_combo.currentText()
        html = template_manager.render_html(
            body, topic=topic, templates_dir=TEMPLATES_DIR
        )
        dlg = QDialog(self)
        dlg.setWindowTitle("Vorlagenvorschau")
        dlg.resize(500, 400)
        dlg_layout = QVBoxLayout(dlg)
        subject = self._subject_edit.text()
        if subject:
            dlg_layout.addWidget(QLabel(f"<b>Betreff:</b> {subject}"))
        view = QTextEdit()
        view.setReadOnly(True)
        view.setHtml(html)
        dlg_layout.addWidget(view)
        btn_close = QPushButton("Schliessen")
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
        dlg.setWindowTitle("Link einfuegen")
        dlg_layout = QVBoxLayout(dlg)

        dlg_layout.addWidget(QLabel("Anzeigetext:"))
        text_edit = QLineEdit(selected if selected else "")
        text_edit.setPlaceholderText("z.B. Hier klicken")
        dlg_layout.addWidget(text_edit)

        dlg_layout.addWidget(QLabel("URL:"))
        url_edit = QLineEdit()
        url_edit.setPlaceholderText("z.B. https://example.com")
        dlg_layout.addWidget(url_edit)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(dlg.accept)
        buttons.rejected.connect(dlg.reject)
        dlg_layout.addWidget(buttons)

        if dlg.exec() != QDialog.DialogCode.Accepted:
            return

        link_text = text_edit.text().strip() or "Link"
        url = url_edit.text().strip()
        if not url:
            return

        cursor.insertText(f"[{link_text}]({url})")
        self._body_edit.setFocus()

    def _insert_image(self) -> None:
        """Pick an image file, copy it to the topic's images folder, and insert a reference."""
        topic = self._topic_combo.currentText()
        if not topic:
            QMessageBox.information(self, "Bild", "Bitte zuerst ein Thema auswaehlen.")
            return

        path, _ = QFileDialog.getOpenFileName(
            self,
            "Bild auswaehlen",
            "",
            "Bilder (*.png *.jpg *.jpeg *.gif *.bmp)",
        )
        if not path:
            return

        src = Path(path)
        images_dir = TEMPLATES_DIR / topic / "images"
        images_dir.mkdir(parents=True, exist_ok=True)

        dest = images_dir / src.name
        if dest.exists():
            reply = QMessageBox.question(
                self,
                "Bild vorhanden",
                f"'{src.name}' existiert bereits. Ueberschreiben?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            if reply != QMessageBox.StandardButton.Yes:
                return
        shutil.copy2(src, dest)

        desc, ok = QInputDialog.getText(
            self, "Bildbeschreibung", "Beschreibung (fuer Barrierefreiheit):"
        )
        desc = desc.strip() if ok and desc.strip() else src.stem

        cursor = self._body_edit.textCursor()
        cursor.insertText(f"![{desc}](image:{src.name})")
        self._body_edit.setFocus()

    def _on_new_topic(self) -> None:
        """Prompt for a new topic name."""
        name, ok = QInputDialog.getText(
            self, "Neues Thema", "Themenname (z.B. partnerschaft, nachfassung):"
        )
        if ok and name.strip():
            name = name.strip().lower().replace(" ", "-")
            langs = template_manager.list_languages(templates_dir=TEMPLATES_DIR)
            first_lang = langs[0] if langs else "en"
            template_manager.save_template(name, first_lang, "", "", TEMPLATES_DIR)
            self._refresh_templates()
            self._topic_combo.setCurrentText(name)

    def _on_new_language(self) -> None:
        """Prompt for a new language code within the current topic."""
        topic = self._topic_combo.currentText()
        if not topic:
            return
        code, ok = QInputDialog.getText(
            self, "Neue Sprache", "Sprachcode (z.B. fr, es):"
        )
        if ok and code.strip():
            code = code.strip().lower()
            template_manager.save_template(topic, code, "", "", TEMPLATES_DIR)
            self._refresh_templates()
            self._topic_combo.setCurrentText(topic)
            self._lang_combo.setCurrentText(code)

    def _update_placeholder_label(self) -> None:
        """Show placeholders found in the current subject + body."""
        text = self._subject_edit.text() + "\n" + self._body_edit.toPlainText()
        names = template_manager.extract_placeholders(text)
        if names:
            self._placeholder_label.setText(
                "Platzhalter: " + ", ".join(f"{{{n}}}" for n in names)
            )
        else:
            self._placeholder_label.setText("Keine Platzhalter gefunden.")

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
                "⚠ Outlook ist auf dieser Plattform nicht verfuegbar. "
                "Nur der Testlauf-Modus ist funktionsfaehig."
            )
            notice.setStyleSheet(
                "color: #b45309; background: #fef3c7; padding: 6px; border-radius: 4px;"
            )
            notice.setWordWrap(True)
            layout.addWidget(notice)

        # Topic + Signature selector row
        topic_layout = QHBoxLayout()
        topic_layout.addWidget(QLabel("Thema:"))
        self._send_topic_combo = QComboBox()
        topic_layout.addWidget(self._send_topic_combo)

        topic_layout.addWidget(QLabel("Signatur:"))
        self._signature_combo = QComboBox()
        self._signature_combo.addItem("Keine Signatur", None)
        sigs = mailer.list_signatures()
        for sig in sigs:
            self._signature_combo.addItem(sig, sig)
        if sigs:
            self._signature_combo.setCurrentIndex(1)  # Default to first signature
        topic_layout.addWidget(self._signature_combo)

        topic_layout.addStretch()
        layout.addLayout(topic_layout)

        # Controls row
        ctrl_layout = QHBoxLayout()

        self._dry_run_cb = QCheckBox("Testlauf")
        self._dry_run_cb.setChecked(True)
        self._dry_run_cb.toggled.connect(self._on_dry_run_toggled)
        ctrl_layout.addWidget(self._dry_run_cb)

        btn_preview = QPushButton("Vorschau")
        btn_preview.clicked.connect(self._on_preview)
        ctrl_layout.addWidget(btn_preview)

        self._btn_send_sel = QPushButton("Senden")
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
        self._btn_send_sel.setEnabled(enabled)
        if not enabled:
            self._btn_send_sel.setToolTip(
                "Outlook nicht verfuegbar — Testlauf aktivieren zum Testen"
            )
        else:
            self._btn_send_sel.setToolTip("")

    def _log_msg(self, message: str) -> None:
        """Append a timestamped message to the send log."""
        ts = datetime.now().strftime("%H:%M:%S")
        self._log.append(f"[{ts}] {message}")

    def _on_preview(self) -> None:
        """Open a contact picker, then show a preview for the chosen contact."""
        topic = self._send_topic()
        if not topic:
            QMessageBox.information(
                self, "Vorschau", "Bitte zuerst ein Thema auswaehlen."
            )
            return

        contacts = self._all_contacts()
        if not contacts:
            QMessageBox.information(self, "Vorschau", "Keine Kontakte geladen.")
            return

        dlg = ContactPickerDialog(contacts, multi_select=False, parent=self)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        selected = dlg.selected_contacts()
        if not selected:
            return

        row = selected[0]
        lang = row.get("language", "")

        try:
            tpl = template_manager.load_template(topic, lang, TEMPLATES_DIR)
        except FileNotFoundError:
            QMessageBox.warning(
                self,
                "Vorschau",
                f"Keine Vorlage fuer Thema '{topic}', Sprache '{lang}'.",
            )
            return

        try:
            subject = tpl["subject"].format(**row)
            body = tpl["body"].format(**row)
        except KeyError as exc:
            QMessageBox.warning(self, "Vorschau", f"Fehlender Platzhalterwert: {exc}")
            return

        html_body = template_manager.render_html(
            body, topic=topic, templates_dir=TEMPLATES_DIR
        )

        preview = QDialog(self)
        preview.setWindowTitle("E-Mail-Vorschau")
        preview.resize(500, 400)
        preview_layout = QVBoxLayout(preview)
        preview_layout.addWidget(QLabel(f"<b>An:</b> {row.get('email', '')}"))
        preview_layout.addWidget(QLabel(f"<b>Betreff:</b> {subject}"))
        body_view = QTextEdit()
        body_view.setReadOnly(True)
        body_view.setHtml(html_body)
        preview_layout.addWidget(body_view)
        btn_close = QPushButton("Schliessen")
        btn_close.clicked.connect(preview.accept)
        preview_layout.addWidget(btn_close)
        preview.exec()

    def _on_send_selected(self) -> None:
        """Open a contact picker and send to the chosen contacts."""
        contacts = self._all_contacts()
        if not contacts:
            QMessageBox.information(self, "Senden", "Keine Kontakte geladen.")
            return

        dlg = ContactPickerDialog(contacts, multi_select=True, parent=self)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        selected = dlg.selected_contacts()
        if not selected:
            QMessageBox.information(self, "Senden", "Keine Kontakte ausgewaehlt.")
            return
        self._send_emails(selected)

    def _send_emails(self, rows: list[dict[str, str]]) -> None:
        """Execute the send loop for *rows* using the topic selected in the Send tab."""
        topic = self._send_topic()
        if not topic:
            QMessageBox.warning(self, "Senden", "Bitte zuerst ein Thema auswaehlen.")
            return

        dry_run = self._dry_run_cb.isChecked()

        # Validate first
        languages = template_manager.list_languages(templates_dir=TEMPLATES_DIR)
        invalid_rows: list[tuple[int, list[str]]] = []
        for i, row in enumerate(rows):
            errors = contact_manager.validate_row(row, languages)
            if errors:
                invalid_rows.append((i + 1, errors))

        if invalid_rows:
            msg_lines = [f"Row {idx}: {'; '.join(errs)}" for idx, errs in invalid_rows]
            QMessageBox.warning(
                self,
                "Validierungsfehler",
                "Bitte beheben Sie folgende Fehler:\n\n" + "\n".join(msg_lines),
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
                    "Outlook-Fehler",
                    f"Verbindung zu Outlook fehlgeschlagen:\n{exc}",
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

        for i, row in enumerate(rows):
            email = row.get("email", "")
            lang = row.get("language", "")

            # Load template for this topic + language
            try:
                tpl = template_manager.load_template(topic, lang, TEMPLATES_DIR)
            except FileNotFoundError:
                self._log_msg(
                    f"UEBERSPRUNGEN {email} — keine '{topic}'-Vorlage fuer '{lang}'"
                )
                skipped += 1
                self._progress.setValue(i + 1)
                QApplication.processEvents()
                continue

            # Resolve placeholders
            try:
                subject = tpl["subject"].format(**row)
                body = tpl["body"].format(**row)
            except KeyError as exc:
                self._log_msg(f"UEBERSPRUNGEN {email} — fehlender Platzhalter {exc}")
                skipped += 1
                self._progress.setValue(i + 1)
                QApplication.processEvents()
                continue

            html_body = template_manager.render_html(
                body, topic=topic, templates_dir=TEMPLATES_DIR, use_cid=True
            )

            # Collect image paths for embedding
            images_dir = TEMPLATES_DIR / topic / "images"
            image_filenames = re.findall(r'src="cid:([^"]+)"', html_body)
            image_paths = [
                images_dir / fn for fn in image_filenames if (images_dir / fn).exists()
            ]

            # Send or dry-run
            if dry_run:
                result = mailer.dry_run_email(email, subject, body)
                self._log_msg(f"TESTLAUF {email}\n{result}")
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
                    self._log_msg(f"GESENDET {email}")
                    sent += 1
                except Exception as exc:
                    self._log_msg(f"FEHLER {email} — {exc}")
                    errors += 1

            self._progress.setValue(i + 1)
            QApplication.processEvents()

        mode = "Testlauf" if dry_run else "Versand"
        self._summary_label.setText(
            f"Fertig ({mode}): {sent} gesendet, {skipped} uebersprungen, {errors} Fehler"
        )
