"""Contacts tab — CSV-backed contact management with language sub-tabs."""

from pathlib import Path

from PyQt6.QtCore import QTimer, Qt
from PyQt6.QtGui import QAction, QColor
from PyQt6.QtWidgets import (
    QCheckBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QHBoxLayout,
    QHeaderView,
    QInputDialog,
    QLabel,
    QLineEdit,
    QMenu,
    QMessageBox,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from app import contact_manager, template_manager
from app.config import DEFAULT_CSV, PROJECT_DIR, TEMPLATES_DIR
from app.tabs import contacts_ui
from app.widgets import ColumnReorderDialog, ExcelTable


class ContactsTab(QWidget):
    """Widget for the Contacts tab with search, language sub-tabs, and CSV I/O."""

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._contacts_path: Path = DEFAULT_CSV
        self._rows: list[dict[str, str]] = []
        self._headers: list[str] = []
        self._loading_contacts: bool = False
        self._sort_state: dict[str, tuple[int, bool]] = {}
        self._lang_tables: dict[str, ExcelTable] = {}

        # Build UI
        self._ui = contacts_ui.build(self)

        # Wire signals
        self._ui.search.textChanged.connect(self._on_contacts_filter_changed)
        self._ui.btn_import.clicked.connect(self._on_import_csv)
        self._ui.btn_export.clicked.connect(self._on_export_csv)
        self._ui.btn_reorder.clicked.connect(self._on_reorder_columns)

        # Auto-save timer (debounce 500ms)
        self._save_timer = QTimer()
        self._save_timer.setSingleShot(True)
        self._save_timer.setInterval(500)
        self._save_timer.timeout.connect(self._auto_save_contacts)

    # -- Public API --

    def load_csv(self, path: Path) -> None:
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
        self.rebuild_lang_tabs()
        self._loading_contacts = False

    def rebuild_lang_tabs(self) -> None:
        """Rebuild the language sub-tabs from current data."""
        self._ui.lang_tabs.blockSignals(True)
        self._ui.lang_tabs.clear()
        self._lang_tables.clear()

        languages = template_manager.list_languages(templates_dir=TEMPLATES_DIR)

        rows_by_lang: dict[str, list[dict[str, str]]] = {lang: [] for lang in languages}
        for row in self._rows:
            lang = row.get("language", "")
            if lang not in rows_by_lang:
                rows_by_lang[lang] = []
            rows_by_lang[lang].append(row)

        for lang in languages:
            table = self._make_table(lang)
            self._populate_table(table, rows_by_lang.get(lang, []))
            self._ui.lang_tabs.addTab(table, lang)

        other_rows = []
        for lang, rows in rows_by_lang.items():
            if lang not in languages:
                other_rows.extend(rows)
        if other_rows:
            table = self._make_table("other")
            self._populate_table(table, other_rows)
            self._ui.lang_tabs.addTab(table, "other")

        self._ui.lang_tabs.blockSignals(False)

    def headers(self) -> list[str]:
        """Return the current column headers (excluding 'language')."""
        return list(self._headers)

    def all_contacts(self) -> list[dict[str, str]]:
        """Gather all contacts from all language tabs, with language added."""
        rows: list[dict[str, str]] = []
        for i in range(self._ui.lang_tabs.count()):
            lang = self._ui.lang_tabs.tabText(i)
            table = self._lang_tables[lang]
            for r in range(self._data_row_count(table)):
                row = self._table_row_to_dict(table, r)
                row["language"] = lang
                rows.append(row)
        return rows

    # -- Internal helpers --

    def _make_table(self, lang: str) -> ExcelTable:
        """Create a new table widget for a language tab."""
        table = ExcelTable()
        table.cellChanged.connect(
            lambda row, col, la=lang: self._on_cell_changed(la, row, col)
        )
        vheader = table.verticalHeader()
        vheader.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        vheader.customContextMenuRequested.connect(
            lambda pos, t=table: self._on_row_header_context_menu(t, pos)
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

    def _is_bool_column(self, col: int) -> bool:
        """Return True if the column at *col* is a boolean column (name ends with ?)."""
        return col < len(self._headers) and self._headers[col].endswith("?")

    def _set_checkbox(
        self, table: ExcelTable, row: int, col: int, checked: bool
    ) -> None:
        """Place a centered checkbox widget in a table cell."""
        cb = QCheckBox()
        cb.setChecked(checked)
        cb.stateChanged.connect(lambda: self._schedule_contacts_save())
        container = QWidget()
        lay = QHBoxLayout(container)
        lay.addWidget(cb)
        lay.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lay.setContentsMargins(0, 0, 0, 0)
        table.setCellWidget(row, col, container)
        # Set a dummy item so column count logic still works
        item = QTableWidgetItem()
        item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
        table.setItem(row, col, item)

    def _get_checkbox(self, table: ExcelTable, row: int, col: int) -> bool:
        """Read the checked state of a checkbox cell widget."""
        widget = table.cellWidget(row, col)
        if widget:
            cb = widget.findChild(QCheckBox)
            if cb:
                return cb.isChecked()
        return False

    def _populate_table(self, table: ExcelTable, rows: list[dict[str, str]]) -> None:
        """Fill a table from a list of row dicts, plus an empty sentinel row."""
        table.blockSignals(True)
        table.setRowCount(len(rows) + 1)
        table.setColumnCount(len(self._headers))
        table.setHorizontalHeaderLabels(self._headers)

        for r, row in enumerate(rows):
            for c, header in enumerate(self._headers):
                if self._is_bool_column(c):
                    val = row.get(header, "").strip().lower()
                    self._set_checkbox(
                        table, r, c, val in ("true", "1", "ja", "yes", "x")
                    )
                else:
                    item = QTableWidgetItem(row.get(header, ""))
                    table.setItem(r, c, item)

        self._init_sentinel_row(table, len(rows))
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        table.verticalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Fixed)
        table.blockSignals(False)
        self._validate_table(table)

    def _init_sentinel_row(self, table: ExcelTable, row: int) -> None:
        """Set up the empty sentinel row with greyed-out placeholder style."""
        for c in range(len(self._headers)):
            if self._is_bool_column(c):
                self._set_checkbox(table, row, c, checked=False)
            else:
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

    def _table_row_to_dict(self, table: ExcelTable, row_index: int) -> dict[str, str]:
        """Convert a table row to a dict keyed by column headers."""
        result: dict[str, str] = {}
        for c, header in enumerate(self._headers):
            if self._is_bool_column(c):
                result[header] = (
                    "true" if self._get_checkbox(table, row_index, c) else "false"
                )
            else:
                item = table.item(row_index, c)
                result[header] = item.text() if item else ""
        return result

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

    def _schedule_contacts_save(self) -> None:
        """Mark contacts as unsaved and restart the debounce timer."""
        if self._loading_contacts:
            return
        self._ui.save_status.setText("Ungespeichert...")
        self._ui.save_status.setStyleSheet("color: orange;")
        self._save_timer.start()

    def _auto_save_contacts(self) -> None:
        """Save all contacts to the default CSV (called by debounce timer)."""
        rows = self.all_contacts()
        contact_manager.save_csv(self._contacts_path, rows)
        self._ui.save_status.setText("Gespeichert")
        self._ui.save_status.setStyleSheet("color: gray;")

    # -- Slots --

    def _on_import_csv(self) -> None:
        """Open a file dialog and import contacts from a CSV."""
        path, _ = QFileDialog.getOpenFileName(
            self, "CSV importieren", str(PROJECT_DIR), "CSV-Dateien (*.csv)"
        )
        if path:
            self.load_csv(Path(path))
            self._auto_save_contacts()

    def _on_export_csv(self) -> None:
        """Export all contacts to a user-chosen CSV file."""
        path, _ = QFileDialog.getSaveFileName(
            self, "CSV exportieren", str(PROJECT_DIR), "CSV-Dateien (*.csv)"
        )
        if path:
            rows = self.all_contacts()
            contact_manager.save_csv(Path(path), rows)

    def _on_row_header_context_menu(self, table: ExcelTable, pos) -> None:
        """Show a right-click context menu on the row number header."""
        clicked_row = table.verticalHeader().logicalIndexAt(pos)
        if clicked_row >= 0 and clicked_row not in {
            idx.row() for idx in table.selectedIndexes()
        }:
            table.selectRow(clicked_row)

        selected_rows = sorted(
            {
                idx.row()
                for idx in table.selectedIndexes()
                if idx.row() < self._data_row_count(table)
            }
        )
        if not selected_rows:
            return

        menu = QMenu(table)
        if len(selected_rows) == 1:
            label = "Zeile loeschen"
        else:
            label = f"{len(selected_rows)} Zeilen loeschen"
        delete_action = QAction(label, table)
        delete_action.triggered.connect(lambda: self._delete_rows(table, selected_rows))
        menu.addAction(delete_action)
        menu.exec(table.verticalHeader().mapToGlobal(pos))

    def _delete_rows(self, table: ExcelTable, rows: list[int]) -> None:
        """Delete one or more rows from a contacts table."""
        for row in reversed(rows):
            if 0 <= row < self._data_row_count(table):
                table.removeRow(row)
        self._validate_table(table)
        self._schedule_contacts_save()

    def _on_add_column(self) -> None:
        """Add a new column (placeholder) to all language tables."""
        dlg = QDialog(self)
        dlg.setWindowTitle("Spalte hinzufuegen")
        layout = QVBoxLayout(dlg)

        layout.addWidget(QLabel("Spaltenname (z.B. titel, abteilung):"))
        name_edit = QLineEdit()
        layout.addWidget(name_edit)

        bool_cb = QCheckBox("Ja/Nein-Spalte (Boolean)")
        layout.addWidget(bool_cb)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(dlg.accept)
        buttons.rejected.connect(dlg.reject)
        layout.addWidget(buttons)

        if dlg.exec() != QDialog.DialogCode.Accepted:
            return

        name = name_edit.text().strip()
        if not name:
            return
        if bool_cb.isChecked() and not name.endswith("?"):
            name = name + "?"

        if name in self._headers:
            QMessageBox.information(
                self, "Spalte hinzufuegen", f"Spalte '{name}' existiert bereits."
            )
            return

        is_bool = name.endswith("?")
        self._headers.append(name)
        for table in self._lang_tables.values():
            table.blockSignals(True)
            col = table.columnCount()
            table.setColumnCount(col + 1)
            table.setHorizontalHeaderLabels(self._headers)
            for r in range(table.rowCount()):
                if is_bool:
                    self._set_checkbox(table, r, col, checked=False)
                else:
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

            rename_action = QAction(f"Spalte '{col_name}' umbenennen", self)
            rename_action.triggered.connect(lambda: self._rename_column(col))
            menu.addAction(rename_action)

            if col_name.lower() != "email":
                remove_action = QAction(f"Spalte '{col_name}' entfernen", self)
                remove_action.triggered.connect(lambda: self._remove_column(col))
                menu.addAction(remove_action)

        menu.exec(header.mapToGlobal(pos))

    def _on_reorder_columns(self) -> None:
        """Open the column reorder dialog and apply the result."""
        if not self._headers:
            return
        dialog = ColumnReorderDialog(self._headers, parent=self)
        if dialog.exec() != ColumnReorderDialog.DialogCode.Accepted:
            return
        new_order = dialog.result_order()
        if new_order == self._headers:
            return
        self._apply_column_order(new_order)

    def _apply_column_order(self, new_order: list[str]) -> None:
        """Rearrange columns in all language tables to match *new_order*."""
        index_map = [self._headers.index(name) for name in new_order]
        self._headers = new_order

        for table in self._lang_tables.values():
            table.blockSignals(True)
            for r in range(table.rowCount()):
                old_texts = [
                    (table.item(r, c).text() if table.item(r, c) else "")
                    for c in range(table.columnCount())
                ]
                for c, old_c in enumerate(index_map):
                    item = table.item(r, c)
                    if item:
                        item.setText(old_texts[old_c])
                    else:
                        table.setItem(r, c, QTableWidgetItem(old_texts[old_c]))
            table.setHorizontalHeaderLabels(self._headers)
            table.blockSignals(False)
            self._validate_table(table)

        self._schedule_contacts_save()

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
            table.horizontalHeader().setSectionResizeMode(
                QHeaderView.ResizeMode.Stretch
            )
            table.blockSignals(False)
        self._schedule_contacts_save()

    def _rename_column(self, col: int) -> None:
        """Rename a column across all language tables and headers."""
        if col < 0 or col >= len(self._headers):
            return
        old_name = self._headers[col]
        new_name, ok = QInputDialog.getText(
            self,
            "Spalte umbenennen",
            f"Neuer Name fuer Spalte '{old_name}':",
            text=old_name,
        )
        if not ok or not new_name.strip():
            return
        new_name = new_name.strip()
        if old_name.endswith("?") and not new_name.endswith("?"):
            new_name = new_name + "?"
        if new_name == old_name:
            return
        if new_name in self._headers:
            QMessageBox.information(
                self, "Spalte umbenennen", f"Spalte '{new_name}' existiert bereits."
            )
            return
        self._headers[col] = new_name
        for table in self._lang_tables.values():
            table.blockSignals(True)
            table.setHorizontalHeaderLabels(self._headers)
            table.blockSignals(False)
        self._schedule_contacts_save()

    def _on_cell_changed(self, lang: str, row: int, _col: int) -> None:
        """Re-validate the edited row, promote sentinel if typed into, and auto-save."""
        table = self._lang_tables.get(lang)
        if table is None:
            return

        sentinel_row = table.rowCount() - 1
        if row == sentinel_row:
            row_dict = self._table_row_to_dict(table, row)
            if any(v.strip() for v in row_dict.values()):
                table.blockSignals(True)
                for c in range(table.columnCount()):
                    item = table.item(row, c)
                    if item:
                        item.setData(Qt.ItemDataRole.ForegroundRole, None)
                new_sentinel = table.rowCount()
                table.insertRow(new_sentinel)
                self._init_sentinel_row(table, new_sentinel)
                table.blockSignals(False)
                if table.enter_pressed:
                    col = table.currentColumn()
                    QTimer.singleShot(
                        0, lambda: table.setCurrentCell(new_sentinel, col)
                    )

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
        """Sort a language table by column, keeping the sentinel row pinned."""
        table = self._lang_tables.get(lang)
        if table is None:
            return

        prev_col, prev_asc = self._sort_state.get(lang, (None, True))
        ascending = not prev_asc if prev_col == col else True
        self._sort_state[lang] = (col, ascending)

        data_count = self._data_row_count(table)
        rows: list[list[str]] = []
        for r in range(data_count):
            row_data = [
                (table.item(r, c).text() if table.item(r, c) else "")
                for c in range(table.columnCount())
            ]
            rows.append(row_data)

        rows.sort(key=lambda row: row[col].lower(), reverse=not ascending)

        table.blockSignals(True)
        for r, row_data in enumerate(rows):
            for c, text in enumerate(row_data):
                item = table.item(r, c)
                if item:
                    item.setText(text)
                else:
                    table.setItem(r, c, QTableWidgetItem(text))
        table.blockSignals(False)

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
            sentinel = table.rowCount() - 1
            table.setRowHidden(sentinel, False)
