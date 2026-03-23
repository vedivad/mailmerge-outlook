"""Reusable widgets used across the MailMerge GUI."""

from PyQt6.QtCore import QTimer, Qt
from PyQt6.QtGui import QKeySequence
from PyQt6.QtWidgets import (
    QApplication,
    QCheckBox,
    QDialog,
    QDialogButtonBox,
    QHBoxLayout,
    QHeaderView,
    QLineEdit,
    QListWidget,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)


class ExcelTable(QTableWidget):
    """QTableWidget with Enter-to-move-down, copy, and paste support."""

    enter_pressed: bool = False
    """True when the most recent commit was triggered by Enter/Return."""

    def keyPressEvent(self, event) -> None:  # noqa: N802
        """Handle Enter, Ctrl+C, and Ctrl+V."""
        if event.matches(QKeySequence.StandardKey.Paste):
            self._paste_clipboard()
        elif event.matches(QKeySequence.StandardKey.Copy):
            self._copy_selection()
        elif event.key() in (Qt.Key.Key_Return, Qt.Key.Key_Enter):
            self.enter_pressed = True
            row = self.currentRow()
            col = self.currentColumn()
            super().keyPressEvent(event)
            next_row = row + 1
            if next_row < self.rowCount():
                QTimer.singleShot(0, lambda: self.setCurrentCell(next_row, col))
        else:
            self.enter_pressed = False
            super().keyPressEvent(event)

    def _copy_selection(self) -> None:
        """Copy selected cells to clipboard as tab-separated text."""
        selection = self.selectedRanges()
        if not selection:
            return
        sr = selection[0]
        lines: list[str] = []
        for r in range(sr.topRow(), sr.bottomRow() + 1):
            row_texts: list[str] = []
            for c in range(sr.leftColumn(), sr.rightColumn() + 1):
                item = self.item(r, c)
                row_texts.append(item.text() if item else "")
            lines.append("\t".join(row_texts))
        clipboard = QApplication.clipboard()
        if clipboard:
            clipboard.setText("\n".join(lines))

    def _paste_clipboard(self) -> None:
        """Paste tab-separated clipboard text into cells starting at the current cell."""
        clipboard = QApplication.clipboard()
        if not clipboard:
            return
        text = clipboard.text()
        if not text:
            return
        start_row = self.currentRow()
        start_col = self.currentColumn()
        for r, line in enumerate(text.split("\n")):
            if not line:
                continue
            for c, value in enumerate(line.split("\t")):
                target_row = start_row + r
                target_col = start_col + c
                if target_row < self.rowCount() and target_col < self.columnCount():
                    item = self.item(target_row, target_col)
                    if item:
                        item.setText(value)
                    else:
                        self.setItem(target_row, target_col, QTableWidgetItem(value))


class ContactPickerDialog(QDialog):
    """Dialog for picking one or more contacts with search filtering."""

    def __init__(
        self,
        contacts: list[dict[str, str]],
        multi_select: bool = True,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle(
            "Kontakte auswaehlen" if multi_select else "Kontakt auswaehlen"
        )
        self.resize(600, 450)
        self._contacts = contacts
        self._multi_select = multi_select
        self._checkboxes: list[QCheckBox] = []
        self._selected: list[dict[str, str]] = []

        layout = QVBoxLayout(self)

        # Search bar
        self._search = QLineEdit()
        self._search.setPlaceholderText("Kontakte filtern...")
        self._search.textChanged.connect(self._apply_filter)
        layout.addWidget(self._search)

        # Select all (multi-select only)
        if multi_select:
            self._select_all_cb = QCheckBox("Alle auswaehlen")
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


class ColumnReorderDialog(QDialog):
    """Dialog for reordering table columns via up/down buttons."""

    def __init__(self, headers: list[str], parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Spalten ordnen")
        self.resize(350, 400)
        self._result_order: list[str] = list(headers)

        layout = QHBoxLayout(self)

        # Column list
        self._list = QListWidget()
        self._list.addItems(headers)
        if headers:
            self._list.setCurrentRow(0)
        layout.addWidget(self._list)

        # Button column
        btn_layout = QVBoxLayout()
        btn_layout.addStretch()

        btn_up = QPushButton("\u25b2")
        btn_up.setToolTip("Nach oben")
        btn_up.setFixedWidth(40)
        btn_up.clicked.connect(self._move_up)
        btn_layout.addWidget(btn_up)

        btn_down = QPushButton("\u25bc")
        btn_down.setToolTip("Nach unten")
        btn_down.setFixedWidth(40)
        btn_down.clicked.connect(self._move_down)
        btn_layout.addWidget(btn_down)

        btn_layout.addStretch()

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self._on_accept)
        buttons.rejected.connect(self.reject)
        btn_layout.addWidget(buttons)

        layout.addLayout(btn_layout)

    def _move_up(self) -> None:
        """Move the selected item up one position."""
        row = self._list.currentRow()
        if row <= 0:
            return
        item = self._list.takeItem(row)
        self._list.insertItem(row - 1, item)
        self._list.setCurrentRow(row - 1)

    def _move_down(self) -> None:
        """Move the selected item down one position."""
        row = self._list.currentRow()
        if row < 0 or row >= self._list.count() - 1:
            return
        item = self._list.takeItem(row)
        self._list.insertItem(row + 1, item)
        self._list.setCurrentRow(row + 1)

    def _on_accept(self) -> None:
        """Store the new order and accept."""
        self._result_order = [
            self._list.item(i).text() for i in range(self._list.count())
        ]
        self.accept()

    def result_order(self) -> list[str]:
        """Return the column names in the user-chosen order."""
        return self._result_order


class ListManagerDialog(QDialog):
    """Dialog for managing a list of named items (add, rename, delete)."""

    def __init__(
        self,
        title: str,
        items: list[str],
        add_label: str = "Hinzufuegen",
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle(title)
        self.resize(350, 350)
        self._changed = False

        layout = QHBoxLayout(self)

        self._list = QListWidget()
        self._list.addItems(items)
        if items:
            self._list.setCurrentRow(0)
        layout.addWidget(self._list)

        btn_layout = QVBoxLayout()

        btn_add = QPushButton(add_label)
        btn_add.clicked.connect(self._on_add)
        btn_layout.addWidget(btn_add)

        btn_rename = QPushButton("Umbenennen")
        btn_rename.clicked.connect(self._on_rename)
        btn_layout.addWidget(btn_rename)

        btn_delete = QPushButton("Loeschen")
        btn_delete.clicked.connect(self._on_delete)
        btn_layout.addWidget(btn_delete)

        btn_layout.addStretch()

        btn_close = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        btn_close.rejected.connect(self.accept)
        btn_layout.addWidget(btn_close)

        layout.addLayout(btn_layout)

        self._add_label = add_label

    def _on_add(self) -> None:
        """Prompt for a new item name and add it."""
        from PyQt6.QtWidgets import QInputDialog

        name, ok = QInputDialog.getText(self, self._add_label, "Name:")
        if not ok or not name.strip():
            return
        name = name.strip().lower().replace(" ", "-")
        existing = [self._list.item(i).text() for i in range(self._list.count())]
        if name in existing:
            QMessageBox.information(
                self, self._add_label, f"'{name}' existiert bereits."
            )
            return
        self._list.addItem(name)
        self._list.setCurrentRow(self._list.count() - 1)
        self._changed = True

    def _on_rename(self) -> None:
        """Rename the selected item."""
        from PyQt6.QtWidgets import QInputDialog

        row = self._list.currentRow()
        if row < 0:
            return
        old_name = self._list.item(row).text()
        new_name, ok = QInputDialog.getText(
            self, "Umbenennen", f"Neuer Name fuer '{old_name}':", text=old_name
        )
        if not ok or not new_name.strip():
            return
        new_name = new_name.strip().lower().replace(" ", "-")
        if new_name == old_name:
            return
        existing = [self._list.item(i).text() for i in range(self._list.count())]
        if new_name in existing:
            QMessageBox.information(
                self, "Umbenennen", f"'{new_name}' existiert bereits."
            )
            return
        self._list.item(row).setText(new_name)
        self._changed = True

    def _on_delete(self) -> None:
        """Delete the selected item."""
        row = self._list.currentRow()
        if row < 0:
            return
        name = self._list.item(row).text()
        reply = QMessageBox.question(
            self,
            "Loeschen",
            f"'{name}' wirklich loeschen?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return
        self._list.takeItem(row)
        self._changed = True

    def result_items(self) -> list[str]:
        """Return the current list of item names."""
        return [self._list.item(i).text() for i in range(self._list.count())]

    def was_changed(self) -> bool:
        """Return True if any add/rename/delete was performed."""
        return self._changed
