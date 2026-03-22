"""Reusable widgets used across the MailMerge GUI."""

from PyQt6.QtCore import QTimer, Qt
from PyQt6.QtWidgets import (
    QCheckBox,
    QDialog,
    QDialogButtonBox,
    QHBoxLayout,
    QHeaderView,
    QLineEdit,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)


class ExcelTable(QTableWidget):
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
