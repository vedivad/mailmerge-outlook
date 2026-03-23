"""Layout definition for the Contacts tab."""

from dataclasses import dataclass

from PyQt6.QtWidgets import (
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)


@dataclass
class ContactsWidgets:
    """References to all widgets the ContactsTab controller needs."""

    search: QLineEdit
    btn_import: QPushButton
    btn_export: QPushButton
    save_status: QLabel
    lang_tabs: QTabWidget


def build(parent: QWidget) -> ContactsWidgets:
    """Build the Contacts tab layout on *parent* and return widget references."""
    layout = QVBoxLayout(parent)

    # Top row: search + buttons
    top_layout = QHBoxLayout()

    search = QLineEdit()
    search.setPlaceholderText("Kontakte filtern...")
    search.setClearButtonEnabled(True)
    top_layout.addWidget(search)

    btn_import = QPushButton("CSV importieren")
    btn_export = QPushButton("CSV exportieren")
    for btn in (btn_import, btn_export):
        top_layout.addWidget(btn)

    save_status = QLabel()
    top_layout.addWidget(save_status)

    layout.addLayout(top_layout)

    # Language sub-tabs
    lang_tabs = QTabWidget()
    layout.addWidget(lang_tabs)

    return ContactsWidgets(
        search=search,
        btn_import=btn_import,
        btn_export=btn_export,
        save_status=save_status,
        lang_tabs=lang_tabs,
    )
