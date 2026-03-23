"""Layout definition for the Templates tab."""

from dataclasses import dataclass

from PyQt6.QtWidgets import (
    QComboBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)


@dataclass
class TemplatesWidgets:
    """References to all widgets the TemplatesTab controller needs."""

    topic_combo: QComboBox
    lang_combo: QComboBox
    subject_edit: QLineEdit
    body_edit: QTextEdit
    font_combo: QComboBox
    font_size_combo: QComboBox
    btn_bold: QPushButton
    btn_italic: QPushButton
    btn_link: QPushButton
    btn_image: QPushButton
    btn_placeholder: QPushButton
    btn_manage_topics: QPushButton
    btn_manage_langs: QPushButton
    btn_preview: QPushButton
    save_status_label: QLabel
    placeholder_label: QLabel


def build(parent: QWidget) -> TemplatesWidgets:
    """Build the Templates tab layout on *parent* and return widget references."""
    layout = QVBoxLayout(parent)

    # Topic + Language selector row
    sel_layout = QHBoxLayout()

    sel_layout.addWidget(QLabel("Thema:"))
    topic_combo = QComboBox()
    sel_layout.addWidget(topic_combo)

    btn_manage_topics = QPushButton("\u2699")
    btn_manage_topics.setFixedWidth(28)
    btn_manage_topics.setToolTip("Themen verwalten")
    sel_layout.addWidget(btn_manage_topics)

    sel_layout.addWidget(QLabel("Sprache:"))
    lang_combo = QComboBox()
    sel_layout.addWidget(lang_combo)

    btn_manage_langs = QPushButton("\u2699")
    btn_manage_langs.setFixedWidth(28)
    btn_manage_langs.setToolTip("Sprachen verwalten")
    sel_layout.addWidget(btn_manage_langs)

    sel_layout.addStretch()
    layout.addLayout(sel_layout)

    # Subject
    layout.addWidget(QLabel("Betreff:"))
    subject_edit = QLineEdit()
    layout.addWidget(subject_edit)

    # Body
    layout.addWidget(QLabel("Inhalt:"))

    # Formatting toolbar
    fmt_layout = QHBoxLayout()

    font_combo = QComboBox()
    font_combo.setToolTip("Schriftart")
    for label, css in [
        ("Calibri", "'Calibri', Arial, sans-serif"),
        ("Arial", "'Arial', sans-serif"),
        ("Verdana", "'Verdana', sans-serif"),
        ("Tahoma", "'Tahoma', sans-serif"),
        ("Segoe UI", "'Segoe UI', sans-serif"),
        ("Aptos", "'Aptos', sans-serif"),
        ("Trebuchet MS", "'Trebuchet MS', sans-serif"),
        ("Times New Roman", "'Times New Roman', serif"),
        ("Georgia", "'Georgia', serif"),
        ("Courier New", "'Courier New', monospace"),
    ]:
        font_combo.addItem(label, css)
    font_combo.setCurrentText("Verdana")
    fmt_layout.addWidget(font_combo)

    font_size_combo = QComboBox()
    font_size_combo.setToolTip("Schriftgroesse")
    for size in ["8pt", "9pt", "10pt", "11pt", "12pt", "14pt", "16pt", "18pt"]:
        font_size_combo.addItem(size)
    font_size_combo.setCurrentText("10pt")
    fmt_layout.addWidget(font_size_combo)

    btn_bold = QPushButton("B")
    btn_bold.setFixedWidth(32)
    btn_bold.setStyleSheet("font-weight: bold;")
    btn_bold.setToolTip("Fett")

    btn_italic = QPushButton("I")
    btn_italic.setFixedWidth(32)
    btn_italic.setStyleSheet("font-style: italic;")
    btn_italic.setToolTip("Kursiv")

    btn_link = QPushButton("Link")
    btn_link.setToolTip("Link einfuegen")

    btn_image = QPushButton("Bild")
    btn_image.setToolTip("Bild einfuegen")

    btn_placeholder = QPushButton("{x}")
    btn_placeholder.setToolTip("Platzhalter einfuegen")

    for btn in (btn_bold, btn_italic, btn_link, btn_image, btn_placeholder):
        fmt_layout.addWidget(btn)
    fmt_layout.addStretch()
    layout.addLayout(fmt_layout)

    body_edit = QTextEdit()
    layout.addWidget(body_edit)

    # Bottom row: status + placeholder info + preview
    bottom_layout = QHBoxLayout()

    save_status_label = QLabel()
    bottom_layout.addWidget(save_status_label)

    bottom_layout.addStretch()

    placeholder_label = QLabel()
    placeholder_label.setWordWrap(True)
    bottom_layout.addWidget(placeholder_label)

    bottom_layout.addStretch()

    btn_preview = QPushButton("Vorschau")
    bottom_layout.addWidget(btn_preview)

    layout.addLayout(bottom_layout)

    return TemplatesWidgets(
        topic_combo=topic_combo,
        lang_combo=lang_combo,
        subject_edit=subject_edit,
        body_edit=body_edit,
        font_combo=font_combo,
        font_size_combo=font_size_combo,
        btn_bold=btn_bold,
        btn_italic=btn_italic,
        btn_link=btn_link,
        btn_image=btn_image,
        btn_placeholder=btn_placeholder,
        btn_manage_topics=btn_manage_topics,
        btn_manage_langs=btn_manage_langs,
        btn_preview=btn_preview,
        save_status_label=save_status_label,
        placeholder_label=placeholder_label,
    )
