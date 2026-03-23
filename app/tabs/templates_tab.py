"""Templates tab — topic/language template editing with formatting toolbar."""

import shutil
from pathlib import Path

from PyQt6.QtCore import QTimer, pyqtSignal
from PyQt6.QtWidgets import (
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from app import template_manager
from app.config import TEMPLATES_DIR


class TemplatesTab(QWidget):
    """Widget for the Templates tab with topic/language selectors and editor."""

    templates_changed = pyqtSignal()

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._loading_template: bool = False

        layout = QVBoxLayout(self)

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

        self._font_combo = QComboBox()
        self._font_combo.setToolTip("Schriftart")
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
            self._font_combo.addItem(label, css)
        self._font_combo.setCurrentText("Verdana")
        fmt_layout.addWidget(self._font_combo)

        self._font_size_combo = QComboBox()
        self._font_size_combo.setToolTip("Schriftgroesse")
        for size in ["8pt", "9pt", "10pt", "11pt", "12pt", "14pt", "16pt", "18pt"]:
            self._font_size_combo.addItem(size)
        self._font_size_combo.setCurrentText("10pt")
        fmt_layout.addWidget(self._font_size_combo)

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

        self._subject_edit.textChanged.connect(self._on_template_edited)
        self._body_edit.textChanged.connect(self._on_template_edited)

    # -- Public API --

    def font_kwargs(self) -> dict[str, str]:
        """Return the current font settings as kwargs for render_html."""
        return {
            "font_family": self._font_combo.currentData()
            or self._font_combo.currentText(),
            "font_size": self._font_size_combo.currentText(),
        }

    def current_topic(self) -> str:
        """Return the currently selected topic name."""
        return self._topic_combo.currentText()

    def refresh(self) -> list[str]:
        """Reload topic/language combos from disk. Returns the topic list."""
        topics = template_manager.list_topics(TEMPLATES_DIR)

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

        return topics

    # -- Slots --

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
            body, topic=topic, templates_dir=TEMPLATES_DIR, **self.font_kwargs()
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
            self.templates_changed.emit()
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
            self.templates_changed.emit()
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
