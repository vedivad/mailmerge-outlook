"""Send tab — dry-run, draft, send, and Outlook preview for bulk emails."""

import re
from collections.abc import Callable
from datetime import datetime

from PyQt6.QtCore import QTimer
from PyQt6.QtWidgets import (
    QComboBox,
    QDialog,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QSizePolicy,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from app import contact_manager, mailer, template_manager
from app.config import TEMPLATES_DIR
from app.widgets import ContactPickerDialog


class SendTab(QWidget):
    """Widget for the Send tab with action buttons, progress, and log."""

    def __init__(
        self,
        get_all_contacts: Callable[[], list[dict[str, str]]],
        get_font_kwargs: Callable[[], dict[str, str]],
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self._get_all_contacts = get_all_contacts
        self._get_font_kwargs = get_font_kwargs
        self._send_state: dict = {}

        layout = QVBoxLayout(self)

        # Outlook availability notice
        if not mailer.OUTLOOK_AVAILABLE:
            notice = QLabel(
                "\u26a0 Outlook ist auf dieser Plattform nicht verfuegbar. "
                "Nur der Testlauf-Modus ist funktionsfaehig."
            )
            notice.setStyleSheet(
                "color: #b45309; background: #fef3c7; padding: 6px; border-radius: 4px;"
            )
            notice.setWordWrap(True)
            layout.addWidget(notice)

        # Topic selector row
        topic_layout = QHBoxLayout()
        topic_layout.addWidget(QLabel("Thema:"))
        self._send_topic_combo = QComboBox()
        topic_layout.addWidget(self._send_topic_combo)
        topic_layout.addStretch()
        layout.addLayout(topic_layout)

        # Controls row
        ctrl_layout = QHBoxLayout()

        btn_dry_run = QPushButton("Testlauf")
        btn_dry_run.clicked.connect(lambda: self._on_action("dry_run"))
        ctrl_layout.addWidget(btn_dry_run)

        btn_draft = QPushButton("Entwurf")
        btn_draft.clicked.connect(lambda: self._on_action("draft"))
        btn_draft.setEnabled(mailer.OUTLOOK_AVAILABLE)
        if not mailer.OUTLOOK_AVAILABLE:
            btn_draft.setToolTip("Outlook nicht verfuegbar")
        ctrl_layout.addWidget(btn_draft)

        btn_send = QPushButton("Senden")
        btn_send.clicked.connect(lambda: self._on_action("send"))
        btn_send.setEnabled(mailer.OUTLOOK_AVAILABLE)
        if not mailer.OUTLOOK_AVAILABLE:
            btn_send.setToolTip("Outlook nicht verfuegbar")
        ctrl_layout.addWidget(btn_send)

        ctrl_layout.addStretch()

        btn_preview = QPushButton("Vorschau (Outlook)")
        btn_preview.clicked.connect(self._on_preview)
        btn_preview.setEnabled(mailer.OUTLOOK_AVAILABLE)
        if not mailer.OUTLOOK_AVAILABLE:
            btn_preview.setToolTip("Outlook nicht verfuegbar")
        ctrl_layout.addWidget(btn_preview)
        layout.addLayout(ctrl_layout)

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

    # -- Public API --

    def refresh_topics(self, topics: list[str]) -> None:
        """Update the topic combo with a new list of topics."""
        prev = self._send_topic_combo.currentText()
        self._send_topic_combo.blockSignals(True)
        self._send_topic_combo.clear()
        self._send_topic_combo.addItems(topics)
        self._send_topic_combo.blockSignals(False)
        if prev and prev in topics:
            self._send_topic_combo.setCurrentText(prev)
        elif topics:
            self._send_topic_combo.setCurrentIndex(0)

    # -- Internal helpers --

    def _send_topic(self) -> str:
        """Return the topic selected in the Send tab."""
        return self._send_topic_combo.currentText()

    def _log_msg(self, message: str) -> None:
        """Append a timestamped message to the send log."""
        ts = datetime.now().strftime("%H:%M:%S")
        self._log.append(f"[{ts}] {message}")

    def _set_send_buttons_enabled(self, enabled: bool) -> None:
        """Enable or disable all action buttons during a send loop."""
        for btn in self.findChildren(QPushButton):
            text = btn.text()
            if text in ("Testlauf", "Entwurf", "Senden", "Vorschau (Outlook)"):
                if enabled and text != "Testlauf" and not mailer.OUTLOOK_AVAILABLE:
                    btn.setEnabled(False)
                else:
                    btn.setEnabled(enabled)

    # -- Slots --

    def _on_preview(self) -> None:
        """Open a contact picker, then display the email in Outlook for inspection."""
        topic = self._send_topic()
        if not topic:
            QMessageBox.information(
                self, "Vorschau", "Bitte zuerst ein Thema auswaehlen."
            )
            return

        contacts = self._get_all_contacts()
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
        email = row.get("email", "")
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
            body,
            topic=topic,
            templates_dir=TEMPLATES_DIR,
            use_cid=True,
            **self._get_font_kwargs(),
        )

        images_dir = TEMPLATES_DIR / topic / "images"
        image_filenames = re.findall(r'src="cid:([^"]+)"', html_body)
        image_paths = [
            images_dir / fn for fn in image_filenames if (images_dir / fn).exists()
        ]

        try:
            mailer.display_email(email, subject, html_body, image_paths=image_paths)
            self._log_msg(f"VORSCHAU {email}")
        except Exception as exc:
            QMessageBox.critical(self, "Vorschau", f"Outlook-Fehler:\n{exc}")

    def _on_action(self, mode: str) -> None:
        """Open a contact picker and execute the chosen mode (dry_run/draft/send)."""
        contacts = self._get_all_contacts()
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

        count = len(selected)
        if mode == "draft":
            answer = QMessageBox.question(
                self,
                "Entwurf erstellen",
                f"{count} E-Mail(s) werden als Entwurf in Outlook gespeichert.\n\n"
                "Fortfahren?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if answer != QMessageBox.StandardButton.Yes:
                return
        elif mode == "send":
            answer = QMessageBox.warning(
                self,
                "E-Mails senden",
                f"{count} E-Mail(s) werden jetzt ueber Outlook versendet.\n\n"
                "Dieser Vorgang kann nicht rueckgaengig gemacht werden.\n\n"
                "Fortfahren?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if answer != QMessageBox.StandardButton.Yes:
                return

        self._send_emails(selected, mode)

    def _send_emails(self, rows: list[dict[str, str]], mode: str = "dry_run") -> None:
        """Execute the send loop for *rows*.

        COM calls (draft/send) are spaced 1.5 s apart via a QTimer to prevent
        Outlook from freezing.  Dry-run mode runs without delay.
        """
        topic = self._send_topic()
        if not topic:
            QMessageBox.warning(self, "Senden", "Bitte zuerst ein Thema auswaehlen.")
            return

        dry_run = mode == "dry_run"

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

        total = len(rows)
        self._progress.setMaximum(total)
        self._progress.setValue(0)
        self._log.clear()
        self._summary_label.clear()

        self._send_state = {
            "rows": rows,
            "mode": mode,
            "dry_run": dry_run,
            "topic": topic,
            "outlook_app": outlook_app,
            "index": 0,
            "sent": 0,
            "skipped": 0,
            "errors": 0,
        }
        self._set_send_buttons_enabled(False)
        self._send_step()

    def _send_step(self) -> None:
        """Process one row of the send loop, then schedule the next via QTimer."""
        s = self._send_state
        rows = s["rows"]

        if s["index"] >= len(rows):
            self._send_finish()
            return

        i = s["index"]
        row = rows[i]
        email = row.get("email", "")
        lang = row.get("language", "")
        topic = s["topic"]
        mode = s["mode"]

        try:
            tpl = template_manager.load_template(topic, lang, TEMPLATES_DIR)
        except FileNotFoundError:
            self._log_msg(
                f"UEBERSPRUNGEN {email} — keine '{topic}'-Vorlage fuer '{lang}'"
            )
            s["skipped"] += 1
            self._progress.setValue(i + 1)
            s["index"] += 1
            QTimer.singleShot(0, self._send_step)
            return

        try:
            subject = tpl["subject"].format(**row)
            body = tpl["body"].format(**row)
        except KeyError as exc:
            self._log_msg(f"UEBERSPRUNGEN {email} — fehlender Platzhalter {exc}")
            s["skipped"] += 1
            self._progress.setValue(i + 1)
            s["index"] += 1
            QTimer.singleShot(0, self._send_step)
            return

        html_body = template_manager.render_html(
            body,
            topic=topic,
            templates_dir=TEMPLATES_DIR,
            use_cid=True,
            **self._get_font_kwargs(),
        )

        images_dir = TEMPLATES_DIR / topic / "images"
        image_filenames = re.findall(r'src="cid:([^"]+)"', html_body)
        image_paths = [
            images_dir / fn for fn in image_filenames if (images_dir / fn).exists()
        ]

        if s["dry_run"]:
            result = mailer.dry_run_email(email, subject, body)
            self._log_msg(f"TESTLAUF {email}\n{result}")
            s["sent"] += 1
        else:
            try:
                mailer.send_email(
                    email,
                    subject,
                    html_body,
                    outlook_app=s["outlook_app"],
                    image_paths=image_paths,
                    draft=(mode == "draft"),
                )
                label = "ENTWURF" if mode == "draft" else "GESENDET"
                self._log_msg(f"{label} {email}")
                s["sent"] += 1
            except Exception as exc:
                self._log_msg(f"FEHLER {email} — {exc}")
                s["errors"] += 1

        self._progress.setValue(i + 1)
        s["index"] += 1

        delay_ms = 0 if s["dry_run"] else 1500
        QTimer.singleShot(delay_ms, self._send_step)

    def _send_finish(self) -> None:
        """Show the summary after the send loop completes."""
        s = self._send_state
        mode = s["mode"]
        mode_label = {"dry_run": "Testlauf", "draft": "Entwurf", "send": "Versand"}[
            mode
        ]
        action_label = "gespeichert" if mode == "draft" else "gesendet"
        self._summary_label.setText(
            f"Fertig ({mode_label}): {s['sent']} {action_label},"
            f" {s['skipped']} uebersprungen, {s['errors']} Fehler"
        )
        self._set_send_buttons_enabled(True)
        self._send_state = {}
