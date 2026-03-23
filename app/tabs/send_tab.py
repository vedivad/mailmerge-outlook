"""Send tab — dry-run, draft, send, and Outlook preview for bulk emails."""

import re
from collections.abc import Callable
from datetime import datetime

from PyQt6.QtCore import QTimer
from PyQt6.QtWidgets import (
    QDialog,
    QMessageBox,
    QWidget,
)

from app import contact_manager, mailer, template_manager
from app.config import TEMPLATES_DIR
from app.tabs import send_ui
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

        # Build UI
        self._ui = send_ui.build(self)

        # Wire signals
        self._ui.btn_dry_run.clicked.connect(lambda: self._on_action("dry_run"))
        self._ui.btn_draft.clicked.connect(lambda: self._on_action("draft"))
        self._ui.btn_send.clicked.connect(lambda: self._on_action("send"))
        self._ui.btn_preview.clicked.connect(self._on_preview)

    # -- Public API --

    def refresh_topics(self, topics: list[str]) -> None:
        """Update the topic combo with a new list of topics."""
        prev = self._ui.topic_combo.currentText()
        self._ui.topic_combo.blockSignals(True)
        self._ui.topic_combo.clear()
        self._ui.topic_combo.addItems(topics)
        self._ui.topic_combo.blockSignals(False)
        if prev and prev in topics:
            self._ui.topic_combo.setCurrentText(prev)
        elif topics:
            self._ui.topic_combo.setCurrentIndex(0)

    # -- Internal helpers --

    def _send_topic(self) -> str:
        """Return the topic selected in the Send tab."""
        return self._ui.topic_combo.currentText()

    def _log_msg(self, message: str) -> None:
        """Append a timestamped message to the send log."""
        ts = datetime.now().strftime("%H:%M:%S")
        self._ui.log.append(f"[{ts}] {message}")

    def _set_send_buttons_enabled(self, enabled: bool) -> None:
        """Enable or disable all action buttons during a send loop."""
        action_buttons = {
            self._ui.btn_dry_run,
            self._ui.btn_draft,
            self._ui.btn_send,
            self._ui.btn_preview,
        }
        for btn in action_buttons:
            if enabled and btn is not self._ui.btn_dry_run and not mailer.OUTLOOK_AVAILABLE:
                btn.setEnabled(False)
            else:
                btn.setEnabled(enabled)

    # -- Slots --

    def _on_preview(self) -> None:
        """Open a contact picker, then display the email in Outlook for inspection."""
        topic = self._send_topic()
        if not topic:
            QMessageBox.information(
                self, self.tr("Preview"), self.tr("Please select a topic first.")
            )
            return

        contacts = self._get_all_contacts()
        if not contacts:
            QMessageBox.information(self, self.tr("Preview"), self.tr("No contacts loaded."))
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
                self.tr("Preview"),
                self.tr("No template for topic '{topic}', language '{lang}'.").format(
                    topic=topic, lang=lang
                ),
            )
            return

        try:
            subject = template_manager.resolve_template(tpl["subject"], row)
            body = template_manager.resolve_template(tpl["body"], row)
        except Exception as exc:
            QMessageBox.warning(
                self,
                self.tr("Preview"),
                self.tr("Template error: {error}").format(error=exc),
            )
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
            self._log_msg(self.tr("PREVIEW {email}").format(email=email))
        except Exception as exc:
            QMessageBox.critical(
                self,
                self.tr("Preview"),
                self.tr("Outlook error:\n{error}").format(error=exc),
            )

    def _on_action(self, mode: str) -> None:
        """Open a contact picker and execute the chosen mode (dry_run/draft/send)."""
        contacts = self._get_all_contacts()
        if not contacts:
            QMessageBox.information(self, self.tr("Send"), self.tr("No contacts loaded."))
            return

        dlg = ContactPickerDialog(contacts, multi_select=True, parent=self)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        selected = dlg.selected_contacts()
        if not selected:
            QMessageBox.information(self, self.tr("Send"), self.tr("No contacts selected."))
            return

        count = len(selected)
        if mode == "draft":
            answer = QMessageBox.question(
                self,
                self.tr("Create drafts"),
                self.tr(
                    "{count} email(s) will be saved as drafts in Outlook.\n\n"
                    "Continue?"
                ).format(count=count),
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if answer != QMessageBox.StandardButton.Yes:
                return
        elif mode == "send":
            answer = QMessageBox.warning(
                self,
                self.tr("Send emails"),
                self.tr(
                    "{count} email(s) will now be sent via Outlook.\n\n"
                    "This action cannot be undone.\n\n"
                    "Continue?"
                ).format(count=count),
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
            QMessageBox.warning(self, self.tr("Send"), self.tr("Please select a topic first."))
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
            msg_lines = [
                self.tr("Row {row}: {errors}").format(row=idx, errors="; ".join(errs))
                for idx, errs in invalid_rows
            ]
            QMessageBox.warning(
                self,
                self.tr("Validation errors"),
                self.tr("Please fix the following errors:") + "\n\n" + "\n".join(msg_lines),
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
                    self.tr("Outlook error"),
                    self.tr("Failed to connect to Outlook:\n{error}").format(error=exc),
                )
                return

        total = len(rows)
        self._ui.progress.setMaximum(total)
        self._ui.progress.setValue(0)
        self._ui.log.clear()
        self._ui.summary_label.clear()

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
                self.tr("SKIPPED {email} — no '{topic}' template for '{lang}'").format(
                    email=email,
                    topic=topic,
                    lang=lang,
                )
            )
            s["skipped"] += 1
            self._ui.progress.setValue(i + 1)
            s["index"] += 1
            QTimer.singleShot(0, self._send_step)
            return

        try:
            subject = template_manager.resolve_template(tpl["subject"], row)
            body = template_manager.resolve_template(tpl["body"], row)
        except Exception as exc:
            self._log_msg(
                self.tr("SKIPPED {email} — template error: {error}").format(
                    email=email,
                    error=exc,
                )
            )
            s["skipped"] += 1
            self._ui.progress.setValue(i + 1)
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
            self._log_msg(
                self.tr("DRY RUN {email}\n{result}").format(email=email, result=result)
            )
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
                label = self.tr("DRAFT") if mode == "draft" else self.tr("SENT")
                self._log_msg(self.tr("{label} {email}").format(label=label, email=email))
                s["sent"] += 1
            except Exception as exc:
                self._log_msg(
                    self.tr("ERROR {email} — {error}").format(email=email, error=exc)
                )
                s["errors"] += 1

        self._ui.progress.setValue(i + 1)
        s["index"] += 1

        delay_ms = 0 if s["dry_run"] else 1500
        QTimer.singleShot(delay_ms, self._send_step)

    def _send_finish(self) -> None:
        """Show the summary after the send loop completes."""
        s = self._send_state
        mode = s["mode"]
        mode_label = {
            "dry_run": self.tr("Dry run"),
            "draft": self.tr("Draft"),
            "send": self.tr("Send"),
        }[mode]
        action_label = self.tr("saved") if mode == "draft" else self.tr("sent")
        self._ui.summary_label.setText(
            self.tr("Done ({mode}): {sent} {action}, {skipped} skipped, {errors} errors").format(
                mode=mode_label,
                sent=s["sent"],
                action=action_label,
                skipped=s["skipped"],
                errors=s["errors"],
            )
        )
        self._set_send_buttons_enabled(True)
        self._send_state = {}
