"""Layout definition for the Send tab."""

from dataclasses import dataclass

from PyQt6.QtWidgets import (
    QComboBox,
    QHBoxLayout,
    QLabel,
    QProgressBar,
    QPushButton,
    QSizePolicy,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from app import delivery


@dataclass
class SendWidgets:
    """References to all widgets the SendTab controller needs."""

    topic_combo: QComboBox
    btn_dry_run: QPushButton
    btn_draft: QPushButton
    btn_send: QPushButton
    btn_preview: QPushButton
    progress: QProgressBar
    log: QTextEdit
    summary_label: QLabel


def build(parent: QWidget) -> SendWidgets:
    """Build the Send tab layout on *parent* and return widget references."""
    layout = QVBoxLayout(parent)
    caps = delivery.capabilities()

    # Provider availability notice
    if not caps.available:
        notice = QLabel(
            "\u26a0 "
            + parent.tr(
                "Email delivery backend is unavailable. "
                "Only dry-run mode is functional."
            )
        )
        if caps.unavailable_reason:
            notice.setText(notice.text() + "\n" + caps.unavailable_reason)
        notice.setStyleSheet(
            "color: #b45309; background: #fef3c7; padding: 6px; border-radius: 4px;"
        )
        notice.setWordWrap(True)
        layout.addWidget(notice)

    # Topic selector row
    topic_layout = QHBoxLayout()
    topic_layout.addWidget(QLabel(parent.tr("Topic:")))
    topic_combo = QComboBox()
    topic_layout.addWidget(topic_combo)
    topic_layout.addStretch()
    layout.addLayout(topic_layout)

    # Controls row
    ctrl_layout = QHBoxLayout()

    btn_dry_run = QPushButton(parent.tr("Dry run"))
    ctrl_layout.addWidget(btn_dry_run)

    btn_draft = QPushButton(parent.tr("Draft"))
    btn_draft.setEnabled(caps.available and caps.supports_draft)
    if not caps.available:
        btn_draft.setToolTip(caps.unavailable_reason)
    elif not caps.supports_draft:
        btn_draft.setToolTip(parent.tr("Draft mode is only available with Outlook"))
    ctrl_layout.addWidget(btn_draft)

    btn_send = QPushButton(parent.tr("Send"))
    btn_send.setEnabled(caps.available)
    if not caps.available:
        btn_send.setToolTip(caps.unavailable_reason)
    ctrl_layout.addWidget(btn_send)

    ctrl_layout.addStretch()

    btn_preview = QPushButton(parent.tr("Preview"))
    btn_preview.setEnabled(caps.available and caps.supports_preview)
    if not caps.available:
        btn_preview.setToolTip(caps.unavailable_reason)
    elif not caps.supports_preview:
        btn_preview.setToolTip(parent.tr("Preview is only available with Outlook"))
    ctrl_layout.addWidget(btn_preview)
    layout.addLayout(ctrl_layout)

    # Progress bar
    progress = QProgressBar()
    progress.setValue(0)
    layout.addWidget(progress)

    # Log panel
    log = QTextEdit()
    log.setReadOnly(True)
    log.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
    layout.addWidget(log)

    # Summary
    summary_label = QLabel()
    layout.addWidget(summary_label)

    return SendWidgets(
        topic_combo=topic_combo,
        btn_dry_run=btn_dry_run,
        btn_draft=btn_draft,
        btn_send=btn_send,
        btn_preview=btn_preview,
        progress=progress,
        log=log,
        summary_label=summary_label,
    )
