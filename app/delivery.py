"""Email delivery backends (Outlook or SMTP) behind a simple interface."""

import mimetypes
import os
import re
import smtplib
from dataclasses import dataclass
from email.message import EmailMessage
from pathlib import Path

from PyQt6.QtCore import QSettings

from app import mailer


@dataclass(frozen=True)
class DeliveryCapabilities:
    """Describes what the active provider can do."""

    provider_name: str
    available: bool
    supports_preview: bool
    supports_draft: bool
    unavailable_reason: str = ""


@dataclass(frozen=True)
class _SmtpSettings:
    """Resolved SMTP settings from environment variables."""

    host: str
    port: int
    from_address: str
    username: str
    password: str
    use_ssl: bool
    use_starttls: bool


def _settings() -> QSettings:
    """Return application settings store."""
    return QSettings("MailMerge", "MailMerge")


def _as_bool(value: object, default: bool) -> bool:
    """Parse a settings/environment value into a bool."""
    if value is None:
        return default
    if isinstance(value, bool):
        return value
    text = str(value).strip().lower()
    if text in {"1", "true", "yes", "on"}:
        return True
    if text in {"0", "false", "no", "off"}:
        return False
    return default


def _value(key: str, env_name: str, default: str) -> str:
    """Read string config from settings first, then environment."""
    setting_value = _settings().value(key, "", type=str)
    if setting_value is not None and setting_value.strip() != "":
        return setting_value.strip()
    env_value = os.getenv(env_name, "").strip()
    if env_value:
        return env_value
    return default


def _provider() -> str:
    """Return the configured provider name."""
    provider = _value("delivery/provider", "MAILMERGE_PROVIDER", "outlook").lower()
    return provider if provider in {"outlook", "smtp"} else "outlook"


def _smtp_settings() -> _SmtpSettings:
    """Read SMTP settings from settings with environment fallback."""
    settings = _settings()

    use_ssl_value = settings.value("delivery/smtp/use_ssl", None)
    if use_ssl_value is None:
        use_ssl_value = os.getenv("MAILMERGE_SMTP_SSL", "0")
    use_ssl = _as_bool(use_ssl_value, default=False)

    default_port = "465" if use_ssl else "587"
    port_value = _value("delivery/smtp/port", "MAILMERGE_SMTP_PORT", default_port)
    try:
        port = int(port_value)
    except ValueError:
        port = int(default_port)

    use_starttls_value = settings.value("delivery/smtp/use_starttls", None)
    if use_starttls_value is None:
        use_starttls_value = os.getenv("MAILMERGE_SMTP_STARTTLS", "1")
    use_starttls = _as_bool(use_starttls_value, default=True)

    return _SmtpSettings(
        host=_value("delivery/smtp/host", "MAILMERGE_SMTP_HOST", ""),
        port=port,
        from_address=_value("delivery/smtp/from", "MAILMERGE_SMTP_FROM", ""),
        username=_value("delivery/smtp/user", "MAILMERGE_SMTP_USER", ""),
        password=_value("delivery/smtp/password", "MAILMERGE_SMTP_PASSWORD", ""),
        use_ssl=use_ssl,
        use_starttls=use_starttls,
    )


def capabilities() -> DeliveryCapabilities:
    """Return capabilities for the currently selected provider."""
    provider = _provider()

    if provider == "smtp":
        settings = _smtp_settings()
        if not settings.host or not settings.from_address:
            return DeliveryCapabilities(
                provider_name="smtp",
                available=False,
                supports_preview=False,
                supports_draft=False,
                unavailable_reason=(
                    "SMTP is not configured. Set MAILMERGE_SMTP_HOST and "
                    "MAILMERGE_SMTP_FROM."
                ),
            )
        if bool(settings.username) != bool(settings.password):
            return DeliveryCapabilities(
                provider_name="smtp",
                available=False,
                supports_preview=False,
                supports_draft=False,
                unavailable_reason=(
                    "SMTP auth is incomplete. Set both MAILMERGE_SMTP_USER and "
                    "MAILMERGE_SMTP_PASSWORD, or neither."
                ),
            )
        return DeliveryCapabilities(
            provider_name="smtp",
            available=True,
            supports_preview=False,
            supports_draft=False,
        )

    if mailer.OUTLOOK_AVAILABLE:
        return DeliveryCapabilities(
            provider_name="outlook",
            available=True,
            supports_preview=True,
            supports_draft=True,
        )

    return DeliveryCapabilities(
        provider_name="outlook",
        available=False,
        supports_preview=False,
        supports_draft=False,
        unavailable_reason=(
            "Outlook is not available on this platform. "
            "pywin32 and Microsoft Outlook are required."
        ),
    )


def create_session() -> object | None:
    """Create a reusable provider session, if supported."""
    active = capabilities()
    if not active.available:
        raise RuntimeError(active.unavailable_reason)

    if active.provider_name == "outlook":
        import win32com.client  # type: ignore[import-untyped]

        return win32com.client.Dispatch("Outlook.Application")

    return None


def dry_run_email(to: str, subject: str, body: str) -> str:
    """Return a dry-run preview string."""
    return mailer.dry_run_email(to, subject, body)


def display_email(
    to: str,
    subject: str,
    html_body: str,
    image_paths: list[Path] | None = None,
    delivery_session: object | None = None,
) -> None:
    """Display an email preview when the provider supports it."""
    active = capabilities()
    if not active.available:
        raise RuntimeError(active.unavailable_reason)
    if not active.supports_preview:
        raise RuntimeError("Preview is only available with the Outlook provider.")

    mailer.display_email(
        to,
        subject,
        html_body,
        outlook_app=delivery_session,
        image_paths=image_paths,
    )


def send_email(
    to: str,
    subject: str,
    html_body: str,
    image_paths: list[Path] | None = None,
    delivery_session: object | None = None,
    draft: bool = False,
) -> None:
    """Send an email through the active provider."""
    active = capabilities()
    if not active.available:
        raise RuntimeError(active.unavailable_reason)

    if active.provider_name == "outlook":
        mailer.send_email(
            to,
            subject,
            html_body,
            outlook_app=delivery_session,
            image_paths=image_paths,
            draft=draft,
        )
        return

    if draft:
        raise RuntimeError("Draft mode is only available with the Outlook provider.")

    _send_smtp_email(to, subject, html_body, image_paths=image_paths)


def _send_smtp_email(
    to: str,
    subject: str,
    html_body: str,
    image_paths: list[Path] | None = None,
) -> None:
    """Send an HTML email through SMTP with optional inline images."""
    settings = _smtp_settings()

    message = EmailMessage()
    message["Subject"] = subject
    message["From"] = settings.from_address
    message["To"] = to

    plain_body = re.sub(r"<[^>]+>", " ", html_body)
    plain_body = re.sub(r"\s+", " ", plain_body).strip() or " "

    message.set_content(plain_body)
    message.add_alternative(html_body, subtype="html")
    html_part = message.get_payload()[-1]

    for image_path in image_paths or []:
        if not image_path.exists():
            continue
        content_type, _encoding = mimetypes.guess_type(image_path.name)
        if content_type:
            maintype, subtype = content_type.split("/", maxsplit=1)
        else:
            maintype, subtype = "application", "octet-stream"

        with image_path.open("rb") as image_file:
            payload = image_file.read()

        html_part.add_related(
            payload,
            maintype=maintype,
            subtype=subtype,
            cid=f"<{image_path.name}>",
            filename=image_path.name,
            disposition="inline",
        )

    server_factory = smtplib.SMTP_SSL if settings.use_ssl else smtplib.SMTP

    with server_factory(settings.host, settings.port, timeout=30) as smtp:
        if not settings.use_ssl and settings.use_starttls:
            smtp.starttls()
        if settings.username and settings.password:
            smtp.login(settings.username, settings.password)
        smtp.send_message(message)
