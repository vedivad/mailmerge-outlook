"""Outlook COM automation for sending emails.

On platforms where ``pywin32`` is not available the module still imports
successfully — ``OUTLOOK_AVAILABLE`` will be ``False`` and ``send_email``
will raise ``RuntimeError``.
"""

try:
    import win32com.client  # type: ignore[import-untyped]

    OUTLOOK_AVAILABLE = True
except ImportError:
    OUTLOOK_AVAILABLE = False


def send_email(
    to: str,
    subject: str,
    html_body: str,
    outlook_app: object | None = None,
) -> None:
    """Send an HTML email through Outlook.

    Parameters
    ----------
    to:
        Recipient email address.
    subject:
        Email subject line.
    html_body:
        Email body as HTML.
    outlook_app:
        An existing ``win32com.client.Dispatch('Outlook.Application')``
        instance.  If *None*, a new one is created.

    Raises
    ------
    RuntimeError
        If Outlook / pywin32 is not available on this platform.
    """
    if not OUTLOOK_AVAILABLE:
        raise RuntimeError(
            "Outlook is not available on this platform. "
            "pywin32 must be installed and Microsoft Outlook must be running."
        )

    if outlook_app is None:
        outlook_app = win32com.client.Dispatch("Outlook.Application")  # type: ignore[possibly-undefined]

    mail = outlook_app.CreateItem(0)  # 0 = olMailItem
    mail.To = to
    mail.Subject = subject
    mail.HTMLBody = html_body
    mail.Send()


def dry_run_email(to: str, subject: str, body: str) -> str:
    """Return a formatted string showing what *would* be sent.

    Works on all platforms — no COM interaction.
    """
    return f"To: {to}\nSubject: {subject}\n{'─' * 40}\n{body}\n{'─' * 40}"
