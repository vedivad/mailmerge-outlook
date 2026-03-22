"""Outlook COM automation for sending emails.

On platforms where ``pywin32`` is not available the module still imports
successfully — ``OUTLOOK_AVAILABLE`` will be ``False`` and ``send_email``
will raise ``RuntimeError``.
"""

import os
import re
from pathlib import Path

try:
    import win32com.client  # type: ignore[import-untyped]

    OUTLOOK_AVAILABLE = True
except ImportError:
    OUTLOOK_AVAILABLE = False


def list_signatures() -> list[str]:
    """Return the names of Outlook signatures installed on this machine.

    Signatures are stored as files in ``%APPDATA%/Microsoft/Signatures/``.
    Returns an empty list on non-Windows platforms.
    """
    appdata = os.environ.get("APPDATA", "")
    if not appdata:
        return []
    sig_dir = Path(appdata) / "Microsoft" / "Signatures"
    if not sig_dir.is_dir():
        return []
    return sorted({p.stem for p in sig_dir.glob("*.htm")})


def _load_signature_html(name: str) -> str:
    """Read the HTML content of a named Outlook signature."""
    appdata = os.environ.get("APPDATA", "")
    if not appdata:
        return ""
    path = Path(appdata) / "Microsoft" / "Signatures" / f"{name}.htm"
    if not path.exists():
        return ""
    return path.read_text(encoding="utf-8", errors="replace")


def send_email(
    to: str,
    subject: str,
    html_body: str,
    outlook_app: object | None = None,
    image_paths: list | None = None,
    signature: str | None = None,
) -> None:
    """Send an HTML email through Outlook.

    Parameters
    ----------
    to:
        Recipient email address.
    subject:
        Email subject line.
    html_body:
        Email body as HTML. Images should use ``cid:filename`` references.
    outlook_app:
        An existing ``win32com.client.Dispatch('Outlook.Application')``
        instance.  If *None*, a new one is created.
    image_paths:
        List of ``pathlib.Path`` objects for images to embed. Each image is
        attached with a Content-ID matching its filename.
    signature:
        Name of an Outlook signature to use. If *None*, the default signature
        is used (populated by Outlook via GetInspector). If set, the named
        signature's HTML is read from disk and appended.

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

    if signature is not None:
        # Use the explicitly chosen signature
        sig_html = _load_signature_html(signature)
        mail.HTMLBody = html_body + sig_html
    else:
        # Use the default signature — GetInspector triggers Outlook to populate it
        _ = mail.GetInspector  # noqa: B018
        signature_html = mail.HTMLBody or ""

        # Insert our content before the signature. Outlook wraps the signature in
        # a full HTML document; we inject our body right after the <body> tag.
        if "<body" in signature_html.lower():
            mail.HTMLBody = re.sub(
                r"(<body[^>]*>)",
                rf"\1{html_body}",
                signature_html,
                count=1,
                flags=re.IGNORECASE,
            )
        else:
            mail.HTMLBody = html_body + signature_html

    mail.To = to
    mail.Subject = subject

    # Embed images as hidden attachments with Content-ID references
    if image_paths:
        for img_path in image_paths:
            attachment = mail.Attachments.Add(str(img_path))
            attachment.PropertyAccessor.SetProperty(
                "http://schemas.microsoft.com/mapi/proptag/0x3712001F",
                img_path.name,
            )

    mail.Send()


def dry_run_email(to: str, subject: str, body: str) -> str:
    """Return a formatted string showing what *would* be sent.

    Works on all platforms — no COM interaction.
    """
    return f"To: {to}\nSubject: {subject}\n{'─' * 40}\n{body}\n{'─' * 40}"
