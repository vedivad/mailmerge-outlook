"""Outlook COM automation for sending emails.

On platforms where ``pywin32`` is not available the module still imports
successfully — ``OUTLOOK_AVAILABLE`` will be ``False`` and ``send_email``
will raise ``RuntimeError``.

Signature handling uses the ``GetInspector`` trick: instead of manually
parsing signature HTML files from disk, we let Outlook render the new-mail
template (which includes the user's default signature, fonts, and inline
images), then inject our content before the signature in the generated HTML.
"""

import logging
import re

_log = logging.getLogger(__name__)


def _debug(msg: str) -> None:
    """Log and print a debug message for send troubleshooting."""
    _log.debug(msg)
    print(f"[mailer] {msg}")


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
    image_paths: list | None = None,
    draft: bool = False,
) -> None:
    """Send (or draft) an HTML email through Outlook, preserving the default signature.

    Uses the ``GetInspector`` trick to let Outlook natively populate the
    user's default signature (with images, fonts, and layout intact), then
    injects the template content before the signature.

    Parameters
    ----------
    to:
        Recipient email address.
    subject:
        Email subject line.
    html_body:
        Email body as HTML.  Should already be wrapped in a ``<div>`` with
        inline font styles (as produced by ``template_manager.render_html``).
        Images should use ``cid:filename`` references.
    outlook_app:
        An existing ``win32com.client.Dispatch('Outlook.Application')``
        instance.  If *None*, a new one is created.
    image_paths:
        List of ``pathlib.Path`` objects for template images to embed.
        Each image is attached with a Content-ID matching its filename.
    draft:
        If *True*, save the email to the Drafts folder instead of sending it.

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
    _debug("CreateItem OK")

    mail.To = to
    _debug(f"To={to} OK")
    mail.Subject = subject
    _debug("Subject OK")

    # Force Outlook to render the default signature into mail.HTMLBody.
    # Accessing GetInspector triggers the internal template engine which
    # populates the HTML with the user's default new-mail signature,
    # including any inline images and Word-specific CSS.
    _inspector = mail.GetInspector  # noqa: F841 — side-effect access
    _debug("GetInspector OK — default signature populated")

    existing_body = mail.HTMLBody

    # Encode non-ASCII characters as HTML entities so Outlook's internal
    # encoding (windows-1252) doesn't corrupt umlauts and other characters.
    safe_content = html_body.encode("ascii", "xmlcharrefreplace").decode("ascii")

    # Inject our content right after the opening <body> tag, before the
    # signature that Outlook already placed there.
    if re.search(r"<body[^>]*>", existing_body, re.IGNORECASE):
        new_body = re.sub(
            r"(<body[^>]*>)",
            rf"\1{safe_content}<br>",
            existing_body,
            count=1,
            flags=re.IGNORECASE,
        )
    else:
        # Fallback if no body tag is found (shouldn't happen)
        new_body = safe_content + existing_body

    mail.HTMLBody = new_body
    _debug("HTMLBody OK")

    # Embed template images as hidden attachments with Content-ID references
    for img_path in image_paths or []:
        _debug(f"Attaching {img_path}")
        attachment = mail.Attachments.Add(str(img_path))
        attachment.PropertyAccessor.SetProperty(
            "http://schemas.microsoft.com/mapi/proptag/0x3712001F",
            img_path.name,
        )
        _debug(f"Attached {img_path.name} OK")

    if draft:
        _debug("Saving to Drafts")
        mail.Save()
        _debug("Draft saved OK")
    else:
        _debug("Calling Send")
        mail.Send()
        _debug("Send OK")


def dry_run_email(to: str, subject: str, body: str) -> str:
    """Return a formatted string showing what *would* be sent.

    Works on all platforms — no COM interaction.
    """
    return f"To: {to}\nSubject: {subject}\n{'─' * 40}\n{body}\n{'─' * 40}"
