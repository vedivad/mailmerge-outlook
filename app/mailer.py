"""Outlook COM automation for sending emails.

On platforms where ``pywin32`` is not available the module still imports
successfully — ``OUTLOOK_AVAILABLE`` will be ``False`` and ``send_email``
will raise ``RuntimeError``.
"""

import logging
import os
import re
from pathlib import Path
from urllib.parse import unquote

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


def _signatures_dir() -> Path | None:
    """Return the Outlook Signatures directory, or None if unavailable."""
    appdata = os.environ.get("APPDATA", "")
    if not appdata:
        return None
    sig_dir = Path(appdata) / "Microsoft" / "Signatures"
    if not sig_dir.is_dir():
        return None
    return sig_dir


def list_signatures() -> list[str]:
    """Return the names of Outlook signatures installed on this machine.

    Signatures are stored as files in ``%APPDATA%/Microsoft/Signatures/``.
    Returns an empty list on non-Windows platforms.
    """
    sig_dir = _signatures_dir()
    if sig_dir is None:
        return []
    return sorted({p.stem for p in sig_dir.glob("*.htm")})


def _load_signature(name: str) -> tuple[str, list[Path]]:
    """Load a signature's HTML body content and its image files.

    Returns a tuple of (sig_body_html, image_paths). The sig_body_html has
    relative image paths replaced with ``cid:`` references. Image paths point
    to the actual files in the signature's support folder.
    """
    sig_dir = _signatures_dir()
    if sig_dir is None:
        return "", []

    htm_path = sig_dir / f"{name}.htm"
    if not htm_path.exists():
        return "", []

    # Detect encoding from the file — Outlook signatures typically use
    # windows-1252, not UTF-8. Read as bytes first, sniff the charset.
    raw_bytes = htm_path.read_bytes()
    charset_match = re.search(
        rb'charset=(["\']?)([^"\';>\s]+)\1', raw_bytes, re.IGNORECASE
    )
    encoding = charset_match.group(2).decode("ascii") if charset_match else "utf-8"
    raw_html = raw_bytes.decode(encoding, errors="replace")

    # Extract just the <body> content
    body_match = re.search(
        r"<body[^>]*>(.*)</body>", raw_html, re.IGNORECASE | re.DOTALL
    )
    if not body_match:
        return "", []

    sig_body = body_match.group(1)

    image_files: list[Path] = []
    seen_files: set[str] = set()

    def _replace_src(match: re.Match) -> str:
        attr = match.group(1)  # "src" or "src"
        quote = match.group(2)  # quote character
        rel_path = match.group(3)  # the relative path

        # Skip if already a cid:, http:, or https: reference
        if rel_path.startswith(("cid:", "http:", "https:")):
            return match.group(0)

        decoded = unquote(rel_path)
        img_path = sig_dir / decoded

        if img_path.exists() and img_path.name not in seen_files:
            seen_files.add(img_path.name)
            image_files.append(img_path)

        # Replace with cid: reference
        return f"{attr}={quote}cid:{img_path.name}{quote}"

    # Replace src="..." in both <img> and VML <v:imagedata>
    sig_body = re.sub(
        r'(src)=(["\'])([^"\']+)\2',
        _replace_src,
        sig_body,
        flags=re.IGNORECASE,
    )

    # Also handle o:href="cid:..." in VML imagedata (leave as-is if already cid)
    sig_body = re.sub(
        r'(o:href)=(["\'])([^"\']+)\2',
        _replace_src,
        sig_body,
        flags=re.IGNORECASE,
    )

    # Strip VML conditional comments — they duplicate the <img> tags and
    # contain complex markup that can confuse Outlook when re-injected.
    sig_body = re.sub(
        r"<!--\[if gte vml 1\]>.*?<!\[endif\]-->",
        "",
        sig_body,
        flags=re.IGNORECASE | re.DOTALL,
    )

    # Encode non-ASCII to HTML entities (same as template body)
    sig_body = sig_body.encode("ascii", "xmlcharrefreplace").decode("ascii")

    return sig_body, image_files


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
        Name of an Outlook signature to append. If *None*, no signature is
        added. The signature is loaded from disk and its images are embedded
        as CID attachments.

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

    # Load signature from disk (if requested)
    sig_body = ""
    sig_images: list[Path] = []
    if signature:
        sig_body, sig_images = _load_signature(signature)
        _debug(f"Signature loaded: {len(sig_body)} chars, {len(sig_images)} images")

    # Encode non-ASCII characters as HTML entities so Outlook's internal
    # encoding (windows-1252) doesn't corrupt umlauts and other characters.
    safe_body = html_body.encode("ascii", "xmlcharrefreplace").decode("ascii")

    meta = '<meta http-equiv="Content-Type" content="text/html; charset=utf-8">'
    # Reset paragraph margins for Outlook's MsoNormal class used in signatures
    sig_style = "<style>p.MsoNormal, li.MsoNormal, div.MsoNormal { margin: 0; }</style>"
    mail.HTMLBody = (
        f"<html><head>{meta}{sig_style}</head><body>{safe_body}{sig_body}</body></html>"
    )
    _debug("HTMLBody OK")

    # Embed template images as hidden attachments with Content-ID references
    all_images = list(image_paths or []) + sig_images
    for img_path in all_images:
        _debug(f"Attaching {img_path}")
        attachment = mail.Attachments.Add(str(img_path))
        attachment.PropertyAccessor.SetProperty(
            "http://schemas.microsoft.com/mapi/proptag/0x3712001F",
            img_path.name,
        )
        _debug(f"Attached {img_path.name} OK")

    _debug("Calling Send")
    mail.Send()
    _debug("Send OK")


def dry_run_email(to: str, subject: str, body: str) -> str:
    """Return a formatted string showing what *would* be sent.

    Works on all platforms — no COM interaction.
    """
    return f"To: {to}\nSubject: {subject}\n{'─' * 40}\n{body}\n{'─' * 40}"
