"""
baptismal/validation.py — Input Sanitization, XSS Protection & Upload Security
=============================================================================
Protects against:
- HTML / XSS injection in Streamlit UI rendering
- Excel Formula Injection (CWE-1236)
- Path traversal in filenames
- MIME spoofing, oversized uploads, malicious PDF JavaScript, and ZIP bombs
"""

import html
import io
import os
import re
import secrets
import zipfile
from typing import Iterable, Tuple

from .config import ALLOWED_UPLOAD_DOC_EXTS, ALLOWED_UPLOAD_IMAGE_EXTS, MAX_UPLOAD_SIZE_MB


def escape_html_text(value: object) -> str:
    """Escapes user-controlled strings before any HTML rendering."""
    if value is None:
        return ""
    return html.escape(str(value), quote=True)


def sanitize_excel_cell(value: object) -> str:
    """Neutralizes spreadsheet formula injection (=, +, -, @, tab, CR) while preserving dates and text."""
    if value is None:
        return ""
    s = str(value).strip()
    if not s:
        return ""
    s = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", s)
    if s[0] in ("=", "+", "@", "\t", "\r"):
        if not re.match(r"^\+?[0-9\s\-()]{7,20}$", s):
            return "'" + s
    elif s.startswith("-") and len(s) > 1 and not re.match(r"^-[0-9.,]+$", s):
        return "'" + s
    return s


def sanitize_filename(original_name: str, preserve_original: bool = True) -> str:
    """Strips directory components and dangerous characters to prevent path traversal."""
    base = os.path.basename(str(original_name or "upload.bin")).replace("\x00", "")
    base = re.sub(r"[\\/:*?\"<>|]+", "_", base).strip(" .")
    if not base or base in (".", ".."):
        base = f"baptismal_file_{secrets.token_hex(4)}.bin"
    if not preserve_original:
        stem, ext = os.path.splitext(base)
        safe_stem = re.sub(r"[^A-Za-z0-9._-]+", "_", stem)[:60] or "file"
        return f"{safe_stem}_{secrets.token_hex(4)}{ext.lower()}"
    return base[:180]


def validate_uploaded_file(
    file_bytes: bytes,
    filename: str,
    declared_mime: str = "",
    allowed_exts: Iterable[str] | None = None,
    max_mb: int = MAX_UPLOAD_SIZE_MB,
) -> Tuple[bool, str, str]:
    """Validates file size, extension, path-traversal safety, magic-byte MIME signature,
    embedded PDF scripts, and ZIP-bomb bounds."""
    if not file_bytes:
        return False, "Empty file uploaded.", ""

    size_limit = int(max_mb) * 1024 * 1024
    if len(file_bytes) > size_limit:
        return False, f"File exceeds maximum allowed size of {max_mb} MB.", ""

    raw_name = str(filename or "")
    if ".." in raw_name or "/" in raw_name or "\\" in raw_name or "\x00" in raw_name:
        return False, "Invalid filename: path traversal characters detected.", ""

    safe_name = sanitize_filename(raw_name, preserve_original=True)
    ext = safe_name.rsplit(".", 1)[-1].lower() if "." in safe_name else ""
    valid_exts = set(allowed_exts) if allowed_exts is not None else (ALLOWED_UPLOAD_IMAGE_EXTS | ALLOWED_UPLOAD_DOC_EXTS)
    if ext not in valid_exts:
        return False, f"Unsupported file extension '.{ext}'. Allowed: {sorted(valid_exts)}", ""

    head = file_bytes[:32]
    if ext in ("jpg", "jpeg"):
        if not head.startswith(b"\xff\xd8\xff"):
            return False, "File content does not match JPEG signature.", ""
    elif ext == "png":
        if not head.startswith(b"\x89PNG\r\n\x1a\n"):
            return False, "File content does not match PNG signature.", ""
    elif ext == "webp":
        if not (head.startswith(b"RIFF") and b"WEBP" in head[:16]):
            return False, "File content does not match WEBP signature.", ""
    elif ext in ("tif", "tiff"):
        if not (head.startswith(b"II*\x00") or head.startswith(b"MM\x00*")):
            return False, "File content does not match TIFF signature.", ""
    elif ext == "pdf":
        if not head.startswith(b"%PDF-"):
            return False, "File content does not match PDF signature.", ""
        upper_scan = file_bytes[:262144].upper()
        if b"/JAVASCRIPT" in upper_scan or b"/JS " in upper_scan or b"/LAUNCH" in upper_scan:
            return False, "Blocked PDF containing embedded JavaScript or Launch actions.", ""
    elif ext == "xlsx":
        if not head.startswith(b"PK\x03\x04"):
            return False, "File content does not match XLSX/ZIP signature.", ""
        try:
            with zipfile.ZipFile(io.BytesIO(file_bytes), "r") as zf:
                total_uncompressed = sum(zinfo.file_size for zinfo in zf.infolist())
                if total_uncompressed > 150 * 1024 * 1024:
                    return False, "Blocked XLSX file exceeding safe uncompressed size limit.", ""
        except Exception:
            return False, "Corrupted or unreadable XLSX archive.", ""

    return True, "OK", safe_name
