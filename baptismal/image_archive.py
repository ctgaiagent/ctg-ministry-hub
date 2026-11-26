"""
baptismal/image_archive.py — BOOK 1–9 Visual Archive & Image Manipulation
=========================================================================
Supports image thumbnails, full-size preview, rotation (0/90/180/270),
zoom, original filename & book folder display, and multi-image / multi-record
associations without auto-trusting filename page numbers.
"""

import io
import os
from typing import Any, Dict, List, Optional
from PIL import Image

from .database import DatabaseManager


def render_image_preview_bytes(
    image_bytes: bytes,
    rotation_degrees: int = 0,
    max_width: Optional[int] = None,
) -> bytes:
    """Renders a rotated or thumbnail-resized JPEG preview from raw image bytes."""
    im = Image.open(io.BytesIO(image_bytes)).convert("RGB")
    rot = int(rotation_degrees) % 360
    if rot:
        im = im.rotate(-rot, expand=True)
    if max_width and im.width > max_width:
        scale = float(max_width) / float(im.width)
        im = im.resize((int(max_width), max(1, int(im.height * scale))), Image.Resampling.LANCZOS)
    buf = io.BytesIO()
    im.save(buf, format="JPEG", quality=88)
    return buf.getvalue()


class ImageArchiveService:
    """Manages BOOK 1–9 scanned archive listings and record-to-image links."""

    def __init__(self, db: DatabaseManager):
        self.db = db

    def list_associations_for_record(self, record_id: str) -> List[Dict[str, Any]]:
        return self.db.fetch_all(
            """
            SELECT association_id, record_id, book_folder, image_filename, drive_file_id,
                   extracted_page_hint, record_page_number, match_status, confirmed_by,
                   confirmed_at, notes, created_at
            FROM historical_image_associations
            WHERE record_id = %s
            ORDER BY created_at DESC
            """,
            (record_id,),
        )

    def list_records_for_image(self, book_folder: str, image_filename: str) -> List[Dict[str, Any]]:
        """Supports multiple historical records linked to a single scanned page."""
        return self.db.fetch_all(
            """
            SELECT a.*, h.child_name, h.date_of_baptism, h.book_number, h.page_number
            FROM historical_image_associations a
            LEFT JOIN historical_records h ON a.record_id = h.record_id
            WHERE a.book_folder = %s AND a.image_filename = %s
            ORDER BY a.created_at DESC
            """,
            (book_folder, image_filename),
        )

    def list_by_book(self, book_folder: str, status_filter: str = "ALL") -> List[Dict[str, Any]]:
        if status_filter and status_filter != "ALL":
            return self.db.fetch_all(
                """
                SELECT * FROM historical_image_associations
                WHERE book_folder = %s AND match_status = %s
                ORDER BY image_filename ASC
                """,
                (book_folder, status_filter),
            )
        return self.db.fetch_all(
            """
            SELECT * FROM historical_image_associations
            WHERE book_folder = %s
            ORDER BY image_filename ASC
            """,
            (book_folder,),
        )

    @staticmethod
    def scan_local_book_folder(base_dir: str, book_folder: str) -> List[str]:
        target = os.path.join(base_dir, book_folder)
        if not os.path.isdir(target):
            return []
        exts = (".jpg", ".jpeg", ".png", ".webp", ".tif", ".tiff")
        return sorted([f for f in os.listdir(target) if f.lower().endswith(exts)])
