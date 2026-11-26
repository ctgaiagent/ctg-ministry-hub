"""
templates/baptismal_certificate.py — Professional CTG Certificate of Baptism PDF Generator
==========================================================================================
- Church-inspired Navy (#0F2C59), White (#FFFFFF), and Gold (#D4AF37) theme.
- Supports A4 Portrait and A4 Landscape layouts.
- Double border, church logo area, certificate heading, typography, signature blocks,
  official seal area, certificate number, and issuance date.
- NEVER invents missing historical information; leaves missing fields blank or marked
  for administrator completion without altering the source record.
"""

import io
import os
from typing import Any, Dict, Optional

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape, portrait
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas

NAVY_HEX = "#0F2C59"
GOLD_HEX = "#D4AF37"
SLATE_HEX = "#334155"


def render_baptismal_certificate_pdf(
    record: Dict[str, Any],
    certificate_id: str,
    issuance_date: str,
    orientation: str = "portrait",
    custom_heading: str = "CERTIFICATE OF BAPTISM",
    custom_wording: str = "This is to certify that the person named below received the Sacrament of Christian Baptism according to the rites of the Church of the Temple of God.",
    logo_path: str = "logo.pg",
    blank_marker: str = "____________________________",
) -> bytes:
    """Generates a print-ready A4 PDF Certificate of Baptism in Portrait or Landscape."""
    is_land = str(orientation).strip().lower() == "landscape"
    page_size = landscape(A4) if is_land else portrait(A4)
    width, height = page_size

    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=page_size)
    c.setTitle(f"CTG Certificate of Baptism - {certificate_id}")

    navy = colors.HexColor(NAVY_HEX)
    gold = colors.HexColor(GOLD_HEX)
    slate = colors.HexColor(SLATE_HEX)

    # Outer & inner ceremonial borders
    margin = 14 * mm
    c.setStrokeColor(navy)
    c.setLineWidth(3.0)
    c.rect(margin, margin, width - 2 * margin, height - 2 * margin)

    c.setStrokeColor(gold)
    c.setLineWidth(1.4)
    c.rect(margin + 3.5 * mm, margin + 3.5 * mm, width - 2 * (margin + 3.5 * mm), height - 2 * (margin + 3.5 * mm))

    # Header & optional logo area
    top_y = height - margin - 14 * mm
    if logo_path and os.path.exists(logo_path):
        try:
            c.drawImage(logo_path, width / 2 - 11 * mm, top_y - 16 * mm, width=22 * mm, height=22 * mm, preserveAspectRatio=True, mask="auto")
            top_y -= 20 * mm
        except Exception:
            pass

    c.setFillColor(navy)
    c.setFont("Times-Bold", 15)
    c.drawCentredString(width / 2, top_y, "CHURCH OF THE TEMPLE OF GOD")
    c.setFont("Times-Roman", 10)
    c.setFillColor(slate)
    parish_val = str(record.get("Parish") or "").strip() or blank_marker
    c.drawCentredString(width / 2, top_y - 6 * mm, f"Parish: {parish_val}")

    c.setFillColor(gold)
    c.setFont("Times-Bold", 24 if is_land else 22)
    c.drawCentredString(width / 2, top_y - 18 * mm, custom_heading.upper())

    c.setStrokeColor(gold)
    c.setLineWidth(1.0)
    c.line(width * 0.22, top_y - 21 * mm, width * 0.78, top_y - 21 * mm)

    c.setFillColor(slate)
    c.setFont("Times-Italic", 10.5)
    c.drawCentredString(width / 2, top_y - 29 * mm, custom_wording[:115])

    def _val(key: str) -> str:
        v = str(record.get(key) or "").strip()
        return v if v else blank_marker

    # Certificate Body Fields (never inventing missing historical info)
    fields_left = [
        ("Name of Baptized Person:", _val("Name of the Child")),
        ("Date of Baptism:", _val("Date of Baptism")),
        ("Date of Birth:", _val("Date of Birth")),
        ("Place of Baptism:", _val("Place of Baptism")),
        ("Name of Father:", _val("Name of the Father")),
        ("Name of Mother:", _val("Name of the Mother")),
    ]
    fields_right = [
        ("Sponsors:", _val("Sponsors")),
        ("Officiating Priest:", _val("Officiating Priest")),
        ("Book Number:", _val("Book Number")),
        ("Page Number:", _val("Page Number")),
        ("Certificate No.:", certificate_id),
        ("Date Issued:", issuance_date),
    ]

    body_top = top_y - 42 * mm
    row_step = 10.5 * mm if is_land else 12.0 * mm

    if is_land:
        col1_x = margin + 14 * mm
        col2_x = width / 2 + 6 * mm
        for idx, (lbl, val) in enumerate(fields_left):
            y_pos = body_top - idx * row_step
            c.setFillColor(navy)
            c.setFont("Times-Bold", 10.5)
            c.drawString(col1_x, y_pos, lbl)
            c.setFillColor(colors.black)
            c.setFont("Times-Roman", 11)
            c.drawString(col1_x + 48 * mm, y_pos, val[:45])
        for idx, (lbl, val) in enumerate(fields_right):
            y_pos = body_top - idx * row_step
            c.setFillColor(navy)
            c.setFont("Times-Bold", 10.5)
            c.drawString(col2_x, y_pos, lbl)
            c.setFillColor(colors.black)
            c.setFont("Times-Roman", 11)
            c.drawString(col2_x + 38 * mm, y_pos, val[:45])
    else:
        col_x = margin + 14 * mm
        all_fields = fields_left + fields_right
        for idx, (lbl, val) in enumerate(all_fields):
            y_pos = body_top - idx * row_step
            c.setFillColor(navy)
            c.setFont("Times-Bold", 10.5)
            c.drawString(col_x, y_pos, lbl)
            c.setFillColor(colors.black)
            c.setFont("Times-Roman", 11)
            c.drawString(col_x + 52 * mm, y_pos, val[:58])

    # Signature & Official Seal Area
    sig_y = margin + 28 * mm
    c.setStrokeColor(navy)
    c.setLineWidth(0.8)
    c.line(margin + 16 * mm, sig_y, margin + 74 * mm, sig_y)
    c.line(width - margin - 74 * mm, sig_y, width - margin - 16 * mm, sig_y)

    c.setFillColor(navy)
    c.setFont("Times-Bold", 9.5)
    c.drawCentredString(margin + 45 * mm, sig_y - 5 * mm, "Parish Priest / Cura Parocco")
    c.drawCentredString(width - margin - 45 * mm, sig_y - 5 * mm, "Master Admin / Archivist")

    # Official Seal Circle
    c.setStrokeColor(gold)
    c.setLineWidth(1.2)
    c.circle(width / 2, sig_y - 1 * mm, 13 * mm)
    c.setFillColor(gold)
    c.setFont("Times-Bold", 8)
    c.drawCentredString(width / 2, sig_y, "OFFICIAL")
    c.drawCentredString(width / 2, sig_y - 4 * mm, "PARISH SEAL")

    c.showPage()
    c.save()
    return buf.getvalue()
