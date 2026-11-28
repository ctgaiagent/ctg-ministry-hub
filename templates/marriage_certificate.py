"""
templates/marriage_certificate.py — Official CTG Certificate of Marriage PDF Generator
======================================================================================
- Church-inspired Navy (#0F2C59), White (#FFFFFF), and Gold (#D4AF37) theme.
- Supports A4 Portrait and A4 Landscape layouts.
- Includes complete Municipal Form No. 97 / Church of the Temple of God sacramental
  marriage fields: Husband & Wife details, Parents, Place & Date of Marriage,
  Marriage License No., Witnesses/Sponsors, Officiating Priest/Minister, Registry/MC No.,
  Book & Page Number, Certificate Number, and Official Seal.
"""

import importlib
import io
import os
from typing import Any, Dict

NAVY_HEX = "#0F2C59"
GOLD_HEX = "#D4AF37"
SLATE_HEX = "#334155"


def _render_marriage_via_pillow_pdf(
    record: Dict[str, Any],
    certificate_id: str,
    issuance_date: str,
    is_land: bool,
    custom_heading: str,
    custom_wording: str,
    blank_marker: str,
) -> bytes:
    """Fallback A4 PDF generator using Pillow (PIL) when ReportLab is unavailable."""
    from PIL import Image, ImageDraw

    w, h = (1754, 1240) if is_land else (1240, 1754)
    im = Image.new("RGB", (w, h), (255, 255, 255))
    dr = ImageDraw.Draw(im)

    navy_rgb = (15, 44, 89)
    gold_rgb = (212, 175, 55)
    slate_rgb = (51, 65, 85)

    dr.rectangle([45, 45, w - 45, h - 45], outline=navy_rgb, width=6)
    dr.rectangle([62, 62, w - 62, h - 62], outline=gold_rgb, width=3)

    dr.text((w // 2 - 190, 95), "CHURCH OF THE TEMPLE OF GOD", fill=navy_rgb)
    parish_val = str(record.get("Parish") or "").strip() or blank_marker
    dr.text((w // 2 - 140, 130), f"Parish: {parish_val}", fill=slate_rgb)
    dr.text((w // 2 - 180, 175), custom_heading.upper(), fill=gold_rgb)
    dr.text((110, 220), custom_wording[:120], fill=slate_rgb)

    def _val(key: str) -> str:
        v = str(record.get(key) or "").strip()
        return v if v else blank_marker

    fields = [
        ("Husband (Groom):", _val("Husband Name")),
        ("Husband Age / Status:", f"{_val('Husband Age / Birthdate')} / {_val('Husband Single / Widowed')}"),
        ("Husband Residence:", _val("Husband Residence")),
        ("Husband Parents:", f"Father: {_val('Husband Father')} | Mother: {_val('Husband Mother')}"),
        ("Wife (Bride):", _val("Wife Name")),
        ("Wife Age / Status:", f"{_val('Wife Age / Birthdate')} / {_val('Wife Single / Widowed')}"),
        ("Wife Residence:", _val("Wife Residence")),
        ("Wife Parents:", f"Father: {_val('Wife Father')} | Mother: {_val('Wife Mother')}"),
        ("Date of Marriage:", _val("Date of Marriage")),
        ("Place of Marriage:", _val("Place of Marriage")),
        ("Marriage License No.:", f"{_val('Marriage License No.')} ({_val('License Issued At / Date')})"),
        ("Officiating Minister:", _val("Officiating Priest / Minister")),
        ("Witnesses / Sponsors:", _val("Witnesses / Sponsors")),
        ("MC / Book / Page:", f"MC: {_val('MC Number')} | Book: {_val('Book Number')} | Page: {_val('Page Number')}"),
        ("Certificate No.:", f"{certificate_id} (Issued: {issuance_date})"),
    ]

    y = 275
    step = 48 if is_land else 68
    for lbl, val in fields:
        dr.text((115, y), f"{lbl}  {val}", fill=navy_rgb)
        y += step

    dr.line([(120, h - 130), (450, h - 130)], fill=navy_rgb, width=2)
    dr.text((160, h - 115), "Officiating Minister / Parish Priest", fill=navy_rgb)
    dr.line([(w - 450, h - 130), (w - 120, h - 130)], fill=navy_rgb, width=2)
    dr.text((w - 400, h - 115), "Master Admin / Church Registrar", fill=navy_rgb)

    buf = io.BytesIO()
    im.save(buf, format="PDF", resolution=150.0)
    return buf.getvalue()


def render_marriage_certificate_pdf(
    record: Dict[str, Any],
    certificate_id: str,
    issuance_date: str,
    orientation: str = "Portrait",
    logo_path: str = "",
    custom_heading: str = "CERTIFICATE OF MARRIAGE",
    custom_wording: str = (
        "This is to certify that the following contracting parties were solemnly united in the "
        "Holy Bonds of Matrimony according to the rites of the Church of the Temple of God "
        "and in due form of law, as recorded in the Marriage Register of this Church:"
    ),
    blank_marker: str = "________________________",
) -> bytes:
    """
    Renders a printable official CTG Certificate of Marriage PDF in Portrait or Landscape.
    """
    is_land = str(orientation or "Portrait").strip().lower() == "landscape"

    try:
        rl_colors = importlib.import_module("reportlab.lib.colors")
        rl_pagesizes = importlib.import_module("reportlab.lib.pagesizes")
        rl_canvas = importlib.import_module("reportlab.pdfgen.canvas")
    except Exception:
        return _render_marriage_via_pillow_pdf(
            record=record,
            certificate_id=certificate_id,
            issuance_date=issuance_date,
            is_land=is_land,
            custom_heading=custom_heading,
            custom_wording=custom_wording,
            blank_marker=blank_marker,
        )

    A4 = rl_pagesizes.A4
    landscape = rl_pagesizes.landscape
    HexColor = rl_colors.HexColor

    page_size = landscape(A4) if is_land else A4
    width, height = page_size

    buf = io.BytesIO()
    c = rl_canvas.Canvas(buf, pagesize=page_size)
    c.setTitle(f"CTG Marriage Certificate - {certificate_id}")

    navy = HexColor(NAVY_HEX)
    gold = HexColor(GOLD_HEX)
    slate = HexColor(SLATE_HEX)

    c.setFillColor(HexColor("#FFFFFF"))
    c.rect(0, 0, width, height, stroke=0, fill=1)

    c.setStrokeColor(navy)
    c.setLineWidth(3.0)
    c.rect(24, 24, width - 48, height - 48, stroke=1, fill=0)

    c.setStrokeColor(gold)
    c.setLineWidth(1.4)
    c.rect(31, 31, width - 62, height - 62, stroke=1, fill=0)

    c.setFillColor(gold)
    for cx, cy in [(31, 31), (width - 31, 31), (31, height - 31), (width - 31, height - 31)]:
        c.circle(cx, cy, 4.0, stroke=0, fill=1)

    top_y = height - 62
    if logo_path and os.path.exists(logo_path):
        try:
            c.drawImage(logo_path, 50, top_y - 28, width=48, height=48, preserveAspectRatio=True, mask="auto")
        except Exception:
            pass
    else:
        c.setStrokeColor(gold)
        c.setLineWidth(1.5)
        c.circle(74, top_y - 4, 21, stroke=1, fill=0)
        c.setFillColor(navy)
        c.setFont("Times-Bold", 9)
        c.drawCentredString(74, top_y - 7, "CTG")

    c.setFillColor(navy)
    c.setFont("Times-Bold", 19 if not is_land else 17)
    c.drawCentredString(width / 2.0, top_y, "CHURCH OF THE TEMPLE OF GOD")

    parish_label = str(record.get("Parish") or "").strip() or "Central / Parish Registry"
    c.setFont("Times-Roman", 11)
    c.setFillColor(slate)
    c.drawCentredString(width / 2.0, top_y - 16, f"Sacramental Marriage Registry — {parish_label}")

    c.setStrokeColor(gold)
    c.setLineWidth(1.2)
    c.line(70, top_y - 26, width - 70, top_y - 26)

    c.setFillColor(navy)
    c.setFont("Times-Bold", 21 if not is_land else 18)
    c.drawCentredString(width / 2.0, top_y - 50, custom_heading.upper())

    c.setFillColor(slate)
    c.setFont("Times-Italic", 10.5 if not is_land else 9.8)
    words = custom_wording.split()
    lines = []
    cur = []
    max_chars = 95 if is_land else 76
    for w in words:
        if sum(len(x) + 1 for x in cur) + len(w) > max_chars:
            lines.append(" ".join(cur))
            cur = [w]
        else:
            cur.append(w)
    if cur:
        lines.append(" ".join(cur))

    wy = top_y - 72
    for ln in lines[:3]:
        c.drawCentredString(width / 2.0, wy, ln)
        wy -= 13

    def _val(key: str) -> str:
        raw = str(record.get(key) or "").strip()
        return raw if raw else blank_marker

    # Couple Highlight Banner
    banner_y = wy - 18
    c.setStrokeColor(gold)
    c.setFillColor(HexColor("#F8FAFC"))
    c.roundRect(58, banner_y - 24, width - 116, 32, 6, stroke=1, fill=1)
    c.setFillColor(navy)
    c.setFont("Times-Bold", 13.5 if not is_land else 12.5)
    couple_str = f"{_val('Husband Name')}   &   {_val('Wife Name')}"
    c.drawCentredString(width / 2.0, banner_y - 13, couple_str[:85])

    fields_start_y = banner_y - 46
    row_gap = 20 if is_land else 27
    left_label_x = 60
    left_val_x = 210

    field_rows = [
        ("Husband (Groom):", f"{_val('Husband Name')}  (Age/Birth: {_val('Husband Age / Birthdate')} | Status: {_val('Husband Single / Widowed')})"),
        ("Husband Residence & Nationality:", f"{_val('Husband Residence')}  ({_val('Husband Nationality')})"),
        ("Husband's Parents:", f"Father: {_val('Husband Father')}   |   Mother: {_val('Husband Mother')}"),
        ("Wife (Bride):", f"{_val('Wife Name')}  (Age/Birth: {_val('Wife Age / Birthdate')} | Status: {_val('Wife Single / Widowed')})"),
        ("Wife Residence & Nationality:", f"{_val('Wife Residence')}  ({_val('Wife Nationality')})"),
        ("Wife's Parents:", f"Father: {_val('Wife Father')}   |   Mother: {_val('Wife Mother')}"),
        ("Date of Marriage:", _val("Date of Marriage")),
        ("Place of Marriage:", f"{_val('Place of Marriage')} ({_val('Parish')})"),
        ("Marriage License No.:", f"{_val('Marriage License No.')}   |   Issued: {_val('License Issued At / Date')}"),
        ("Officiating Priest / Minister:", f"{_val('Officiating Priest / Minister')} ({_val('Priest Position / Religion')})"),
        ("Witnesses / Sponsors:", _val("Witnesses / Sponsors")),
    ]

    cy = fields_start_y
    for label, value in field_rows:
        c.setFillColor(navy)
        c.setFont("Times-Bold", 10.2 if not is_land else 9.5)
        c.drawString(left_label_x, cy, label)

        c.setFillColor(HexColor("#111827"))
        c.setFont("Times-Roman", 10.5 if not is_land else 9.8)
        c.drawString(left_val_x, cy, value[:88])

        c.setStrokeColor(HexColor("#CBD5E1"))
        c.setLineWidth(0.5)
        c.line(left_val_x, cy - 3, width - 60, cy - 3)
        cy -= row_gap

    reg_y = cy - 6
    c.setFillColor(navy)
    c.setFont("Times-Bold", 9.8)
    c.drawString(left_label_x, reg_y, f"MC No.: {_val('MC Number')}")
    c.drawString(left_label_x + 135, reg_y, f"Book No.: {_val('Book Number')}")
    c.drawString(left_label_x + 265, reg_y, f"Page No.: {_val('Page Number')}")
    c.drawString(left_label_x + 380, reg_y, f"Record ID: {_val('Record ID')}")

    meta_y = reg_y - 18
    c.setFont("Times-Italic", 9.3)
    c.setFillColor(slate)
    c.drawString(left_label_x, meta_y, f"Certificate No.: {certificate_id}")
    c.drawString(left_label_x + 265, meta_y, f"Date Issued: {issuance_date}")

    sig_y = 74
    c.setStrokeColor(navy)
    c.setLineWidth(1.0)
    c.line(65, sig_y + 18, 255, sig_y + 18)
    c.setFillColor(navy)
    c.setFont("Times-Bold", 9.8)
    c.drawCentredString(160, sig_y + 4, "Officiating Minister / Parish Priest")
    c.setFont("Times-Roman", 8.8)
    c.drawCentredString(160, sig_y - 8, _val("Officiating Priest / Minister")[:38])

    seal_cx = width / 2.0
    seal_cy = sig_y + 14
    c.setStrokeColor(gold)
    c.setLineWidth(1.5)
    c.circle(seal_cx, seal_cy, 24, stroke=1, fill=0)
    c.circle(seal_cx, seal_cy, 20, stroke=1, fill=0)
    c.setFillColor(gold)
    c.setFont("Times-Bold", 7.2)
    c.drawCentredString(seal_cx, seal_cy + 3, "OFFICIAL")
    c.drawCentredString(seal_cx, seal_cy - 6, "CHURCH SEAL")

    right_sig_x1 = width - 255
    right_sig_x2 = width - 65
    right_center = (right_sig_x1 + right_sig_x2) / 2.0
    c.setStrokeColor(navy)
    c.setLineWidth(1.0)
    c.line(right_sig_x1, sig_y + 18, right_sig_x2, sig_y + 18)
    c.setFillColor(navy)
    c.setFont("Times-Bold", 9.8)
    c.drawCentredString(right_center, sig_y + 4, "Master Admin / Church Registrar")
    c.setFont("Times-Roman", 8.8)
    c.drawCentredString(right_center, sig_y - 8, "Authorized Signatory")

    c.showPage()
    c.save()
    return buf.getvalue()
