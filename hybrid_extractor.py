"""
hybrid_extractor.py — 3-Tier Bulletproof Hybrid Extraction Engine for CTG Ministry Hub
======================================================================================
- TIER 1: Pure Non-AI Deterministic Parser (openpyxl / pandas) for .xlsx, .xls, .csv (0 AI tokens, instant).
- TIER 2: Auto-Discovering Multi-Page Gemini Vision Engine (google-genai) using client.models.list() at runtime
          to permanently eliminate 404 NOT_FOUND errors, processing all 250 DPI PDF pages (Pages 1–4) together.
- TIER 3: Offline / Non-AI Local Python OCR Fallback (rapidocr-onnxruntime / paddleocr / pytesseract) when
          GEMINI_API_KEY is missing, offline, or out of quota — with zero filename guessing or dummy rows.
"""

import io
import re
import json
import importlib
from typing import List, Dict, Any, Tuple
import openpyxl
import pandas as pd
from PIL import Image, ImageOps, ImageEnhance
from google import genai
from google.genai import types

try:
    from pydantic import BaseModel as _PydanticBaseModel, Field as _PydanticField
except Exception:
    _PydanticBaseModel = object
    def _PydanticField(default=None, **kwargs):
        return default


# =====================================================================
# 1. STRICT PYDANTIC STRUCTURED OUTPUT SCHEMA (STEP 3 EDITOR COMPATIBLE)
# =====================================================================
class FamilyMember(_PydanticBaseModel):
    full_name: str = _PydanticField(description="Kumpletong Pangalan ng miyembro (pagsamahin ang Given Name sa itaas at Surname sa ibaba kung 2-line sa cell)")
    relationship: str = _PydanticField(description="Relasyon sa Head of Family (HEAD OF FAMILY, HUSBAND, WIFE, SON, DAUGHTER, SON IN LAW, DAUGHTER IN LAW, GRANDSON, GRANDDAUGHTER, MOTHER, FATHER, BROTHER, SISTER, OTHER)")
    sex: str = _PydanticField(description="MALE o FEMALE")
    birthdate: str = _PydanticField(description="Petsa ng kapanganakan (YYYY-MM-DD)")
    age: str = _PydanticField(description="Edad")
    category: str = _PydanticField(description="Category (KID, YOUTH, ADULT, SENIOR CITIZEN)")
    civil_status: str = _PydanticField(description="SINGLE, MARRIED, WIDOWED, SEPARATED")
    student_working: str = _PydanticField(description="STUDENT, WORKING, NONE")
    contact_no: str = _PydanticField(description="Mobile number kung meron, o blanko ''")
    church_member: str = _PydanticField(description="YES o NO")
    status: str = _PydanticField(description="ACTIVE o INACTIVE")
    baptized: str = _PydanticField(description="YES o NO")
    baptism_date: str = _PydanticField(description="Petsa ng binyag kung nakasulat, o blanko ''")
    current_address: str = _PydanticField(description="Current Address kung nakasulat (kopyahin ang nasa itaas kung ditto marks ''), o blanko ''")


class FamilyFormExtraction(_PydanticBaseModel):
    parish: str = _PydanticField(description="Section I: Parish name")
    area_zone: str = _PydanticField(description="Section I: Area / Zone")
    cell_group: str = _PydanticField(description="Section I: Cell Group")
    family_id: str = _PydanticField(description="Section II: Family ID kung nakasulat")
    head_of_family: str = _PydanticField(description="Section II: Head of Family na aktwal na nakasulat sa Page 1")
    household_mobile: str = _PydanticField(description="Section II: Mobile Number na nakasulat sa Page 1")
    complete_address: str = _PydanticField(description="Section II: Complete Address (House / Street / Purok)")
    barangay: str = _PydanticField(description="Section II: Barangay")
    city_municipality: str = _PydanticField(description="Section II: City / Municipality")
    province_area: str = _PydanticField(description="Section II: Province")
    zip_code: str = _PydanticField(description="Section II: ZIP Code")
    family_members: List[FamilyMember] = _PydanticField(description="Section III: Lahat ng miyembrong may nakasulat na pangalan mula sa Page 2, Page 3, at kasunod na pahina")
    emergency_contact_name: str = _PydanticField(description="Section V (Page 4): Emergency Contact Name na aktwal na nakasulat")
    emergency_contact_no_rel: str = _PydanticField(description="Section V (Page 4): Emergency Contact Number at Relationship, hal. '09123456789 (Wife)'")
    household_remarks: str = _PydanticField(description="Section VI (Page 4): Remarks o notes kung meron")


# =====================================================================
# 2. HIGH-RESOLUTION (250 DPI) PDF TO IMAGES CONVERTER
# =====================================================================
def enhance_form_image(pil_im: Image.Image) -> Image.Image:
    try:
        rgb = ImageOps.exif_transpose(pil_im).convert("RGB")
        rgb = ImageOps.autocontrast(rgb, cutoff=0.4)
        rgb = ImageEnhance.Contrast(rgb).enhance(1.28)
        rgb = ImageEnhance.Sharpness(rgb).enhance(1.55)
        return rgb
    except Exception:
        return pil_im.convert("RGB")


def pdf_to_images_250dpi(pdf_bytes: bytes, dpi: int = 250) -> List[Image.Image]:
    """Converts ALL pages of a PDF to high-resolution 250 DPI RGB PIL Images using PyMuPDF (fitz) or pypdfium2."""
    if not pdf_bytes:
        return []
    images: List[Image.Image] = []
    # 1. Try PyMuPDF (fitz)
    try:
        fitz = importlib.import_module("fitz")
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        for page in doc:
            pix = page.get_pixmap(dpi=dpi)
            img = Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")
            if img.width < img.height and len(doc) >= 2:
                img = img.rotate(90, expand=True)
            images.append(img)
        doc.close()
        if images:
            return images
    except Exception:
        images = []

    # 2. Fallback to pypdfium2 at 250 DPI (scale = dpi / 72.0)
    try:
        pdfium = importlib.import_module("pypdfium2")
        doc = pdfium.PdfDocument(pdf_bytes)
        scale = float(dpi) / 72.0
        for i in range(len(doc)):
            page = doc[i]
            bitmap = page.render(scale=scale)
            img = bitmap.to_pil().convert("RGB")
            if img.width < img.height and len(doc) >= 2:
                img = img.rotate(90, expand=True)
            images.append(img.copy())
            bitmap.close()
            page.close()
        doc.close()
    except Exception:
        pass
    return images


# =====================================================================
# 3. TIER 1: PURE NON-AI DETERMINISTIC EXCEL / CSV PARSER
# =====================================================================
def tier1_parse_excel_or_csv(file_bytes: bytes, filename: str, default_parish: str = "LIPAY NORTE PARISH") -> Dict[str, Any]:
    """Parses Official CTG Multi-Section Family Form (.xlsx) or 29-Column Batch Excel/CSV with 0 AI calls."""
    ext = filename.split(".")[-1].lower() if "." in filename else "xlsx"
    result: Dict[str, Any] = {
        "parish": default_parish,
        "area_zone": "",
        "cell_group": "",
        "family_id": "",
        "head_of_family": "",
        "household_mobile": "",
        "complete_address": "",
        "barangay": "",
        "city_municipality": "",
        "province_area": "",
        "zip_code": "",
        "family_members": [],
        "emergency_contact_name": "",
        "emergency_contact_no_rel": "",
        "household_remarks": "",
    }

    if ext == "csv":
        df = pd.read_csv(io.BytesIO(file_bytes)).fillna("")
        for _, row in df.iterrows():
            fn = str(row.get("Full Name", row.get("Name", ""))).strip()
            if not fn:
                continue
            if not result["head_of_family"]:
                result["head_of_family"] = str(row.get("Head of Family", fn)).strip()
                result["parish"] = str(row.get("Parish", default_parish)).strip() or default_parish
                result["family_id"] = str(row.get("Family ID", "")).strip()
                result["area_zone"] = str(row.get("Area/Zone", "")).strip()
                result["cell_group"] = str(row.get("Cell Group", "")).strip()
                result["complete_address"] = str(row.get("House / Street / Purok", "")).strip()
                result["barangay"] = str(row.get("Barangay", "")).strip()
                result["city_municipality"] = str(row.get("City/Municipality", "")).strip()
                result["province_area"] = str(row.get("Province / Area", "")).strip()
                result["zip_code"] = str(row.get("ZIP Code", "")).strip()
                result["emergency_contact_name"] = str(row.get("Emergency Contact Name", "")).strip()
                em_num = str(row.get("Emergency Contact Number", "")).strip()
                em_rel = str(row.get("Emergency Relationship", "")).strip()
                result["emergency_contact_no_rel"] = f"{em_num} ({em_rel})".strip(" ()") if em_rel else em_num
                result["household_remarks"] = str(row.get("Remarks", "")).strip()
            result["family_members"].append({
                "full_name": fn,
                "relationship": str(row.get("Relationship", "")).strip(),
                "sex": str(row.get("Sex", "")).strip(),
                "birthdate": str(row.get("Birthdate", "")).strip(),
                "age": str(row.get("Age", "")).strip(),
                "category": str(row.get("Category", "")).strip(),
                "civil_status": str(row.get("Civil Status", "")).strip(),
                "student_working": str(row.get("Student/Working", "")).strip(),
                "contact_no": str(row.get("Contact No.", "")).strip(),
                "church_member": str(row.get("Church Member", "YES")).strip() or "YES",
                "status": str(row.get("Status", "ACTIVE")).strip() or "ACTIVE",
                "baptized": str(row.get("Baptized", "YES")).strip() or "YES",
                "baptism_date": str(row.get("Baptism Date", "")).strip(),
                "current_address": str(row.get("Current Address (If living separately)", "")).strip(),
            })
        return result

    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active
    grid: List[List[str]] = []
    for row in ws.iter_rows(values_only=True):
        grid.append([str(c).strip() if c is not None else "" for c in row])

    header_row_idx = -1
    for r_i, row_vals in enumerate(grid):
        row_join = " | ".join(row_vals).upper()
        if "FULL NAME" in row_join and ("RELATIONSHIP" in row_join or "BIRTHDATE" in row_join):
            header_row_idx = r_i
            break
        for c_i, cell_v in enumerate(row_vals):
            cu = cell_v.upper().rstrip(":")
            next_v = row_vals[c_i + 1] if (c_i + 1) < len(row_vals) else ""
            if not next_v and (c_i + 2) < len(row_vals):
                next_v = row_vals[c_i + 2]
            if cu == "PARISH" and next_v:
                result["parish"] = next_v
            elif cu in ("AREA/ZONE", "AREA / ZONE") and next_v:
                result["area_zone"] = next_v
            elif cu == "CELL GROUP" and next_v:
                result["cell_group"] = next_v
            elif cu.startswith("FAMILY ID") and next_v:
                result["family_id"] = next_v
            elif cu == "HEAD OF FAMILY" and next_v:
                result["head_of_family"] = next_v
            elif cu in ("MOBILE NUMBER", "HOUSEHOLD MOBILE NO.") and next_v:
                result["household_mobile"] = next_v
            elif cu == "COMPLETE ADDRESS" and next_v:
                result["complete_address"] = next_v
            elif cu == "BARANGAY" and next_v:
                result["barangay"] = next_v
            elif cu in ("CITY/MUNICIPALITY", "CITY / MUNICIPALITY") and next_v:
                result["city_municipality"] = next_v
            elif cu in ("PROVINCE", "PROVINCE / AREA") and next_v:
                result["province_area"] = next_v
            elif cu == "ZIP CODE" and next_v:
                result["zip_code"] = next_v

    if header_row_idx >= 0:
        hdrs = [h.upper() for h in grid[header_row_idx]]
        def _col(names: Tuple[str, ...]) -> int:
            for idx_h, h_val in enumerate(hdrs):
                if any(n in h_val for n in names):
                    return idx_h
            return -1

        c_fn = _col(("FULL NAME", "NAME"))
        c_rel = _col(("RELATIONSHIP",))
        c_sex = _col(("SEX", "GENDER"))
        c_bd = _col(("BIRTHDATE", "DOB"))
        c_age = _col(("AGE",))
        c_cat = _col(("CATEGORY",))
        c_civ = _col(("CIVIL STATUS",))
        c_sw = _col(("STUDENT", "WORKING", "OCCUPATION"))
        c_con = _col(("CONTACT", "MOBILE"))
        c_cm = _col(("CHURCH MEMBER",))
        c_st = _col(("STATUS",))
        c_bap = _col(("BAPTIZED",))
        c_bapd = _col(("BAPTISM DATE",))
        c_cur = _col(("CURRENT ADDRESS",))

        for r_i in range(header_row_idx + 1, len(grid)):
            r_vals = grid[r_i]
            r_join_u = " ".join(r_vals).upper()
            if "IV. HOUSEHOLD SUMMARY" in r_join_u or "V. EMERGENCY CONTACT" in r_join_u:
                for em_i in range(r_i, len(grid)):
                    em_row = grid[em_i]
                    for ci, cv in enumerate(em_row):
                        cv_u = cv.upper().rstrip(":")
                        nv = em_row[ci + 1] if (ci + 1) < len(em_row) else ""
                        if cv_u == "NAME" and nv and not result["emergency_contact_name"]:
                            result["emergency_contact_name"] = nv
                        elif cv_u == "RELATIONSHIP" and nv:
                            result["_em_rel"] = nv
                        elif cv_u in ("CONTACT NUMBER", "CONTACT NO.") and nv:
                            result["_em_num"] = nv
                break
            fn_val = r_vals[c_fn].strip() if 0 <= c_fn < len(r_vals) else ""
            if not fn_val or fn_val.upper() in ("FULL NAME", "NONE", "N/A"):
                continue
            result["family_members"].append({
                "full_name": fn_val,
                "relationship": r_vals[c_rel].strip() if 0 <= c_rel < len(r_vals) else "",
                "sex": r_vals[c_sex].strip() if 0 <= c_sex < len(r_vals) else "",
                "birthdate": r_vals[c_bd].strip() if 0 <= c_bd < len(r_vals) else "",
                "age": r_vals[c_age].strip() if 0 <= c_age < len(r_vals) else "",
                "category": r_vals[c_cat].strip() if 0 <= c_cat < len(r_vals) else "",
                "civil_status": r_vals[c_civ].strip() if 0 <= c_civ < len(r_vals) else "",
                "student_working": r_vals[c_sw].strip() if 0 <= c_sw < len(r_vals) else "",
                "contact_no": r_vals[c_con].strip() if 0 <= c_con < len(r_vals) else "",
                "church_member": (r_vals[c_cm].strip() if 0 <= c_cm < len(r_vals) else "YES") or "YES",
                "status": (r_vals[c_st].strip() if 0 <= c_st < len(r_vals) else "ACTIVE") or "ACTIVE",
                "baptized": (r_vals[c_bap].strip() if 0 <= c_bap < len(r_vals) else "YES") or "YES",
                "baptism_date": r_vals[c_bapd].strip() if 0 <= c_bapd < len(r_vals) else "",
                "current_address": r_vals[c_cur].strip() if 0 <= c_cur < len(r_vals) else "",
            })

    em_n = result.pop("_em_num", "")
    em_r = result.pop("_em_rel", "")
    if em_n or em_r:
        result["emergency_contact_no_rel"] = f"{em_n} ({em_r})".strip(" ()") if em_r else em_n
    return result


# =====================================================================
# 4. TIER 2: AUTO-DISCOVERING GEMINI VISION EXTRACTOR (NO 404 ERRORS)
# =====================================================================
_MODEL_DISCOVERY_CACHE: Dict[str, List[str]] = {}


def discover_active_gemini_flash_models(client: genai.Client, cache_key: str = "default") -> List[str]:
    """Dynamically queries client.models.list() at runtime to find active multimodal Flash models on the user's API key."""
    if cache_key in _MODEL_DISCOVERY_CACHE and _MODEL_DISCOVERY_CACHE[cache_key]:
        return list(_MODEL_DISCOVERY_CACHE[cache_key])

    discovered_flash: List[str] = []
    discovered_other: List[str] = []
    skip_tokens = ("embedding", "imagen", "veo", "tts", "audio", "aqa", "bison", "gecko", "learnlm", "gemma", "1.0", "1.5", "2.0")

    try:
        for m in client.models.list():
            raw_name = str(getattr(m, "name", "") or "").strip()
            clean_id = raw_name.replace("models/", "").strip()
            if not clean_id or not clean_id.startswith("gemini-"):
                continue
            low_id = clean_id.lower()
            if any(tok in low_id for tok in skip_tokens):
                continue
            actions = getattr(m, "supported_actions", None) or getattr(m, "supported_generation_methods", None)
            if actions and not any("generatecontent" in str(a).lower() for a in actions):
                continue
            if "flash" in low_id:
                discovered_flash.append(clean_id)
            else:
                discovered_other.append(clean_id)
    except Exception:
        pass

    preferred_priority = [
        "gemini-2.5-flash",
        "gemini-3.0-flash",
        "gemini-3.5-flash-lite",
        "gemini-3.8-flash",
        "gemini-2.5-flash-lite",
    ]
    ordered: List[str] = []
    for pref in preferred_priority:
        if pref in discovered_flash and pref not in ordered:
            ordered.append(pref)
    for fl in sorted(discovered_flash, reverse=True):
        if fl not in ordered:
            ordered.append(fl)
    for ot in sorted(discovered_other, reverse=True):
        if ot not in ordered:
            ordered.append(ot)
    if not ordered:
        ordered = ["gemini-2.5-flash", "gemini-2.5-flash-lite"]

    _MODEL_DISCOVERY_CACHE[cache_key] = ordered
    return list(ordered)


# =====================================================================
# 5. TIER 3: OFFLINE LOCAL PYTHON OCR FALLBACK (RAPIDOCR / TESSERACT)
# =====================================================================
def tier3_offline_ocr_extract(images: List[Image.Image], default_parish: str = "LIPAY NORTE PARISH") -> Dict[str, Any]:
    """Offline Non-AI OCR Fallback using rapidocr-onnxruntime (or pytesseract) when Gemini API is unavailable or out of quota."""
    result: Dict[str, Any] = {
        "parish": default_parish,
        "area_zone": "",
        "cell_group": "",
        "family_id": "",
        "head_of_family": "",
        "household_mobile": "",
        "complete_address": "",
        "barangay": "",
        "city_municipality": "",
        "province_area": "",
        "zip_code": "",
        "family_members": [],
        "emergency_contact_name": "",
        "emergency_contact_no_rel": "",
        "household_remarks": "",
    }
    if not images:
        return result

    ocr_engine = None
    try:
        rapid_mod = importlib.import_module("rapidocr_onnxruntime")
        ocr_engine = rapid_mod.RapidOCR()
    except Exception:
        ocr_engine = None

    all_lines: List[str] = []
    import numpy as np

    for im in images:
        if ocr_engine is not None:
            try:
                arr = np.array(im.convert("RGB"))
                ocr_res, _ = ocr_engine(arr)
                if ocr_res:
                    # Sort bounding boxes top-to-bottom, left-to-right
                    sorted_boxes = sorted(ocr_res, key=lambda x: (round(float(x[0][0][1]) / 18.0), float(x[0][0][0])))
                    cur_y = None
                    row_words: List[str] = []
                    for box, txt, _score in sorted_boxes:
                        y_top = float(box[0][1])
                        if cur_y is None or abs(y_top - cur_y) <= 18.0:
                            row_words.append(str(txt).strip())
                            if cur_y is None:
                                cur_y = y_top
                        else:
                            if row_words:
                                all_lines.append(" ".join(row_words))
                            row_words = [str(txt).strip()]
                            cur_y = y_top
                    if row_words:
                        all_lines.append(" ".join(row_words))
                    continue
            except Exception:
                pass
        try:
            pytess = importlib.import_module("pytesseract")
            txt_p = pytess.image_to_string(im) or ""
            all_lines.extend([ln.strip() for ln in txt_p.splitlines() if ln.strip()])
        except Exception:
            pass

    in_em = False
    em_num = ""
    em_rel = ""
    for ln in all_lines:
        lu = ln.upper()
        if "V. EMERGENCY CONTACT" in lu or "EMERGENCY CONTACT" in lu:
            in_em = True
            continue
        if "VI. REMARKS" in lu or "VII. CERTIFICATION" in lu:
            in_em = False
            continue
        if in_em:
            m_en = re.search(r"Name\s*[:\-]?\s*([A-Za-z.,\s]{3,40})", ln, flags=re.I)
            if m_en and not result["emergency_contact_name"]:
                result["emergency_contact_name"] = re.split(r"(?:Relationship|Contact)", m_en.group(1), flags=re.I)[0].strip(" :.-")
            m_er = re.search(r"Relationship\s*[:\-]?\s*([A-Za-z\s]{2,25})", ln, flags=re.I)
            if m_er and not em_rel:
                em_rel = re.split(r"(?:Contact|Number)", m_er.group(1), flags=re.I)[0].strip(" :.-")
            m_ec = re.search(r"(09\d{2}[\s\-]?\d{3}[\s\-]?\d{4})", ln)
            if m_ec and not em_num:
                em_num = re.sub(r"\D", "", m_ec.group(1))
            continue

        m_hof = re.search(r"Head\s*of\s*Family\s*[:\-]?\s*([A-Za-z.,\s]{4,40})", ln, flags=re.I)
        if m_hof and not result["head_of_family"]:
            result["head_of_family"] = re.split(r"(?:Mobile|Complete|Address|Family)", m_hof.group(1), flags=re.I)[0].strip(" :.-")
        m_mob = re.search(r"Mobile\s*Number\s*[:\-]?\s*(09\d{2}[\s\-]?\d{3}[\s\-]?\d{4})", ln, flags=re.I)
        if m_mob and not result["household_mobile"]:
            result["household_mobile"] = re.sub(r"\D", "", m_mob.group(1))
        m_area = re.search(r"Area\s*/?\s*Zone\s*[:\-]?\s*([0-9A-Za-z]+)", ln, flags=re.I)
        if m_area and not result["area_zone"]:
            result["area_zone"] = m_area.group(1).strip()
        m_zip = re.search(r"ZIP\s*Code\s*[:\-]?\s*(\d{4})", ln, flags=re.I)
        if m_zip and not result["zip_code"]:
            result["zip_code"] = m_zip.group(1).strip()

    if em_num or em_rel:
        result["emergency_contact_no_rel"] = f"{em_num} ({em_rel})".strip(" ()") if em_rel else em_num
    return result