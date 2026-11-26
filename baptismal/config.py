"""
baptismal/config.py — Configuration, Official 9 Parishes (Hong Kong Excluded) & Theme Palette
=============================================================================================
IMPORTANT RULE:
Only the 9 official Philippine parishes are included in the Baptismal System.
Hong Kong Parish is strictly excluded from all parish dashboards, parish admin
accounts, new registration options, and operational parish reporting.
"""

import json
import os
import re
from typing import Any, Dict, List

# Church-inspired Navy, White, and Gold Theme Palette
THEME_NAVY = "#0F2C59"
THEME_NAVY_DARK = "#091B38"
THEME_GOLD = "#D4AF37"
THEME_GOLD_SOFT = "#F6E7B2"
THEME_WHITE = "#FFFFFF"
THEME_SLATE_BG = "#F8FAFC"

# Verified CTG Baptismal Google Drive Folder & Workbook IDs
DEFAULT_BAPTISMAL_ROOT_FOLDER_ID = "1A_KW6rioDjckR_b9eU5A4eFVJAMx6o7e"
DEFAULT_BAPTISMAL_WORKBOOK_FILE_ID = "1GrO8Vpolg8KiHHWRoU1jlLv1iVtFR5Ic"

KNOWN_BAPTISMAL_SUBFOLDER_IDS: Dict[str, str] = {
    "BOOK 1": "1Y3xyOXC0MOuqtDnN1_Sn0yQIEXGb0KWV",
    "BOOK 2": "1LEwh-KS6JNemDb_uB9t7yC07p9SppKcr",
    "BOOK 3": "127B4WBu7Ul8m1q9VfvnTNupgkq1w6h1U",
    "BOOK 5": "1lFo5AoPuahzz2TSiVTZzL8wxBmbzoP-e",
    "NEW BAPTISM IMAGES": "1NF0xcDbN3AFXFqXHgMqMgyVV4WntfLkn",
    "GENERATED CERTIFICATES": "1IsoSJ4ZCaKcFk_Ki3CNiFY5e6hHK24U3",
    "CERTIFICATE REQUESTS": "1hIlmtf8SZEOkK25GcX8ukc-tWeu-oJ_5",
}

# OFFICIAL NINE PARISHES (Hong Kong Parish is strictly excluded)
OFFICIAL_NINE_PARISHES: List[str] = [
    "Lipay Norte Parish",
    "Malanas Parish",
    "Poblacion Parish",
    "Duplas Parish",
    "Santiago Parish",
    "Caba Parish",
    "Manila Parish",
    "Andabuen Parish",
    "Apayao Parish",
]

DEFAULT_NINE_PARISH_ADMIN_USERNAMES: Dict[str, str] = {
    "Lipay Norte Parish": "lipay_norte_bap_admin",
    "Malanas Parish": "malanas_bap_admin",
    "Poblacion Parish": "poblacion_bap_admin",
    "Duplas Parish": "duplas_bap_admin",
    "Santiago Parish": "santiago_bap_admin",
    "Caba Parish": "caba_bap_admin",
    "Manila Parish": "manila_bap_admin",
    "Andabuen Parish": "andabuen_bap_admin",
    "Apayao Parish": "apayao_bap_admin",
}

EXCLUDED_PARISH_KEYWORDS = ("HONG KONG", "HONGKONG")


def is_excluded_parish(parish_name: object) -> bool:
    """Returns True if the parish name refers to Hong Kong Parish (which is strictly excluded)."""
    s = str(parish_name or "").strip().upper()
    return any(k in s for k in EXCLUDED_PARISH_KEYWORDS)


def normalize_official_nine_parish(parish_name: object) -> str:
    """Matches a parish string against OFFICIAL_NINE_PARISHES; raises ValueError if Hong Kong or unknown."""
    raw = str(parish_name or "").strip()
    if not raw:
        raise ValueError("Parish is required.")
    if is_excluded_parish(raw):
        raise ValueError("Hong Kong Parish is excluded from the CTG Baptismal System.")
    raw_up = raw.upper().replace("PARISH", "").strip()
    for p in OFFICIAL_NINE_PARISHES:
        p_core = p.upper().replace("PARISH", "").strip()
        if raw.upper() == p.upper() or raw_up == p_core:
            return p
    raise ValueError(f"Parish '{raw}' is not one of the 9 official CTG Baptismal parishes.")


# Exact 27 Historical Columns from OVERALL BAPTISMAL CERTIFICATE LIST.xlsx
HISTORICAL_27_COLUMNS: List[str] = [
    "Parish",
    "Address",
    "Province",
    "Name of the Child",
    "Date of Baptism",
    "Place of Birth",
    "Filiation/Sex",
    "Nationality",
    "Date of Birth",
    "Name of the Father",
    "Father's Residence",
    "Name of the Mother",
    "Mother's Residence",
    "Paternal Grandparents",
    "Maternal Grandparents",
    "Vecindad",
    "Sponsors",
    "Civil Status",
    "Residences",
    "Officiating Priest",
    "Place of Baptism",
    "Day of Baptism",
    "Month of Baptism",
    "Year of Baptism",
    "Book Number",
    "Page Number",
    "Cura Parocco",
]

DRIVE_ROOT_FOLDER_NAME = "CTG BAPTISMAL RECORDS"
DRIVE_SUBFOLDERS_TREE: Dict[str, List[str]] = {
    "HISTORICAL RECORDS": [],
    "NEW BAPTISM IMAGES": ["2026", "2027", "2028"],
    "GENERATED CERTIFICATES": ["2026", "2027", "2028"],
    "BACKUPS": [],
}

HISTORICAL_BOOKS: List[str] = [f"BOOK {i}" for i in range(1, 10)]
YEARLY_WORKSHEET_PREFIX = "NEW BAPTISMS - "
SUPPORTED_NEW_YEARS: List[str] = ["2026", "2027", "2028"]

# Canonical Request Statuses (supports both Title Case and Uppercase lookups)
CANONICAL_REQUEST_STATUSES: List[str] = [
    "Pending",
    "Under Review",
    "Additional Information Required",
    "Approved",
    "Rejected",
    "Completed",
    "Expired",
    "Revoked",
]

REQUEST_STATUSES: List[str] = CANONICAL_REQUEST_STATUSES + [
    "PENDING",
    "UNDER REVIEW",
    "ADDITIONAL INFORMATION REQUIRED",
    "APPROVED",
    "REJECTED",
    "IN PROGRESS",
    "COMPLETED",
    "EXPIRED",
    "REVOKED",
]

REGISTRATION_VERIFICATION_STATUSES: List[str] = [
    "Pending Verification",
    "Additional Information Required",
    "Approved",
    "Rejected",
]

EDITABLE_REGISTRATION_STATUSES = {
    "Pending Verification",
    "Additional Information Required",
    "PENDING_VERIFICATION",
    "ADDITIONAL_INFO_REQUIRED",
    "REGISTERED",
}

TEMP_ACCESS_PERMISSIONS: List[str] = [
    "view_record",
    "view_image",
    "generate_certificate",
    "download_certificate",
    "print_certificate",
]

ALLOWED_UPLOAD_IMAGE_EXTS = {"jpg", "jpeg", "png", "webp", "tif", "tiff"}
ALLOWED_UPLOAD_DOC_EXTS = {"pdf", "xlsx", "csv"}
MAX_UPLOAD_SIZE_MB = 20

# Security defaults
MAX_FAILED_LOGIN_ATTEMPTS = 5
LOCKOUT_DURATION_SECONDS = 900  # 15 minutes
SESSION_TIMEOUT_SECONDS = 3600  # 60 minutes
PBKDF2_ITERATIONS = 310000


def _read_local_hub_config() -> Dict[str, Any]:
    for fn in ("ctg_app_cloud_config.json", ".streamlit/secrets.toml", "secrets.toml"):
        if os.path.exists(fn):
            try:
                if fn.endswith(".json"):
                    with open(fn, "r", encoding="utf-8") as f:
                        d = json.load(f)
                    if isinstance(d, dict):
                        return d
                else:
                    with open(fn, "r", encoding="utf-8", errors="ignore") as f:
                        txt = f.read()
                    parsed: Dict[str, Any] = {}
                    for m in re.finditer(r'^([A-Za-z0-9_]+)\s*=\s*["\']([^"\']+)["\']', txt, flags=re.M):
                        parsed[m.group(1)] = m.group(2)
                    if parsed:
                        return parsed
            except Exception:
                continue
    return {}


def get_config_value(key: str, default: str = "") -> str:
    """Safely reads configuration from os.environ, st.secrets, st.session_state.custom_secrets, or local hub config."""
    env_val = os.environ.get(key)
    if env_val is not None and str(env_val).strip() != "":
        return str(env_val).strip()
    try:
        import streamlit as st  # type: ignore
        custom_s = st.session_state.get("custom_secrets", {})
        if isinstance(custom_s, dict) and str(custom_s.get(key, "")).strip():
            return str(custom_s[key]).strip()
        if key in st.secrets:
            val = st.secrets[key]
            if val is not None and str(val).strip() != "":
                return str(val).strip()
    except Exception:
        pass
    local_cfg = _read_local_hub_config()
    if str(local_cfg.get(key, "")).strip():
        return str(local_cfg[key]).strip()
    return default


def get_database_url() -> str:
    """Returns PostgreSQL connection URL from BAPTISMAL_DATABASE_URL or DATABASE_URL, or empty if unconfigured."""
    return (
        get_config_value("BAPTISMAL_DATABASE_URL")
        or get_config_value("DATABASE_URL")
        or ""
    )


def get_google_drive_config() -> Dict[str, Any]:
    """Reads Google Drive OAuth / Service Account configuration and designated file/folder IDs."""
    sa_json: Any = (
        get_config_value("BAPTISMAL_GOOGLE_SERVICE_ACCOUNT_JSON")
        or get_config_value("GOOGLE_SERVICE_ACCOUNT_JSON")
    )
    if not sa_json:
        try:
            import streamlit as st  # type: ignore
            for sec_k in ("baptismal_service_account", "gcp_service_account", "google_service_account", "service_account"):
                if sec_k in st.secrets:
                    sa_json = dict(st.secrets[sec_k])
                    break
        except Exception:
            pass

    return {
        "service_account": sa_json,
        "client_id": (
            get_config_value("BAPTISMAL_GOOGLE_CLIENT_ID")
            or get_config_value("GOOGLE_CLIENT_ID")
            or get_config_value("GMAIL_CLIENT_ID")
        ),
        "client_secret": (
            get_config_value("BAPTISMAL_GOOGLE_CLIENT_SECRET")
            or get_config_value("GOOGLE_CLIENT_SECRET")
            or get_config_value("GMAIL_CLIENT_SECRET")
        ),
        "refresh_token": (
            get_config_value("BAPTISMAL_DRIVE_REFRESH_TOKEN")
            or get_config_value("GOOGLE_DRIVE_REFRESH_TOKEN")
            or get_config_value("GMAIL_REFRESH_TOKEN")
        ),
        "root_folder_id": get_config_value("BAPTISMAL_ROOT_FOLDER_ID", DEFAULT_BAPTISMAL_ROOT_FOLDER_ID),
        "source_files_folder_id": get_config_value("BAPTISMAL_SOURCE_FOLDER_ID", DEFAULT_BAPTISMAL_ROOT_FOLDER_ID),
        "workbook_file_id": get_config_value("BAPTISMAL_WORKBOOK_FILE_ID", DEFAULT_BAPTISMAL_WORKBOOK_FILE_ID),
    }
