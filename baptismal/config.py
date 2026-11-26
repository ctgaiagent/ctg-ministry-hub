"""
baptismal/config.py — Configuration, Constants, Secrets Loader & Theme Palette
==============================================================================
No hardcoded passwords, API keys, OAuth tokens, or invented Google Drive IDs.
All external credentials and Drive IDs are read exclusively from environment
variables or Streamlit Secrets (`st.secrets`).
"""

import os
from typing import Any, Dict, List

# Church-inspired Navy, White, and Gold Theme Palette
THEME_NAVY = "#0F2C59"
THEME_NAVY_DARK = "#091B38"
THEME_GOLD = "#D4AF37"
THEME_GOLD_SOFT = "#F6E7B2"
THEME_WHITE = "#FFFFFF"
THEME_SLATE_BG = "#F8FAFC"

# Exact 27 Historical Columns from OVERALL BAPTISMAL CERTIFICATE LIST.xlsx (Sheet1)
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

# Private Google Drive Folder Architecture
DRIVE_ROOT_FOLDER_NAME = "CTG BAPTISMAL SYSTEM"
DRIVE_SUBFOLDERS_TREE: Dict[str, List[str]] = {
    "HISTORICAL RECORDS": [],
    "NEW BAPTISM IMAGES": ["2026", "2027", "2028"],
    "GENERATED CERTIFICATES": ["2026", "2027", "2028"],
    "BACKUPS": [],
}

HISTORICAL_BOOKS: List[str] = [f"BOOK {i}" for i in range(1, 10)]
YEARLY_WORKSHEET_PREFIX = "NEW BAPTISMS - "
SUPPORTED_NEW_YEARS: List[str] = ["2026", "2027", "2028"]

REQUEST_STATUSES: List[str] = [
    "PENDING",
    "APPROVED",
    "REJECTED",
    "IN PROGRESS",
    "COMPLETED",
    "EXPIRED",
    "REVOKED",
]

TEMP_ACCESS_PERMISSIONS: List[str] = [
    "view_record",
    "view_image",
    "generate_certificate",
    "download_certificate",
    "print_certificate",
]

ALLOWED_UPLOAD_IMAGE_EXTS = {"jpg", "jpeg", "png", "webp", "tif", "tiff"}
ALLOWED_UPLOAD_DOC_EXTS = {"pdf", "xlsx"}
MAX_UPLOAD_SIZE_MB = 20

# Security defaults
MAX_FAILED_LOGIN_ATTEMPTS = 5
LOCKOUT_DURATION_SECONDS = 900  # 15 minutes
SESSION_TIMEOUT_SECONDS = 3600  # 60 minutes
PBKDF2_ITERATIONS = 310000


def get_config_value(key: str, default: str = "") -> str:
    """Safely reads configuration from os.environ first, then st.secrets if available."""
    env_val = os.environ.get(key)
    if env_val is not None and str(env_val).strip() != "":
        return str(env_val).strip()
    try:
        import streamlit as st  # type: ignore
        if key in st.secrets:
            val = st.secrets[key]
            if val is not None and str(val).strip() != "":
                return str(val).strip()
    except Exception:
        pass
    return default


def get_database_url() -> str:
    """Returns PostgreSQL connection URL from BAPTISMAL_DATABASE_URL or DATABASE_URL, or empty if unconfigured."""
    return (
        get_config_value("BAPTISMAL_DATABASE_URL")
        or get_config_value("DATABASE_URL")
        or ""
    )


def get_google_drive_config() -> Dict[str, Any]:
    """Reads Google Drive OAuth / Service Account configuration and designated file/folder IDs.
    Never invents Drive file IDs or credentials."""
    sa_json = get_config_value("BAPTISMAL_GOOGLE_SERVICE_ACCOUNT_JSON") or get_config_value("GOOGLE_SERVICE_ACCOUNT_JSON")
    if not sa_json:
        try:
            import streamlit as st  # type: ignore
            for sec_k in ("baptismal_service_account", "gcp_service_account", "google_service_account"):
                if sec_k in st.secrets:
                    sa_json = dict(st.secrets[sec_k])
                    break
        except Exception:
            pass

    return {
        "service_account": sa_json,
        "client_id": get_config_value("BAPTISMAL_GOOGLE_CLIENT_ID") or get_config_value("GOOGLE_CLIENT_ID") or get_config_value("GMAIL_CLIENT_ID"),
        "client_secret": get_config_value("BAPTISMAL_GOOGLE_CLIENT_SECRET") or get_config_value("GOOGLE_CLIENT_SECRET") or get_config_value("GMAIL_CLIENT_SECRET"),
        "refresh_token": get_config_value("BAPTISMAL_DRIVE_REFRESH_TOKEN") or get_config_value("GOOGLE_DRIVE_REFRESH_TOKEN"),
        "root_folder_id": get_config_value("BAPTISMAL_ROOT_FOLDER_ID"),
        "source_files_folder_id": get_config_value("BAPTISMAL_SOURCE_FOLDER_ID"),
        "workbook_file_id": get_config_value("BAPTISMAL_WORKBOOK_FILE_ID"),
    }
