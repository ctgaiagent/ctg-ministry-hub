# CTG Baptismal Records & Certification System — Deployment & Security Guide

## 1. Architecture Summary
- **Existing Hub Preserved:** `app.py` (`streamlit run app.py`) remains completely untouched.
- **New Separate Entry Point:** `baptismal_app.py` (`streamlit run baptismal_app.py`).
- **Modular Backend Package (`baptismal/`):**
  - `config.py`, `validation.py`, `database.py`, `migrations.py`, `audit_logs.py`
  - `authentication.py` (PBKDF2-HMAC-SHA256 310,000 iterations, lockout after 5 failures, no hardcoded passwords)
  - `authorization.py` (Server-side Master Admin & record-specific Temporary Access checks)
  - `google_drive.py` (Private folder tree `CTG BAPTISMAL SYSTEM`, `canEdit` workbook check, revision concurrency)
  - `excel_manager.py` (Appends to `NEW BAPTISMS - YYYY` while keeping historical `Sheet1` 100% read-only)
  - `historical_records.py`, `image_archive.py`, `image_matching.py`, `registration.py`, `certificates.py`, `temporary_access.py`, `backup_manager.py`
- **PDF Certificate Template:** `templates/baptismal_certificate.py` (A4 Portrait & Landscape, Navy/White/Gold theme, never invents missing fields).

## 2. Local Testing Commands (macOS / VS Code Terminal)
```bash
# 1. Run all automated unit tests (uses temporary SQLite & mocked Drive)
python3 -m unittest discover -s tests -v

# 2. Run the existing CTG Ministry Hub (unchanged)
streamlit run app.py

# 3. Run the new Baptismal Records & Certification System independently
streamlit run baptismal_app.py
```

## 3. PostgreSQL Setup (Production)
1. Provision a managed PostgreSQL database (e.g., Supabase, Neon, or AWS RDS) with SSL required.
2. Add `BAPTISMAL_DATABASE_URL = "postgresql://USER:PASSWORD@HOST:5432/baptismal_db?sslmode=require"` to Streamlit Secrets.
3. On first launch, `baptismal/migrations.py` automatically creates all 12 tables, foreign keys, and indexes idempotently.

## 4. Google Drive OAuth & Permission Verification
1. Grant **Editor** access on `OVERALL BAPTISMAL CERTIFICATE LIST.xlsx` and the private `CTG BAPTISMAL SYSTEM` folder to your designated OAuth account or Google Service Account (`client_email`).
2. Add `BAPTISMAL_WORKBOOK_FILE_ID`, `BAPTISMAL_ROOT_FOLDER_ID`, and either OAuth (`BAPTISMAL_GOOGLE_CLIENT_ID`, `BAPTISMAL_GOOGLE_CLIENT_SECRET`, `BAPTISMAL_DRIVE_REFRESH_TOKEN`) or `[baptismal_service_account]` to Streamlit Secrets using `secrets_baptismal_template.toml` as a guide.
3. Keep all folders private — **never** enable "Anyone with the link".
4. Open **10. Google Drive Storage Status** inside `baptismal_app.py` and click **Verify Workbook Editor Permission & Private Folder Tree**.

## 5. Master Admin Initial Provisioning
- **Option A (Recommended Interactive Setup):** Launch `streamlit run baptismal_app.py` with an empty database; the app prompts you to create the first Master Admin account securely without putting any password in code or logs.
- **Option B (Secrets Bootstrap):** Set `BAPTISMAL_INITIAL_ADMIN_USER` and `BAPTISMAL_INITIAL_ADMIN_PASSWORD_HASH` in Streamlit Secrets before first boot, then remove them afterward.

## 6. Separate Streamlit Community Cloud Deployment
1. Push `baptismal_app.py`, `baptismal/`, `templates/`, `tests/`, and `requirements.txt` to your GitHub repository.
2. In Streamlit Community Cloud, click **New app**, select the same repository, and set **Main file path** to `baptismal_app.py` (leaving your existing `app.py` deployment untouched).
3. Paste your filled `secrets_baptismal_template.toml` values into **App Settings → Secrets**.

## 7. Integrations Requiring Your Authorization (Not Tested Live)
- **Live Google Drive API (`CTG BAPTISMAL SYSTEM` & `OVERALL BAPTISMAL CERTIFICATE LIST.xlsx`):** Tested locally using `unittest.mock`; live Drive OAuth/Service Account verification requires your real Drive IDs and credentials in Streamlit Secrets.
- **Live PostgreSQL Server:** Tested locally using SQLite; live `psycopg2` connection requires your `BAPTISMAL_DATABASE_URL` in Streamlit Secrets.
- **Live Streamlit UI E2E Browser Run:** Requires running `streamlit run baptismal_app.py` in your environment where `streamlit` is installed.
