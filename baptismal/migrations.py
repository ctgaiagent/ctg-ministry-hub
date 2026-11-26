"""
baptismal/migrations.py — Idempotent Database Schema & Migrations
=================================================================
Includes `staged_registration_files` so Parish Leader Admin uploads are staged
for Master Admin confirmation before encoding into Google Drive Excel and Drive folders.
"""

from datetime import datetime, timezone
from typing import List, Tuple

from .database import DatabaseManager

MIGRATIONS: List[Tuple[str, List[str]]] = [
    (
        "001_initial_schema",
        [
            """
            CREATE TABLE IF NOT EXISTS schema_migrations (
                version TEXT PRIMARY KEY,
                applied_at TEXT NOT NULL
            )
            """,
            """
            CREATE TABLE IF NOT EXISTS baptismal_accounts (
                account_id TEXT PRIMARY KEY,
                username TEXT UNIQUE NOT NULL,
                full_name TEXT NOT NULL,
                email TEXT,
                role TEXT NOT NULL,
                assigned_parish TEXT NOT NULL DEFAULT '',
                password_hash TEXT NOT NULL,
                is_active INTEGER NOT NULL DEFAULT 1,
                failed_attempts INTEGER NOT NULL DEFAULT 0,
                locked_until TEXT,
                must_change_password INTEGER NOT NULL DEFAULT 0,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            )
            """,
            """
            CREATE TABLE IF NOT EXISTS baptismal_roles_permissions (
                role_name TEXT NOT NULL,
                permission_key TEXT NOT NULL,
                PRIMARY KEY (role_name, permission_key)
            )
            """,
            """
            CREATE TABLE IF NOT EXISTS historical_records (
                record_id TEXT PRIMARY KEY,
                source_row_number INTEGER UNIQUE NOT NULL,
                parish TEXT,
                address TEXT,
                province TEXT,
                child_name TEXT NOT NULL,
                date_of_baptism TEXT,
                place_of_birth TEXT,
                filiation_sex TEXT,
                nationality TEXT,
                date_of_birth TEXT,
                father_name TEXT,
                fathers_residence TEXT,
                mother_name TEXT,
                mothers_residence TEXT,
                paternal_grandparents TEXT,
                maternal_grandparents TEXT,
                vecindad TEXT,
                sponsors TEXT,
                civil_status TEXT,
                residences TEXT,
                officiating_priest TEXT,
                place_of_baptism TEXT,
                day_of_baptism TEXT,
                month_of_baptism TEXT,
                year_of_baptism TEXT,
                book_number TEXT,
                page_number TEXT,
                cura_parocco TEXT,
                raw_json TEXT NOT NULL,
                created_at TEXT NOT NULL
            )
            """,
            "CREATE INDEX IF NOT EXISTS idx_hist_child_name ON historical_records(child_name)",
            "CREATE INDEX IF NOT EXISTS idx_hist_book_page ON historical_records(book_number, page_number)",
            "CREATE INDEX IF NOT EXISTS idx_hist_parish_year ON historical_records(parish, year_of_baptism)",
            """
            CREATE TABLE IF NOT EXISTS new_baptism_registrations (
                record_id TEXT PRIMARY KEY,
                baptism_year TEXT NOT NULL,
                worksheet_name TEXT NOT NULL,
                registration_status TEXT NOT NULL,
                verification_status TEXT NOT NULL DEFAULT 'Pending Verification',
                verified_by TEXT,
                verified_at TEXT,
                verification_notes TEXT,
                upload_status TEXT NOT NULL,
                workbook_update_status TEXT NOT NULL,
                parish TEXT NOT NULL,
                child_name TEXT NOT NULL,
                date_of_baptism TEXT,
                father_name TEXT,
                mother_name TEXT,
                book_number TEXT,
                page_number TEXT,
                dedup_fingerprint TEXT UNIQUE NOT NULL,
                payload_json TEXT NOT NULL,
                created_by TEXT NOT NULL,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            )
            """,
            "CREATE INDEX IF NOT EXISTS idx_new_bap_year ON new_baptism_registrations(baptism_year)",
            "CREATE INDEX IF NOT EXISTS idx_new_bap_parish ON new_baptism_registrations(parish)",
            """
            CREATE TABLE IF NOT EXISTS staged_registration_files (
                staged_id TEXT PRIMARY KEY,
                record_id TEXT NOT NULL,
                filename TEXT NOT NULL,
                mime_type TEXT NOT NULL,
                file_b64 TEXT NOT NULL,
                file_size_bytes INTEGER NOT NULL,
                uploaded_by TEXT NOT NULL,
                is_pushed_to_drive INTEGER NOT NULL DEFAULT 0,
                drive_file_id TEXT,
                created_at TEXT NOT NULL
            )
            """,
            "CREATE INDEX IF NOT EXISTS idx_staged_rec ON staged_registration_files(record_id)",
            """
            CREATE TABLE IF NOT EXISTS drive_file_mappings (
                mapping_id TEXT PRIMARY KEY,
                record_id TEXT NOT NULL,
                category TEXT NOT NULL,
                original_filename TEXT NOT NULL,
                drive_file_id TEXT NOT NULL,
                drive_folder_id TEXT NOT NULL,
                year_folder TEXT,
                mime_type TEXT NOT NULL,
                file_size_bytes INTEGER NOT NULL,
                upload_status TEXT NOT NULL,
                uploaded_by TEXT NOT NULL,
                created_at TEXT NOT NULL
            )
            """,
            "CREATE INDEX IF NOT EXISTS idx_drive_map_rec ON drive_file_mappings(record_id)",
            """
            CREATE TABLE IF NOT EXISTS historical_image_associations (
                association_id TEXT PRIMARY KEY,
                record_id TEXT NOT NULL,
                book_folder TEXT NOT NULL,
                image_filename TEXT NOT NULL,
                drive_file_id TEXT,
                extracted_page_hint TEXT,
                record_page_number TEXT,
                match_status TEXT NOT NULL,
                confirmed_by TEXT,
                confirmed_at TEXT,
                notes TEXT,
                created_at TEXT NOT NULL,
                UNIQUE(record_id, book_folder, image_filename)
            )
            """,
            "CREATE INDEX IF NOT EXISTS idx_hist_img_rec ON historical_image_associations(record_id)",
            """
            CREATE TABLE IF NOT EXISTS certificate_requests (
                request_id TEXT PRIMARY KEY,
                request_type TEXT NOT NULL DEFAULT 'CERTIFIED_CERTIFICATE',
                requester_account_id TEXT NOT NULL,
                requester_username TEXT NOT NULL,
                assigned_parish TEXT NOT NULL DEFAULT '',
                target_record_id TEXT,
                requested_person_name TEXT NOT NULL,
                father_name TEXT,
                mother_name TEXT,
                approx_baptism_date TEXT,
                approx_birth_date TEXT,
                book_hint TEXT,
                page_hint TEXT,
                purpose TEXT NOT NULL,
                supporting_doc_filename TEXT,
                supporting_doc_drive_id TEXT,
                approved_certificate_id TEXT,
                status TEXT NOT NULL,
                reviewed_by TEXT,
                reviewed_at TEXT,
                review_notes TEXT,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            )
            """,
            "CREATE INDEX IF NOT EXISTS idx_cert_req_status ON certificate_requests(status)",
            """
            CREATE TABLE IF NOT EXISTS temporary_access_grants (
                grant_id TEXT PRIMARY KEY,
                request_id TEXT REFERENCES certificate_requests(request_id),
                account_id TEXT NOT NULL REFERENCES baptismal_accounts(account_id),
                record_id TEXT NOT NULL,
                permissions_json TEXT NOT NULL,
                expires_at TEXT NOT NULL,
                is_revoked INTEGER NOT NULL DEFAULT 0,
                revoked_by TEXT,
                revoked_at TEXT,
                granted_by TEXT NOT NULL,
                created_at TEXT NOT NULL
            )
            """,
            "CREATE INDEX IF NOT EXISTS idx_temp_grant_acc_rec ON temporary_access_grants(account_id, record_id)",
            """
            CREATE TABLE IF NOT EXISTS generated_certificates (
                certificate_id TEXT PRIMARY KEY,
                record_id TEXT NOT NULL,
                request_id TEXT,
                authorized_account_id TEXT,
                assigned_parish TEXT,
                issued_by TEXT NOT NULL,
                issuance_date TEXT NOT NULL,
                issuance_status TEXT NOT NULL,
                approval_status TEXT NOT NULL DEFAULT 'ISSUED',
                approved_by TEXT,
                approved_at TEXT,
                is_released_to_requester INTEGER NOT NULL DEFAULT 0,
                reissue_reason TEXT,
                drive_file_id TEXT,
                pdf_sha256 TEXT NOT NULL,
                layout_orientation TEXT NOT NULL,
                created_at TEXT NOT NULL
            )
            """,
            "CREATE INDEX IF NOT EXISTS idx_gen_cert_rec ON generated_certificates(record_id)",
            """
            CREATE TABLE IF NOT EXISTS baptismal_audit_logs (
                event_id TEXT PRIMARY KEY,
                account_id TEXT NOT NULL,
                event_type TEXT NOT NULL,
                related_record_id TEXT,
                timestamp TEXT NOT NULL,
                outcome TEXT NOT NULL,
                metadata_json TEXT NOT NULL
            )
            """,
            "CREATE INDEX IF NOT EXISTS idx_audit_ts ON baptismal_audit_logs(timestamp DESC)",
            "CREATE INDEX IF NOT EXISTS idx_audit_type ON baptismal_audit_logs(event_type)",
            """
            CREATE TABLE IF NOT EXISTS workbook_update_logs (
                update_id TEXT PRIMARY KEY,
                record_id TEXT NOT NULL,
                worksheet_name TEXT NOT NULL,
                drive_file_id TEXT,
                before_revision_id TEXT,
                after_revision_id TEXT,
                backup_drive_file_id TEXT,
                status TEXT NOT NULL,
                error_message TEXT,
                actor_account_id TEXT NOT NULL,
                created_at TEXT NOT NULL
            )
            """,
        ],
    ),
]

_UPGRADE_COLUMNS: List[Tuple[str, str, str]] = [
    ("baptismal_accounts", "assigned_parish", "TEXT NOT NULL DEFAULT ''"),
    ("new_baptism_registrations", "verification_status", "TEXT NOT NULL DEFAULT 'Pending Verification'"),
    ("new_baptism_registrations", "verified_by", "TEXT"),
    ("new_baptism_registrations", "verified_at", "TEXT"),
    ("new_baptism_registrations", "verification_notes", "TEXT"),
    ("certificate_requests", "request_type", "TEXT NOT NULL DEFAULT 'CERTIFIED_CERTIFICATE'"),
    ("certificate_requests", "assigned_parish", "TEXT NOT NULL DEFAULT ''"),
    ("certificate_requests", "father_name", "TEXT"),
    ("certificate_requests", "mother_name", "TEXT"),
    ("certificate_requests", "approx_baptism_date", "TEXT"),
    ("certificate_requests", "approx_birth_date", "TEXT"),
    ("certificate_requests", "book_hint", "TEXT"),
    ("certificate_requests", "page_hint", "TEXT"),
    ("certificate_requests", "supporting_doc_filename", "TEXT"),
    ("certificate_requests", "supporting_doc_drive_id", "TEXT"),
    ("certificate_requests", "approved_certificate_id", "TEXT"),
    ("generated_certificates", "request_id", "TEXT"),
    ("generated_certificates", "authorized_account_id", "TEXT"),
    ("generated_certificates", "assigned_parish", "TEXT"),
    ("generated_certificates", "approval_status", "TEXT NOT NULL DEFAULT 'ISSUED'"),
    ("generated_certificates", "approved_by", "TEXT"),
    ("generated_certificates", "approved_at", "TEXT"),
    ("generated_certificates", "is_released_to_requester", "INTEGER NOT NULL DEFAULT 0"),
]


def run_migrations(db: DatabaseManager) -> List[str]:
    """Applies any unapplied database migrations and column/table upgrades idempotently."""
    db.execute(
        """
        CREATE TABLE IF NOT EXISTS schema_migrations (
            version TEXT PRIMARY KEY,
            applied_at TEXT NOT NULL
        )
        """
    )
    applied_rows = db.fetch_all("SELECT version FROM schema_migrations")
    applied_set = {r["version"] for r in applied_rows}
    newly_applied: List[str] = []

    for version, statements in MIGRATIONS:
        if version in applied_set:
            continue
        with db.transaction() as cur:
            for stmt in statements:
                cur.execute(db._adapt_sql(stmt))
            cur.execute(
                db._adapt_sql("INSERT INTO schema_migrations (version, applied_at) VALUES (%s, %s)"),
                (version, datetime.now(timezone.utc).isoformat()),
            )
        newly_applied.append(version)

    # Ensure staged_registration_files exists even on older databases
    db.execute(
        """
        CREATE TABLE IF NOT EXISTS staged_registration_files (
            staged_id TEXT PRIMARY KEY,
            record_id TEXT NOT NULL,
            filename TEXT NOT NULL,
            mime_type TEXT NOT NULL,
            file_b64 TEXT NOT NULL,
            file_size_bytes INTEGER NOT NULL,
            uploaded_by TEXT NOT NULL,
            is_pushed_to_drive INTEGER NOT NULL DEFAULT 0,
            drive_file_id TEXT,
            created_at TEXT NOT NULL
        )
        """
    )

    for tbl, col, col_def in _UPGRADE_COLUMNS:
        try:
            db.execute(f"ALTER TABLE {tbl} ADD COLUMN {col} {col_def}")
        except Exception:
            pass

    role_perms = [
        ("Master Admin", "full_system_access"),
        ("Master Admin", "view_historical"),
        ("Master Admin", "view_book_archive"),
        ("Master Admin", "register_new_baptism"),
        ("Master Admin", "verify_registrations"),
        ("Master Admin", "manage_temporary_access"),
        ("Master Admin", "generate_certificate"),
        ("Master Admin", "approve_and_issue_certificate"),
        ("Master Admin", "view_audit_logs"),
        ("Master Admin", "manage_backups"),
        ("Parish Leader Admin", "register_own_parish_baptism"),
        ("Parish Leader Admin", "edit_pending_own_parish_baptism"),
        ("Parish Leader Admin", "request_historical_record"),
        ("Parish Leader Admin", "request_certified_certificate"),
        ("Parish Leader Admin", "download_released_certificate"),
        ("Temporary Requester", "view_approved_record_only"),
    ]
    for role_n, perm_k in role_perms:
        exists = db.fetch_one(
            "SELECT 1 AS ok FROM baptismal_roles_permissions WHERE role_name = %s AND permission_key = %s",
            (role_n, perm_k),
        )
        if not exists:
            db.execute(
                "INSERT INTO baptismal_roles_permissions (role_name, permission_key) VALUES (%s, %s)",
                (role_n, perm_k),
            )
    return newly_applied
