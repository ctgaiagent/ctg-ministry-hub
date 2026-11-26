"""
baptismal/google_drive.py — Least-Privilege Private Google Drive Integration
============================================================================
- Reuses Google OAuth Refresh Token or Service Account credentials from st.secrets / env.
- Verifies Editor permission (`capabilities.canEdit`) on the designated Excel workbook BEFORE writing.
- Never enables public sharing ("Anyone with the link").
- Never invents Drive file IDs or credentials.
- Organizes private folders:
  CTG BAPTISMAL SYSTEM/
    HISTORICAL RECORDS/
    NEW BAPTISM IMAGES/2026..2028/
    GENERATED CERTIFICATES/2026..2028/
    BACKUPS/
"""

import io
import json
import time
from typing import Any, Dict, List, Optional, Tuple

from .config import DRIVE_ROOT_FOLDER_NAME, DRIVE_SUBFOLDERS_TREE, get_google_drive_config


class GoogleDrivePermissionError(PermissionError):
    """Raised when OAuth scopes or Drive file permissions are insufficient for writing."""


class GoogleDriveManager:
    """Backend-only Google Drive API v3 manager with strict permission checks and safe retries."""

    DRIVE_SCOPES = ["https://www.googleapis.com/auth/drive"]

    def __init__(self, drive_service: Optional[Any] = None, config: Optional[Dict[str, Any]] = None):
        self.config = config or get_google_drive_config()
        self._service = drive_service

    def get_service(self):
        if self._service is not None:
            return self._service

        sa_raw = self.config.get("service_account")
        if sa_raw:
            try:
                from google.oauth2 import service_account  # type: ignore
                from googleapiclient.discovery import build  # type: ignore

                sa_info = sa_raw if isinstance(sa_raw, dict) else json.loads(str(sa_raw))
                creds = service_account.Credentials.from_service_account_info(sa_info, scopes=self.DRIVE_SCOPES)
                self._service = build("drive", "v3", credentials=creds, cache_discovery=False)
                return self._service
            except Exception:
                pass

        cid = str(self.config.get("client_id") or "").strip()
        csec = str(self.config.get("client_secret") or "").strip()
        rt = str(self.config.get("refresh_token") or "").strip()
        if cid and csec and rt:
            try:
                from google.oauth2.credentials import Credentials  # type: ignore
                from googleapiclient.discovery import build  # type: ignore

                creds = Credentials(
                    token=None,
                    refresh_token=rt,
                    token_uri="https://oauth2.googleapis.com/token",
                    client_id=cid,
                    client_secret=csec,
                    scopes=self.DRIVE_SCOPES,
                )
                self._service = build("drive", "v3", credentials=creds, cache_discovery=False)
                return self._service
            except Exception as ex:
                raise GoogleDrivePermissionError(f"Failed to initialize Google Drive OAuth client: {ex}")

        raise GoogleDrivePermissionError(
            "Google Drive credentials are not configured. Set BAPTISMAL_GOOGLE_SERVICE_ACCOUNT_JSON or OAuth credentials in Streamlit Secrets."
        )

    def _execute_with_retry(self, request_callable, max_retries: int = 3):
        last_err = None
        for attempt in range(max_retries):
            try:
                return request_callable()
            except Exception as ex:
                last_err = ex
                err_s = str(ex).lower()
                if "403" in err_s or "insufficient" in err_s or "404" in err_s:
                    raise
                if attempt < max_retries - 1:
                    time.sleep(0.8 * (attempt + 1))
        raise RuntimeError(f"Google Drive API request failed after {max_retries} attempts: {last_err}")

    def verify_workbook_write_access(self, file_id: str) -> Dict[str, Any]:
        """Verifies that the backend OAuth/Service Account has Editor permission on the designated workbook."""
        if not file_id:
            raise GoogleDrivePermissionError("BAPTISMAL_WORKBOOK_FILE_ID is not configured.")
        svc = self.get_service()
        meta = self._execute_with_retry(
            lambda: svc.files()
            .get(
                fileId=file_id,
                fields="id, name, mimeType, capabilities(canEdit, canDownload), headRevisionId",
                supportsAllDrives=True,
            )
            .execute()
        )
        caps = meta.get("capabilities") or {}
        if not caps.get("canEdit", False):
            raise GoogleDrivePermissionError(
                f"Insufficient Google Drive permissions on workbook '{meta.get('name', file_id)}': backend account is Viewer-only (canEdit=False). Write operation aborted."
            )
        return meta

    def ensure_private_folder(self, folder_name: str, parent_id: str = "") -> str:
        """Finds or creates a private folder without any public 'anyone' permission."""
        svc = self.get_service()
        safe_q_name = folder_name.replace("'", "\\'")
        q = f"name = '{safe_q_name}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
        if parent_id:
            q += f" and '{parent_id}' in parents"
        res = self._execute_with_retry(
            lambda: svc.files().list(q=q, fields="files(id, name)", supportsAllDrives=True, includeItemsFromAllDrives=True).execute()
        )
        files = res.get("files") or []
        if files:
            return str(files[0]["id"])

        body: Dict[str, Any] = {
            "name": folder_name,
            "mimeType": "application/vnd.google-apps.folder",
        }
        if parent_id:
            body["parents"] = [parent_id]
        created = self._execute_with_retry(
            lambda: svc.files().create(body=body, fields="id", supportsAllDrives=True).execute()
        )
        return str(created["id"])

    def ensure_baptismal_folder_tree(self) -> Dict[str, str]:
        """Ensures the private CTG BAPTISMAL SYSTEM folder hierarchy exists and returns folder ID mappings."""
        root_id = str(self.config.get("root_folder_id") or "").strip()
        if not root_id:
            root_id = self.ensure_private_folder(DRIVE_ROOT_FOLDER_NAME)
        folder_map: Dict[str, str] = {"ROOT": root_id}
        for sub_name, years in DRIVE_SUBFOLDERS_TREE.items():
            sub_id = self.ensure_private_folder(sub_name, parent_id=root_id)
            folder_map[sub_name] = sub_id
            for yr in years:
                yr_id = self.ensure_private_folder(yr, parent_id=sub_id)
                folder_map[f"{sub_name}/{yr}"] = yr_id
        return folder_map

    def download_file_bytes(self, file_id: str) -> bytes:
        svc = self.get_service()
        request = svc.files().get_media(fileId=file_id, supportsAllDrives=True)
        return bytes(self._execute_with_retry(request.execute))

    def upload_private_file(
        self,
        file_bytes: bytes,
        filename: str,
        mime_type: str,
        folder_id: str,
    ) -> Dict[str, Any]:
        """Uploads a private file to a specific Google Drive folder. Never enables public link sharing."""
        from googleapiclient.http import MediaIoBaseUpload  # type: ignore

        svc = self.get_service()
        media = MediaIoBaseUpload(io.BytesIO(file_bytes), mimetype=mime_type, resumable=False)
        body = {"name": filename, "parents": [folder_id]}
        created = self._execute_with_retry(
            lambda: svc.files().create(body=body, media_body=media, fields="id, name, headRevisionId", supportsAllDrives=True).execute()
        )
        return dict(created)

    def update_existing_file_bytes(
        self,
        file_id: str,
        file_bytes: bytes,
        mime_type: str,
        expected_revision_id: str = "",
    ) -> Dict[str, Any]:
        """Updates an existing Google Drive file in-place after verifying canEdit and revision concurrency."""
        from googleapiclient.http import MediaIoBaseUpload  # type: ignore

        meta = self.verify_workbook_write_access(file_id)
        current_rev = str(meta.get("headRevisionId") or "")
        if expected_revision_id and current_rev and expected_revision_id != current_rev:
            raise RuntimeError(
                f"Concurrency conflict: workbook revision changed from '{expected_revision_id}' to '{current_rev}'. Please reload before saving."
            )
        svc = self.get_service()
        media = MediaIoBaseUpload(io.BytesIO(file_bytes), mimetype=mime_type, resumable=False)
        updated = self._execute_with_retry(
            lambda: svc.files().update(fileId=file_id, media_body=media, fields="id, name, headRevisionId", supportsAllDrives=True).execute()
        )
        return dict(updated)
