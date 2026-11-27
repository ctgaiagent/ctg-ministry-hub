"""
baptismal/google_drive.py — Least-Privilege Private Google Drive Integration
============================================================================
- Reuses Google OAuth Refresh Token or Service Account credentials from st.secrets / env / CTG Ministry Hub.
- Verifies Editor permission (`capabilities.canEdit`) on the designated Excel workbook BEFORE writing.
- Lists and downloads historical records (`OVERALL BAPTISMAL CERTIFICATE LIST.xlsx`) and
  `BOOK 1`–`BOOK 9` images directly from the `CTG BAPTISMAL RECORDS` Google Drive folder.
- Never enables public sharing ("Anyone with the link").
"""

import io
import json
import time
from typing import Any, Dict, List, Optional

from .config import (
    DEFAULT_BAPTISMAL_ROOT_FOLDER_ID,
    DEFAULT_BAPTISMAL_WORKBOOK_FILE_ID,
    DRIVE_ROOT_FOLDER_NAME,
    DRIVE_SUBFOLDERS_TREE,
    KNOWN_BAPTISMAL_SUBFOLDER_IDS,
    get_google_drive_config,
)


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

        # Try reusing active Google Drive service from main CTG Ministry Hub app if running in unified mode
        try:
            import sys
            for mod_name in ("__main__", "app"):
                mod = sys.modules.get(mod_name)
                if mod and hasattr(mod, "get_all_candidate_google_drive_services"):
                    svcs = mod.get_all_candidate_google_drive_services()
                    if svcs:
                        self._service = svcs[0]
                        return self._service
        except Exception:
            pass

        raise GoogleDrivePermissionError(
            "Google Drive credentials are not configured. Set Google Drive OAuth or Service Account credentials in App Settings / Streamlit Secrets."
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

    def verify_workbook_write_access(self, file_id: str = "") -> Dict[str, Any]:
        """Verifies that the backend OAuth/Service Account has Editor permission on the designated workbook."""
        target_id = str(file_id or self.config.get("workbook_file_id") or DEFAULT_BAPTISMAL_WORKBOOK_FILE_ID).strip()
        if not target_id:
            raise GoogleDrivePermissionError("BAPTISMAL_WORKBOOK_FILE_ID is not configured.")
        svc = self.get_service()
        meta = self._execute_with_retry(
            lambda: svc.files()
            .get(
                fileId=target_id,
                fields="id, name, mimeType, capabilities(canEdit, canDownload), headRevisionId",
                supportsAllDrives=True,
            )
            .execute()
        )
        caps = meta.get("capabilities") or {}
        if not caps.get("canEdit", False):
            raise GoogleDrivePermissionError(
                f"Insufficient Google Drive permissions on workbook '{meta.get('name', target_id)}': backend account is Viewer-only (canEdit=False). Write operation aborted."
            )
        return meta

    def ensure_private_folder(self, folder_name: str, parent_id: str = "") -> str:
        """Finds or creates a private folder without any public 'anyone' permission."""
        if folder_name in KNOWN_BAPTISMAL_SUBFOLDER_IDS and parent_id in ("", DEFAULT_BAPTISMAL_ROOT_FOLDER_ID):
            return KNOWN_BAPTISMAL_SUBFOLDER_IDS[folder_name]
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
        """Ensures the private CTG BAPTISMAL RECORDS folder hierarchy exists and returns folder ID mappings."""
        root_id = str(self.config.get("root_folder_id") or DEFAULT_BAPTISMAL_ROOT_FOLDER_ID).strip()
        if not root_id:
            root_id = self.ensure_private_folder(DRIVE_ROOT_FOLDER_NAME)
        folder_map: Dict[str, str] = {"ROOT": root_id, **KNOWN_BAPTISMAL_SUBFOLDER_IDS}
        for sub_name, years in DRIVE_SUBFOLDERS_TREE.items():
            sub_id = KNOWN_BAPTISMAL_SUBFOLDER_IDS.get(sub_name) or self.ensure_private_folder(sub_name, parent_id=root_id)
            folder_map[sub_name] = sub_id
            for yr in years:
                yr_id = self.ensure_private_folder(yr, parent_id=sub_id)
                folder_map[f"{sub_name}/{yr}"] = yr_id
        return folder_map

    def list_files_in_folder(self, folder_id: str, page_size: int = 1000) -> List[Dict[str, Any]]:
        """Lists all files inside a specific Google Drive folder (such as BOOK 1..9) with automatic pagination."""
        if not folder_id:
            return []
        svc = self.get_service()
        q = f"'{folder_id}' in parents and trashed = false"
        all_files: List[Dict[str, Any]] = []
        page_token: Optional[str] = None
        max_total = max(int(page_size), 1000)
        while True:
            req_kwargs: Dict[str, Any] = {
                "q": q,
                "pageSize": min(1000, max_total),
                "fields": "nextPageToken, files(id, name, mimeType, size, modifiedTime)",
                "supportsAllDrives": True,
                "includeItemsFromAllDrives": True,
            }
            if page_token:
                req_kwargs["pageToken"] = page_token
            res = self._execute_with_retry(lambda: svc.files().list(**req_kwargs).execute())
            batch = list(res.get("files") or [])
            all_files.extend(batch)
            page_token = res.get("nextPageToken")
            if not page_token or len(all_files) >= max_total:
                break
        return all_files

    def discover_book_folders_in_drive(self) -> Dict[str, str]:
        """Discovers BOOK 1..BOOK 9 folder IDs inside the CTG BAPTISMAL RECORDS root folder."""
        book_map = {k: v for k, v in KNOWN_BAPTISMAL_SUBFOLDER_IDS.items() if k.startswith("BOOK ")}
        root_id = str(self.config.get("root_folder_id") or DEFAULT_BAPTISMAL_ROOT_FOLDER_ID).strip()
        try:
            children = self.list_files_in_folder(root_id, page_size=100)
            for item in children:
                nm = str(item.get("name") or "").strip().upper()
                if nm.startswith("BOOK ") and item.get("mimeType") == "application/vnd.google-apps.folder":
                    book_map[nm] = str(item["id"])
        except Exception:
            pass
        return book_map

    def download_file_bytes(self, file_id: str = "") -> bytes:
        target_id = str(file_id or self.config.get("workbook_file_id") or DEFAULT_BAPTISMAL_WORKBOOK_FILE_ID).strip()
        svc = self.get_service()
        request = svc.files().get_media(fileId=target_id, supportsAllDrives=True)
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
