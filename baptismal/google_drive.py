"""
baptismal/google_drive.py — Least-Privilege Private Google Drive Integration
============================================================================
- Reuses all verified Google Drive OAuth Refresh Tokens (including Cloud Vault tokens) and Service Accounts
  from main CTG Ministry Hub (`app.get_all_candidate_google_drive_services()`) or `st.secrets` / env.
- Never forces `scopes=["https://www.googleapis.com/auth/drive"]` onto OAuth refresh token objects during refresh
  (which previously caused `invalid_scope: Bad Request` on tokens with different scope sets), and automatically
  rotates across all candidate Drive services if one token is Gmail-only.
- Verifies Editor permission (`capabilities.canEdit`) on the designated Excel workbook BEFORE writing.
- Never enables public sharing ("Anyone with the link").
"""

import io
import json
import sys
import time
from typing import Any, Callable, Dict, List, Optional

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
    """Backend-only Google Drive API v3 manager with multi-token rotation and strict permission checks."""

    DRIVE_SCOPES = ["https://www.googleapis.com/auth/drive"]

    def __init__(self, drive_service: Optional[Any] = None, config: Optional[Dict[str, Any]] = None):
        self.config = config or get_google_drive_config()
        self._injected_service = drive_service
        self._service = drive_service

    def _get_hub_app_module(self) -> Optional[Any]:
        for mod_name in ("__main__", "app"):
            mod = sys.modules.get(mod_name)
            if mod and hasattr(mod, "get_all_candidate_google_drive_services"):
                return mod
        return None

    def get_candidate_services(self) -> List[Any]:
        """Returns all candidate Google Drive service instances in priority order:
        1. Injected test/custom service (if provided).
        2. Main CTG Ministry Hub candidate services (Cloud Vault Drive token, verified working Drive RT, Service Accounts).
        3. Explicit Baptismal Service Account.
        4. Explicit Baptismal/Env OAuth Refresh Token (initialized WITHOUT `scopes=` to avoid `invalid_scope: Bad Request`).
        """
        if self._injected_service is not None:
            return [self._injected_service]

        candidates: List[Any] = []
        seen_ids = set()

        # 1. Main CTG Ministry Hub candidate services (includes Cloud Vault Drive token from Master Excel)
        hub_mod = self._get_hub_app_module()
        if hub_mod is not None:
            try:
                # Ensure vault token is extracted if Master Census Excel exists on disk
                if hasattr(hub_mod, "_extract_vault_token_from_excel_bytes"):
                    for xl_cand in ("CTG_HEADS_TOTAL_MASTER_ORDERED.xlsx", "CTG_HEADS_TOTAL_UPDATED FIXED.xlsx"):
                        if os.path.exists(xl_cand):
                            try:
                                with open(xl_cand, "rb") as xf:
                                    hub_mod._extract_vault_token_from_excel_bytes(xf.read())
                                break
                            except Exception:
                                pass
            except Exception:
                pass
            try:
                hub_svcs = hub_mod.get_all_candidate_google_drive_services() or []
                for s in hub_svcs:
                    if s is not None and id(s) not in seen_ids:
                        seen_ids.add(id(s))
                        candidates.append(s)
            except Exception:
                pass

        # 2. Explicit Service Account from config
        sa_raw = self.config.get("service_account")
        if sa_raw:
            try:
                from google.oauth2 import service_account  # type: ignore
                from googleapiclient.discovery import build  # type: ignore

                sa_info = sa_raw if isinstance(sa_raw, dict) else json.loads(str(sa_raw))
                creds = service_account.Credentials.from_service_account_info(sa_info, scopes=self.DRIVE_SCOPES)
                svc_sa = build("drive", "v3", credentials=creds, cache_discovery=False)
                svc_sa._ctg_rt_token = "SERVICE_ACCOUNT"
                candidates.append(svc_sa)
            except Exception:
                pass

        # 3. Explicit OAuth Refresh Token from config (NO `scopes=` argument so refresh never triggers invalid_scope!)
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
                )
                svc_oauth = build("drive", "v3", credentials=creds, cache_discovery=False)
                svc_oauth._ctg_rt_token = rt
                candidates.append(svc_oauth)
            except Exception:
                pass

        if not candidates:
            raise GoogleDrivePermissionError(
                "Google Drive credentials are not configured. Set Google Drive OAuth or Service Account credentials in App Settings / Streamlit Secrets."
            )
        return candidates

    def get_service(self):
        svcs = self.get_candidate_services()
        return svcs[0]

    def _run_across_candidate_services(self, op_Using_svc: Callable[[Any], Any], max_retries_per_svc: int = 2) -> Any:
        """Executes a Drive operation across all candidate services, automatically rotating past any token
        that fails with `invalid_scope`, `403`, `insufficientPermissions`, or `invalid_grant`."""
        svcs = self.get_candidate_services()
        hub_mod = self._get_hub_app_module()
        last_err: Optional[Exception] = None

        for svc in svcs:
            rt_tok = getattr(svc, "_ctg_rt_token", "")
            for attempt in range(max_retries_per_svc):
                try:
                    result = op_Using_svc(svc)
                    if hub_mod and rt_tok and rt_tok != "SERVICE_ACCOUNT" and hasattr(hub_mod, "_mark_token_as_verified_drive_working"):
                        try:
                            hub_mod._mark_token_as_verified_drive_working(rt_tok)
                        except Exception:
                            pass
                    return result
                except Exception as ex:
                    last_err = ex
                    err_s = str(ex).lower()
                    is_scope_or_auth_err = any(
                        k in err_s
                        for k in ("invalid_scope", "insufficient", "403", "invalid_grant", "unauthorized")
                    )
                    if is_scope_or_auth_err:
                        if hub_mod and rt_tok and rt_tok != "SERVICE_ACCOUNT" and hasattr(hub_mod, "_mark_token_as_403_gmail_only"):
                            try:
                                hub_mod._mark_token_as_403_gmail_only(rt_tok)
                            except Exception:
                                pass
                        break  # Immediately rotate to next candidate service
                    if "404" in err_s:
                        raise
                    if attempt < max_retries_per_svc - 1:
                        time.sleep(0.5 * (attempt + 1))

        if isinstance(last_err, GoogleDrivePermissionError):
            raise last_err
        err_msg = str(last_err or "Unknown Google Drive error")
        if "invalid_scope" in err_msg.lower() or "insufficient" in err_msg.lower() or "403" in err_msg.lower():
            raise GoogleDrivePermissionError(
                "Ang kasalukuyang OAuth Refresh Token sa Streamlit Secrets ay Gmail-only (walang Google Drive scope). "
                "Buksan ang '🛠️ App Settings & Security → 🔑 1-Click Google Drive + Gmail Scope Upgrade' sa kaliwang menu "
                "upang i-authorize ang Google Drive access."
            )
        raise RuntimeError(f"Google Drive API request failed: {err_msg}")

    def verify_workbook_write_access(self, file_id: str = "") -> Dict[str, Any]:
        """Verifies that the backend OAuth/Service Account has Editor permission on the designated workbook."""
        target_id = str(file_id or self.config.get("workbook_file_id") or DEFAULT_BAPTISMAL_WORKBOOK_FILE_ID).strip()
        if not target_id:
            raise GoogleDrivePermissionError("BAPTISMAL_WORKBOOK_FILE_ID is not configured.")
        meta = self._run_across_candidate_services(
            lambda svc: svc.files()
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
        safe_q_name = folder_name.replace("'", "\\'")
        q = f"name = '{safe_q_name}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
        if parent_id:
            q += f" and '{parent_id}' in parents"
        res = self._run_across_candidate_services(
            lambda svc: svc.files()
            .list(q=q, fields="files(id, name)", supportsAllDrives=True, includeItemsFromAllDrives=True)
            .execute()
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
        created = self._run_across_candidate_services(
            lambda svc: svc.files().create(body=body, fields="id", supportsAllDrives=True).execute()
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
        q = f"'{folder_id}' in parents and trashed = false"
        max_total = max(int(page_size), 1000)

        def _list_all_with_svc(svc: Any) -> List[Dict[str, Any]]:
            all_files: List[Dict[str, Any]] = []
            page_token: Optional[str] = None
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
                res = svc.files().list(**req_kwargs).execute()
                batch = list(res.get("files") or [])
                all_files.extend(batch)
                page_token = res.get("nextPageToken")
                if not page_token or len(all_files) >= max_total:
                    break
            return all_files

        return self._run_across_candidate_services(_list_all_with_svc)

    def search_files_in_folder(self, folder_id: str, search_terms: List[str], page_size: int = 50) -> List[Dict[str, Any]]:
        """Performs a fast targeted filename search inside a specific Google Drive Book folder."""
        if not folder_id:
            return []
        clean_terms = [str(t or "").strip().replace("'", "\\'") for t in search_terms if str(t or "").strip()]
        if not clean_terms:
            return self.list_files_in_folder(folder_id, page_size=page_size)
        or_clauses = " or ".join([f"name contains '{t}'" for t in clean_terms[:6]])
        q = f"'{folder_id}' in parents and trashed = false and ({or_clauses})"
        res = self._run_across_candidate_services(
            lambda svc: svc.files()
            .list(
                q=q,
                pageSize=min(200, int(page_size)),
                fields="files(id, name, mimeType, size, modifiedTime)",
                supportsAllDrives=True,
                includeItemsFromAllDrives=True,
            )
            .execute()
        )
        return list(res.get("files") or [])

    def discover_book_folders_in_drive(self) -> Dict[str, str]:
        """Returns BOOK 1..BOOK 9 folder IDs inside the CTG BAPTISMAL RECORDS root folder."""
        book_map = {k: v for k, v in KNOWN_BAPTISMAL_SUBFOLDER_IDS.items() if k.startswith("BOOK ")}
        if len(book_map) >= 9:
            return book_map
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
        try:
            data = self._run_across_candidate_services(
                lambda svc: svc.files().get_media(fileId=target_id, supportsAllDrives=True).execute()
            )
            return bytes(data)
        except Exception as primary_ex:
            # Direct HTTPS fallback if file is accessible by link
            try:
                import requests
                for url in (
                    f"https://drive.google.com/uc?export=download&id={target_id}",
                    f"https://docs.google.com/spreadsheets/d/{target_id}/export?format=xlsx",
                ):
                    resp = requests.get(url, timeout=20)
                    if resp.status_code == 200 and resp.content and not resp.content[:100].lstrip().startswith(b"<!DOCTYPE html"):
                        return bytes(resp.content)
            except Exception:
                pass
            raise primary_ex

    def upload_private_file(
        self,
        file_bytes: bytes,
        filename: str,
        mime_type: str,
        folder_id: str,
    ) -> Dict[str, Any]:
        """Uploads a private file to a specific Google Drive folder. Never enables public link sharing."""
        from googleapiclient.http import MediaIoBaseUpload  # type: ignore

        def _do_upload(svc: Any) -> Dict[str, Any]:
            media = MediaIoBaseUpload(io.BytesIO(file_bytes), mimetype=mime_type, resumable=False)
            body = {"name": filename, "parents": [folder_id]}
            return svc.files().create(body=body, media_body=media, fields="id, name, headRevisionId", supportsAllDrives=True).execute()

        created = self._run_across_candidate_services(_do_upload)
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

        def _do_update(svc: Any) -> Dict[str, Any]:
            media = MediaIoBaseUpload(io.BytesIO(file_bytes), mimetype=mime_type, resumable=False)
            return svc.files().update(fileId=file_id, media_body=media, fields="id, name, headRevisionId", supportsAllDrives=True).execute()

        updated = self._run_across_candidate_services(_do_update)
        return dict(updated)
