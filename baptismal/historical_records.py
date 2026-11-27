"""
baptismal/historical_records.py — Master-Admin-Only Read-Only Historical Records Engine
======================================================================================
- Loads `Historical Records` or `Sheet1` from `OVERALL BAPTISMAL CERTIFICATE LIST.xlsx`
  preserving all 27 original column names, spelling, and historical values verbatim.
- Supports searching by Record ID (`HIST-00743`) OR by typing ONLY the last digits/number
  (e.g. `743` -> `HIST-00743`), as well as any of the 27 columns.
- Excludes Hong Kong Parish records from operational search and dashboard statistics
  by default (without deleting any data from the source workbook).
"""

from datetime import date, datetime
import io
import json
import os
from typing import Any, Dict, List, Optional, Tuple
import openpyxl

from .authorization import AuthorizationService
from .config import DEFAULT_BAPTISMAL_WORKBOOK_FILE_ID, HISTORICAL_27_COLUMNS
from .database import DatabaseManager


def _format_cell_verbatim(val: Any) -> str:
    if val is None:
        return ""
    if isinstance(val, datetime):
        return val.strftime("%B %d, %Y")
    if isinstance(val, date):
        return val.strftime("%B %d, %Y")
    return str(val).strip()


def resolve_quick_record_id(number_or_id: str) -> str:
    """Converts a short number like '1', '25', or '743' into 'HIST-00001', 'HIST-00025', 'HIST-00743'."""
    s = str(number_or_id or "").strip().upper()
    if not s:
        return ""
    if s.startswith("HIST-"):
        tail = s[5:].strip()
        if tail.isdigit():
            return f"HIST-{int(tail):05d}"
        return s
    if s.isdigit():
        return f"HIST-{int(s):05d}"
    return s


class HistoricalRecordsService:
    """Manages ingestion and Master-Admin-only search over the 4,093 historical records."""

    def __init__(self, db: DatabaseManager, authz: Optional[AuthorizationService] = None):
        self.db = db
        self.authz = authz or AuthorizationService(db)

    def import_sheet1_from_excel_bytes(self, excel_bytes: bytes) -> int:
        wb = openpyxl.load_workbook(io.BytesIO(excel_bytes), read_only=True, data_only=True)
        if "Historical Records" in wb.sheetnames:
            ws = wb["Historical Records"]
        elif "Sheet1" in wb.sheetnames:
            ws = wb["Sheet1"]
        else:
            ws = wb.active

        rows_iter = ws.iter_rows(values_only=True)
        raw_headers = next(rows_iter, None)
        if not raw_headers:
            return 0
        headers = [str(h).strip() if h is not None else f"Col_{i}" for i, h in enumerate(raw_headers)]
        col_map = {h: idx for idx, h in enumerate(headers)}

        batch_params = []
        for row_num, row_vals in enumerate(rows_iter, start=2):
            row_dict: Dict[str, str] = {}
            has_any_val = False
            for col_name in HISTORICAL_27_COLUMNS:
                c_idx = col_map.get(col_name, -1)
                raw_v = row_vals[c_idx] if (0 <= c_idx < len(row_vals)) else ""
                val_str = _format_cell_verbatim(raw_v)
                if val_str:
                    has_any_val = True
                row_dict[col_name] = val_str
            if not has_any_val:
                continue

            rec_id = f"HIST-{row_num - 1:05d}"
            batch_params.append((
                rec_id,
                row_num,
                row_dict.get("Parish", ""),
                row_dict.get("Address", ""),
                row_dict.get("Province", ""),
                row_dict.get("Name of the Child", ""),
                row_dict.get("Date of Baptism", ""),
                row_dict.get("Place of Birth", ""),
                row_dict.get("Filiation/Sex", ""),
                row_dict.get("Nationality", ""),
                row_dict.get("Date of Birth", ""),
                row_dict.get("Name of the Father", ""),
                row_dict.get("Father's Residence", ""),
                row_dict.get("Name of the Mother", ""),
                row_dict.get("Mother's Residence", ""),
                row_dict.get("Paternal Grandparents", ""),
                row_dict.get("Maternal Grandparents", ""),
                row_dict.get("Vecindad", ""),
                row_dict.get("Sponsors", ""),
                row_dict.get("Civil Status", ""),
                row_dict.get("Residences", ""),
                row_dict.get("Officiating Priest", ""),
                row_dict.get("Place of Baptism", ""),
                row_dict.get("Day of Baptism", ""),
                row_dict.get("Month of Baptism", ""),
                row_dict.get("Year of Baptism", ""),
                row_dict.get("Book Number", ""),
                row_dict.get("Page Number", ""),
                row_dict.get("Cura Parocco", ""),
                json.dumps(row_dict, ensure_ascii=False),
                "historical-import",
            ))

        sql = """
            INSERT OR IGNORE INTO historical_records (
                record_id, source_row_number, parish, address, province, child_name,
                date_of_baptism, place_of_birth, filiation_sex, nationality, date_of_birth,
                father_name, fathers_residence, mother_name, mothers_residence,
                paternal_grandparents, maternal_grandparents, vecindad, sponsors,
                civil_status, residences, officiating_priest, place_of_baptism,
                day_of_baptism, month_of_baptism, year_of_baptism, book_number,
                page_number, cura_parocco, raw_json, created_at
            ) VALUES (
                %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,
                %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s
            )
        """
        if self.db.is_postgres:
            sql = sql.replace("INSERT OR IGNORE", "INSERT") + " ON CONFLICT (record_id) DO NOTHING"
        if batch_params:
            self.db.executemany(sql, batch_params)
        tot_now = int((self.db.fetch_one("SELECT COUNT(*) AS c FROM historical_records") or {}).get("c", 0))
        return tot_now

    def auto_sync_from_drive_or_disk(self, drive_mgr: Any = None, force: bool = False) -> Tuple[bool, int, str]:
        cur_cnt = int((self.db.fetch_one("SELECT COUNT(*) AS c FROM historical_records") or {}).get("c", 0))
        if cur_cnt > 0 and not force:
            return True, cur_cnt, f"Loaded {cur_cnt} historical records."

        local_candidates = [
            "OVERALL BAPTISMAL CERTIFICATE LIST.xlsx",
            "OVERALL_BAPTISMAL_CERTIFICATE_LIST.xlsx",
            "BAPTISMAL RELATED FILES/OVERALL BAPTISMAL CERTIFICATE LIST.xlsx",
            "downloads/OVERALL_BAPTISMAL_CERTIFICATE_LIST.xlsx",
        ]
        for lp in local_candidates:
            if os.path.exists(lp):
                try:
                    with open(lp, "rb") as f:
                        b = f.read()
                    cnt = self.import_sheet1_from_excel_bytes(b)
                    if cnt > 0:
                        return True, cnt, f"Na-load ang {cnt} historical records mula sa local '{lp}'."
                except Exception:
                    pass

        if drive_mgr is not None:
            wb_id = str(drive_mgr.config.get("workbook_file_id") or DEFAULT_BAPTISMAL_WORKBOOK_FILE_ID).strip()
            try:
                wb_bytes = drive_mgr.download_file_bytes(wb_id)
                if wb_bytes:
                    cnt = self.import_sheet1_from_excel_bytes(wb_bytes)
                    return True, cnt, f"Na-sync ang {cnt} historical records mula sa Google Drive (`OVERALL BAPTISMAL CERTIFICATE LIST.xlsx`)."
            except Exception as ex:
                return False, cur_cnt, f"Hindi ma-download mula sa Google Drive ({wb_id}): {ex}"

        return False, cur_cnt, "Walang nakitang local o Google Drive historical workbook."

    def search_records(
        self,
        keyword: str = "",
        filters: Optional[Dict[str, str]] = None,
        page: int = 1,
        page_size: int = 50,
        actor_account_id: str = "",
        include_excluded_source_records: bool = False,
        record_id_query: str = "",
    ) -> Tuple[List[Dict[str, Any]], int]:
        """Master-Admin-only case-insensitive search across Record ID (including short last-digits like '743' -> 'HIST-00743')
        and all 27 historical columns. Excludes Hong Kong Parish records from operational scope by default."""
        if actor_account_id:
            self.authz.require_master_admin(actor_account_id, action="search_historical_archive")

        where_clauses: List[str] = ["1=1"]
        params: List[Any] = []

        if not include_excluded_source_records:
            where_clauses.append("LOWER(COALESCE(parish, '')) NOT LIKE %s")
            params.append("%hong kong%")
            where_clauses.append("LOWER(COALESCE(parish, '')) NOT LIKE %s")
            params.append("%hongkong%")

        rid_q = str(record_id_query or "").strip()
        if rid_q:
            resolved_rid = resolve_quick_record_id(rid_q).lower()
            where_clauses.append("(LOWER(record_id) = %s OR LOWER(record_id) LIKE %s)")
            params.extend([resolved_rid, f"%{rid_q.lower()}%"])

        kw = str(keyword or "").strip().lower()
        if kw:
            resolved_kw_id = resolve_quick_record_id(kw).lower()
            like_kw = f"%{kw}%"
            where_clauses.append("(LOWER(raw_json) LIKE %s OR LOWER(record_id) LIKE %s OR LOWER(record_id) = %s)")
            params.extend([like_kw, like_kw, resolved_kw_id])

        field_to_col = {
            "Record ID": "record_id",
            "Parish": "parish",
            "Address": "address",
            "Province": "province",
            "Name of the Child": "child_name",
            "Date of Baptism": "date_of_baptism",
            "Place of Birth": "place_of_birth",
            "Filiation/Sex": "filiation_sex",
            "Nationality": "nationality",
            "Date of Birth": "date_of_birth",
            "Name of the Father": "father_name",
            "Father's Residence": "fathers_residence",
            "Name of the Mother": "mother_name",
            "Mother's Residence": "mothers_residence",
            "Paternal Grandparents": "paternal_grandparents",
            "Maternal Grandparents": "maternal_grandparents",
            "Vecindad": "vecindad",
            "Sponsors": "sponsors",
            "Civil Status": "civil_status",
            "Residences": "residences",
            "Officiating Priest": "officiating_priest",
            "Place of Baptism": "place_of_baptism",
            "Day of Baptism": "day_of_baptism",
            "Month of Baptism": "month_of_baptism",
            "Year of Baptism": "year_of_baptism",
            "Book Number": "book_number",
            "Page Number": "page_number",
            "Cura Parocco": "cura_parocco",
        }
        for f_key, f_val in (filters or {}).items():
            fv = str(f_val or "").strip().lower()
            if fv and f_key in field_to_col:
                col_n = field_to_col[f_key]
                where_clauses.append(f"LOWER(COALESCE({col_n}, '')) LIKE %s")
                params.append(f"%{fv}%")

        where_sql = " AND ".join(where_clauses)
        cnt_row = self.db.fetch_one(f"SELECT COUNT(*) AS cnt FROM historical_records WHERE {where_sql}", tuple(params))
        total_cnt = int((cnt_row or {}).get("cnt", 0))

        offset = max(0, (int(page) - 1) * int(page_size))
        rows = self.db.fetch_all(
            f"""
            SELECT record_id, source_row_number, raw_json
            FROM historical_records
            WHERE {where_sql}
            ORDER BY source_row_number ASC
            LIMIT %s OFFSET %s
            """,
            tuple(params + [int(page_size), offset]),
        )
        results: List[Dict[str, Any]] = []
        for r in rows:
            parsed = json.loads(r["raw_json"])
            parsed["Record ID"] = r["record_id"]
            parsed["_source_row"] = r["source_row_number"]
            results.append(parsed)
        return results, total_cnt

    def get_record_by_id(self, record_id: str, actor_account_id: str = "") -> Optional[Dict[str, Any]]:
        resolved_id = resolve_quick_record_id(record_id)
        if actor_account_id:
            self.authz.require_record_permission(actor_account_id, resolved_id, "view_record")
        row = self.db.fetch_one(
            "SELECT record_id, source_row_number, raw_json FROM historical_records WHERE record_id = %s",
            (resolved_id,),
        )
        if row:
            parsed = json.loads(row["raw_json"])
            parsed["Record ID"] = row["record_id"]
            parsed["_source_row"] = row["source_row_number"]
            return parsed
        new_row = self.db.fetch_one(
            "SELECT record_id, payload_json FROM new_baptism_registrations WHERE record_id = %s",
            (resolved_id,),
        )
        if new_row:
            parsed = json.loads(new_row["payload_json"])
            parsed["Record ID"] = new_row["record_id"]
            return parsed
        return None

    def get_dashboard_statistics(self, actor_account_id: str = "") -> Dict[str, Any]:
        """Master-Admin-only global statistics across the 9 official parishes (excluding Hong Kong)."""
        if actor_account_id:
            self.authz.require_master_admin(actor_account_id, action="view_master_dashboard_statistics")

        hk_excl = "LOWER(COALESCE(parish, '')) NOT LIKE '%hong kong%' AND LOWER(COALESCE(parish, '')) NOT LIKE '%hongkong%'"
        tot_hist = int((self.db.fetch_one(f"SELECT COUNT(*) AS c FROM historical_records WHERE {hk_excl}") or {}).get("c", 0))
        by_book = self.db.fetch_all(
            f"SELECT COALESCE(NULLIF(book_number, ''), 'Unspecified') AS book_number, COUNT(*) AS cnt FROM historical_records WHERE {hk_excl} GROUP BY book_number ORDER BY book_number"
        )
        by_parish = self.db.fetch_all(
            f"SELECT COALESCE(NULLIF(parish, ''), 'Unspecified') AS parish, COUNT(*) AS cnt FROM historical_records WHERE {hk_excl} GROUP BY parish ORDER BY cnt DESC"
        )
        by_year = self.db.fetch_all(
            f"SELECT COALESCE(NULLIF(year_of_baptism, ''), 'Unspecified') AS year_of_baptism, COUNT(*) AS cnt FROM historical_records WHERE {hk_excl} GROUP BY year_of_baptism ORDER BY year_of_baptism DESC"
        )
        matched_imgs = int((self.db.fetch_one("SELECT COUNT(*) AS c FROM historical_image_associations WHERE match_status = 'CONFIRMED'") or {}).get("c", 0))
        unmatched_imgs = int((self.db.fetch_one("SELECT COUNT(*) AS c FROM historical_image_associations WHERE match_status != 'CONFIRMED'") or {}).get("c", 0))
        new_regs = int((self.db.fetch_one(f"SELECT COUNT(*) AS c FROM new_baptism_registrations WHERE {hk_excl}") or {}).get("c", 0))
        gen_certs = int((self.db.fetch_one("SELECT COUNT(*) AS c FROM generated_certificates") or {}).get("c", 0))
        pending_reqs = int((self.db.fetch_one("SELECT COUNT(*) AS c FROM certificate_requests WHERE UPPER(status) IN ('PENDING', 'UNDER REVIEW', 'ADDITIONAL INFORMATION REQUIRED')") or {}).get("c", 0))

        return {
            "total_historical": tot_hist,
            "by_book": by_book,
            "by_parish": by_parish,
            "by_year": by_year,
            "matched_images": matched_imgs,
            "unmatched_images": unmatched_imgs,
            "new_registrations": new_regs,
            "generated_certificates": gen_certs,
            "pending_requests": pending_reqs,
        }
