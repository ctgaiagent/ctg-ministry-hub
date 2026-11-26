"""
baptismal/historical_records.py — Read-Only Historical Records Engine (~4,093 Records)
======================================================================================
- Loads `Sheet1` from `OVERALL BAPTISMAL CERTIFICATE LIST.xlsx` (or CSV) preserving
  all 27 original column names, spelling, and historical values verbatim.
- Historical records are strictly read-only by default.
- Provides case-insensitive, partial-keyword search across all 27 columns,
  multiple filters, pagination, and dashboard statistics.
"""

import io
import json
from typing import Any, Dict, List, Optional, Tuple
import openpyxl
import pandas as pd

from .config import HISTORICAL_27_COLUMNS
from .database import DatabaseManager


class HistoricalRecordsService:
    """Manages ingestion and read-only multi-column search over the 4,093 historical records."""

    def __init__(self, db: DatabaseManager):
        self.db = db

    def import_sheet1_from_excel_bytes(self, excel_bytes: bytes) -> int:
        """Imports historical records from `Sheet1` idempotently without altering original spelling or values."""
        wb = openpyxl.load_workbook(io.BytesIO(excel_bytes), read_only=True, data_only=True)
        ws = wb["Sheet1"] if "Sheet1" in wb.sheetnames else wb.active
        rows_iter = ws.iter_rows(values_only=True)
        raw_headers = next(rows_iter, None)
        if not raw_headers:
            return 0
        headers = [str(h).strip() if h is not None else f"Col_{i}" for i, h in enumerate(raw_headers)]
        col_map = {h: idx for idx, h in enumerate(headers)}

        inserted_count = 0
        batch_params = []
        for row_num, row_vals in enumerate(rows_iter, start=2):
            row_dict: Dict[str, str] = {}
            has_any_val = False
            for col_name in HISTORICAL_27_COLUMNS:
                c_idx = col_map.get(col_name, -1)
                raw_v = row_vals[c_idx] if (0 <= c_idx < len(row_vals)) else ""
                val_str = "" if raw_v is None else str(raw_v).strip()
                if val_str:
                    has_any_val = True
                row_dict[col_name] = val_str
            if not has_any_val:
                continue

            rec_id = f"HIST-{row_num:05d}"
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
            inserted_count = self.db.executemany(sql, batch_params)
        return inserted_count

    def search_records(
        self,
        keyword: str = "",
        filters: Optional[Dict[str, str]] = None,
        page: int = 1,
        page_size: int = 50,
    ) -> Tuple[List[Dict[str, Any]], int]:
        """Case-insensitive, partial-keyword search across ALL 27 historical columns with pagination."""
        where_clauses: List[str] = ["1=1"]
        params: List[Any] = []

        kw = str(keyword or "").strip().lower()
        if kw:
            like_kw = f"%{kw}%"
            where_clauses.append("LOWER(raw_json) LIKE %s")
            params.append(like_kw)

        field_to_col = {
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

    def get_record_by_id(self, record_id: str) -> Optional[Dict[str, Any]]:
        row = self.db.fetch_one(
            "SELECT record_id, source_row_number, raw_json FROM historical_records WHERE record_id = %s",
            (record_id,),
        )
        if row:
            parsed = json.loads(row["raw_json"])
            parsed["Record ID"] = row["record_id"]
            parsed["_source_row"] = row["source_row_number"]
            return parsed
        new_row = self.db.fetch_one(
            "SELECT record_id, payload_json FROM new_baptism_registrations WHERE record_id = %s",
            (record_id,),
        )
        if new_row:
            parsed = json.loads(new_row["payload_json"])
            parsed["Record ID"] = new_row["record_id"]
            return parsed
        return None

    def get_dashboard_statistics(self) -> Dict[str, Any]:
        tot_hist = int((self.db.fetch_one("SELECT COUNT(*) AS c FROM historical_records") or {}).get("c", 0))
        by_book = self.db.fetch_all(
            "SELECT COALESCE(NULLIF(book_number, ''), 'Unspecified') AS book_number, COUNT(*) AS cnt FROM historical_records GROUP BY book_number ORDER BY book_number"
        )
        by_parish = self.db.fetch_all(
            "SELECT COALESCE(NULLIF(parish, ''), 'Unspecified') AS parish, COUNT(*) AS cnt FROM historical_records GROUP BY parish ORDER BY cnt DESC"
        )
        by_year = self.db.fetch_all(
            "SELECT COALESCE(NULLIF(year_of_baptism, ''), 'Unspecified') AS year_of_baptism, COUNT(*) AS cnt FROM historical_records GROUP BY year_of_baptism ORDER BY year_of_baptism DESC"
        )
        matched_imgs = int((self.db.fetch_one("SELECT COUNT(*) AS c FROM historical_image_associations WHERE match_status = 'CONFIRMED'") or {}).get("c", 0))
        unmatched_imgs = int((self.db.fetch_one("SELECT COUNT(*) AS c FROM historical_image_associations WHERE match_status != 'CONFIRMED'") or {}).get("c", 0))
        new_regs = int((self.db.fetch_one("SELECT COUNT(*) AS c FROM new_baptism_registrations") or {}).get("c", 0))
        gen_certs = int((self.db.fetch_one("SELECT COUNT(*) AS c FROM generated_certificates") or {}).get("c", 0))
        pending_reqs = int((self.db.fetch_one("SELECT COUNT(*) AS c FROM certificate_requests WHERE status = 'PENDING'") or {}).get("c", 0))

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
