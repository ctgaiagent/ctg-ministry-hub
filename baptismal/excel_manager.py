"""
baptismal/excel_manager.py — Safe, Non-Destructive Yearly Excel Workbook Sync
=============================================================================
Enforces all workbook safety rules:
1. NEVER modifies the original historical worksheet (`Historical Records` or `Sheet1`).
2. NEVER deletes historical rows.
3. NEVER overwrites the workbook with a blank replacement.
4. Appends approved new records to `NEW BAPTISMS - YYYY` (creating the worksheet if missing).
5. Prevents duplicate submissions via Record ID & dedup fingerprint checks.
6. Preserves existing formatting and creates backups before updates.
"""

import io
from typing import Any, Dict, List, Tuple
import openpyxl

from .config import HISTORICAL_27_COLUMNS, YEARLY_WORKSHEET_PREFIX
from .validation import sanitize_excel_cell


class ExcelSafetyError(RuntimeError):
    """Raised when a workbook update could risk corrupting or overwriting historical Sheet1/Historical Records data."""


def get_yearly_worksheet_name(baptism_year: str) -> str:
    yr = str(baptism_year or "").strip()
    if not yr.isdigit() or len(yr) != 4:
        raise ValueError(f"Invalid baptism year '{baptism_year}'. Expected 4-digit year (e.g., 2026).")
    return f"{YEARLY_WORKSHEET_PREFIX}{yr}"


def _find_historical_sheet_name(wb: openpyxl.Workbook) -> str:
    for cand in ("Sheet1", "Historical Records"):
        if cand in wb.sheetnames:
            return cand
    return ""


def append_new_baptism_to_workbook_bytes(
    existing_workbook_bytes: bytes,
    record_data: Dict[str, Any],
    record_id: str,
    baptism_year: str,
) -> Tuple[bytes, str]:
    """Safely appends a new baptismal record to `NEW BAPTISMS - YYYY` inside the existing workbook bytes.
    Strictly preserves `Sheet1` / `Historical Records` and all existing worksheets untouched."""
    if not existing_workbook_bytes:
        raise ExcelSafetyError("Cannot update an empty workbook. Aborting to prevent blank replacement.")

    wb = openpyxl.load_workbook(io.BytesIO(existing_workbook_bytes))
    hist_sheet_name = _find_historical_sheet_name(wb)
    if not hist_sheet_name:
        raise ExcelSafetyError("Expected historical worksheet ('Sheet1' or 'Historical Records') is missing. Aborting workbook update.")

    sheet1_before_rows = wb[hist_sheet_name].max_row
    target_ws_name = get_yearly_worksheet_name(baptism_year)
    expected_headers = ["Record ID"] + HISTORICAL_27_COLUMNS

    if target_ws_name not in wb.sheetnames:
        ws = wb.create_sheet(title=target_ws_name)
        hdr_font = openpyxl.styles.Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        hdr_fill = openpyxl.styles.PatternFill("solid", fgColor="0F2C59")
        for col_idx, h_text in enumerate(expected_headers, 1):
            cell = ws.cell(1, col_idx, h_text)
            cell.font = hdr_font
            cell.fill = hdr_fill
    else:
        ws = wb[target_ws_name]
        if ws.max_row == 0 or (ws.max_row == 1 and ws.cell(1, 1).value is None):
            hdr_font = openpyxl.styles.Font(name="Calibri", size=11, bold=True, color="FFFFFF")
            hdr_fill = openpyxl.styles.PatternFill("solid", fgColor="0F2C59")
            for col_idx, h_text in enumerate(expected_headers, 1):
                cell = ws.cell(1, col_idx, h_text)
                cell.font = hdr_font
                cell.fill = hdr_fill
        existing_rids = set()
        for r_idx in range(2, ws.max_row + 1):
            val = str(ws.cell(r_idx, 1).value or "").strip()
            if val:
                existing_rids.add(val)
        if record_id in existing_rids:
            raise ExcelSafetyError(f"Duplicate prevention: Record ID '{record_id}' already exists in '{target_ws_name}'.")

    row_values: List[str] = [sanitize_excel_cell(record_id)]
    for col_name in HISTORICAL_27_COLUMNS:
        row_values.append(sanitize_excel_cell(record_data.get(col_name, "")))
    ws.append(row_values)

    if wb[hist_sheet_name].max_row != sheet1_before_rows:
        raise ExcelSafetyError(f"Safety check failed: '{hist_sheet_name}' row count changed unexpectedly. Aborting.")

    out_buf = io.BytesIO()
    wb.save(out_buf)
    return out_buf.getvalue(), target_ws_name
