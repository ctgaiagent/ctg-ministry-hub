import io
import unittest
import openpyxl
from baptismal.config import HISTORICAL_27_COLUMNS
from baptismal.excel_manager import ExcelSafetyError, append_new_baptism_to_workbook_bytes


class TestExcelManager(unittest.TestCase):
    def _make_base_workbook(self) -> bytes:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        ws.append(HISTORICAL_27_COLUMNS)
        ws.append([f"Hist_{i}" for i in range(1, 28)])
        buf = io.BytesIO()
        wb.save(buf)
        return buf.getvalue()

    def test_create_and_append_yearly_worksheet_preserving_sheet1(self):
        base_bytes = self._make_base_workbook()
        rec1 = {c: "" for c in HISTORICAL_27_COLUMNS}
        rec1["Name of the Child"] = "Synthetic Baby 1"
        rec1["Year of Baptism"] = "2026"

        wb_bytes_1, ws_name_1 = append_new_baptism_to_workbook_bytes(base_bytes, rec1, "BAP-2026-0001", "2026")
        self.assertEqual(ws_name_1, "NEW BAPTISMS - 2026")

        wb1 = openpyxl.load_workbook(io.BytesIO(wb_bytes_1))
        self.assertEqual(wb1["Sheet1"].max_row, 2)
        self.assertEqual(wb1["NEW BAPTISMS - 2026"].max_row, 2)

        rec2 = dict(rec1)
        rec2["Name of the Child"] = "Synthetic Baby 2"
        wb_bytes_2, _ = append_new_baptism_to_workbook_bytes(wb_bytes_1, rec2, "BAP-2026-0002", "2026")
        wb2 = openpyxl.load_workbook(io.BytesIO(wb_bytes_2))
        self.assertEqual(wb2["Sheet1"].max_row, 2)
        self.assertEqual(wb2["NEW BAPTISMS - 2026"].max_row, 3)

        # Duplicate Record ID must be blocked
        with self.assertRaises(ExcelSafetyError):
            append_new_baptism_to_workbook_bytes(wb_bytes_2, rec2, "BAP-2026-0002", "2026")


if __name__ == "__main__":
    unittest.main()
