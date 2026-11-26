import io
import os
import tempfile
import unittest
import openpyxl
from baptismal.authentication import AuthService
from baptismal.authorization import AuthorizationService
from baptismal.config import HISTORICAL_27_COLUMNS
from baptismal.database import DatabaseManager
from baptismal.historical_records import HistoricalRecordsService
from baptismal.image_matching import ImageMatchingService
from baptismal.migrations import run_migrations
from baptismal.registration import RegistrationService


class TestRegistrationAndHistoricalSearch(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.NamedTemporaryFile(suffix=".db", delete=False)
        self.tmp.close()
        self.db = DatabaseManager(f"sqlite:///{self.tmp.name}")
        run_migrations(self.db)
        self.auth = AuthService(self.db)
        self.authz = AuthorizationService(self.db)
        self.admin_id = self.auth.create_account("admin_reg", "Admin", "Master Admin", plain_password="AdminPass!2026")
        self.hist = HistoricalRecordsService(self.db)
        self.reg = RegistrationService(self.db, self.authz)
        self.matcher = ImageMatchingService(self.db, self.authz)

    def tearDown(self):
        if os.path.exists(self.tmp.name):
            os.unlink(self.tmp.name)

    def test_historical_27col_import_search_and_image_mismatch_flag(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        ws.append(HISTORICAL_27_COLUMNS)
        row_data = {col: f"Val_{i}" for i, col in enumerate(HISTORICAL_27_COLUMNS, 1)}
        row_data["Name of the Child"] = "Synthetic Child One"
        row_data["Book Number"] = "2"
        row_data["Page Number"] = "743"
        row_data["Parish"] = "SYNTHETIC PARISH"
        row_data["Year of Baptism"] = "1964"
        ws.append([row_data[c] for c in HISTORICAL_27_COLUMNS])
        buf = io.BytesIO()
        wb.save(buf)

        cnt = self.hist.import_sheet1_from_excel_bytes(buf.getvalue())
        self.assertEqual(cnt, 1)

        found, total = self.hist.search_records(keyword="Synthetic Child")
        self.assertEqual(total, 1)
        self.assertEqual(found[0]["Page Number"], "743")

        # Image filename suggests page 744 while record has 743 -> MUST require Master Admin confirmation
        match_res = self.matcher.suggest_possible_match(found[0]["Record ID"], "BOOK 2", "Synthetic Child BC_744.jpeg")
        self.assertTrue(match_res["page_mismatch"])
        self.assertEqual(match_res["match_status"], "POSSIBLE_MATCH_NEEDS_CONFIRMATION")

        new_st = self.matcher.confirm_image_association(self.admin_id, match_res["association_id"], approve=True)
        self.assertEqual(new_st, "CONFIRMED")

    def test_new_registration_and_duplicate_prevention(self):
        payload = {c: "" for c in HISTORICAL_27_COLUMNS}
        payload.update({
            "Parish": "TEST PARISH",
            "Name of the Child": "Juan Dela Cruz Jr.",
            "Date of Baptism": "2026-05-10",
            "Year of Baptism": "2026",
            "Name of the Father": "Juan Dela Cruz Sr.",
            "Name of the Mother": "Maria Santos",
        })
        res = self.reg.register_new_baptism(self.admin_id, payload, sync_to_drive_excel=False)
        self.assertTrue(res["record_id"].startswith("BAP-2026-"))
        self.assertEqual(res["worksheet_name"], "NEW BAPTISMS - 2026")

        with self.assertRaises(ValueError):
            self.reg.register_new_baptism(self.admin_id, payload, sync_to_drive_excel=False)


if __name__ == "__main__":
    unittest.main()
