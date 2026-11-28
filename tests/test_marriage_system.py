"""
tests/test_marriage_system.py — Unit Tests for CTG Marriage Records & Certification System
==========================================================================================
Verifies:
1. Master-Admin-only RBAC enforcement (Parish Leader Admins are strictly blocked).
2. Quick Marriage Record ID resolution ('5' -> 'MAR-00005', '154' -> 'MAR-00154', 'MC_112' -> 'MAR-00112').
3. Verified Google Drive Marriage Certificate JPEG seed records & filename parser.
4. New Marriage Registration & Portrait/Landscape Official Marriage Certificate PDF generation.
"""

import os
import tempfile
import unittest

from baptismal.authorization import AuthorizationError
from baptismal.database import DatabaseManager
from baptismal.migrations import run_migrations
from marriage.certificates import (
    generate_and_archive_marriage_certificate,
    list_generated_marriage_certificates,
)
from marriage.config import (
    DEFAULT_MARRIAGE_IMAGES_FOLDER_ID,
    DEFAULT_MARRIAGE_ROOT_FOLDER_ID,
    DEFAULT_MARRIAGE_WORKBOOK_FILE_ID,
)
from marriage.records import (
    confirm_marriage_image_link,
    ensure_marriage_tables,
    get_marriage_dashboard_metrics,
    get_marriage_record_by_id,
    parse_mc_filename,
    register_new_marriage_record,
    resolve_quick_marriage_id,
    search_marriage_records,
)


class TestMarriageSystem(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp_dir = tempfile.TemporaryDirectory()
        self.db_path = os.path.join(self.tmp_dir.name, "test_marriage.sqlite3")
        self.db = DatabaseManager(self.db_path)
        run_migrations(self.db)
        ensure_marriage_tables(self.db)
        self.master_user = {
            "account_id": "MASTER-001",
            "username": "ctg.baptismal.admin",
            "role": "Master Admin",
            "assigned_parish": "ALL",
        }
        self.parish_leader_user = {
            "account_id": "PARISH-CABA",
            "username": "caba_bap_admin",
            "role": "Parish Leader Admin",
            "assigned_parish": "Caba Parish",
        }

    def tearDown(self) -> None:
        self.tmp_dir.cleanup()

    def test_google_drive_ids_and_quick_id_resolver(self) -> None:
        self.assertEqual(DEFAULT_MARRIAGE_ROOT_FOLDER_ID, "1QTvhE_zJ52TBAgdgVlybVH5UODbdAA2A")
        self.assertEqual(DEFAULT_MARRIAGE_WORKBOOK_FILE_ID, "1cG0_neGpx89MGYK4iGKhoFJ240R_SGRy")
        self.assertEqual(DEFAULT_MARRIAGE_IMAGES_FOLDER_ID, "1TUDlzDi-LCjxmxYorg099cs-uJlyJPzq")

        self.assertEqual(resolve_quick_marriage_id("5"), "MAR-00005")
        self.assertEqual(resolve_quick_marriage_id("154"), "MAR-00154")
        self.assertEqual(resolve_quick_marriage_id("MC_112"), "MAR-00112")
        self.assertEqual(resolve_quick_marriage_id("MAR-00447"), "MAR-00447")

        mc_num, husband = parse_mc_filename("Francisco Bautista MC_154.jpeg")
        self.assertEqual(mc_num, 154)
        self.assertEqual(husband, "Francisco Bautista")

    def test_master_admin_only_rbac_blocks_parish_leaders(self) -> None:
        # Parish Leader Admin must be strictly blocked from accessing Marriage Records
        with self.assertRaises(AuthorizationError):
            get_marriage_record_by_id(self.db, "154", actor=self.parish_leader_user)

        with self.assertRaises(AuthorizationError):
            search_marriage_records(self.db, query="Bautista", actor=self.parish_leader_user)

        with self.assertRaises(AuthorizationError):
            register_new_marriage_record(
                self.db,
                actor=self.parish_leader_user,
                form_data={"Husband Name": "Test Groom", "Wife Name": "Test Bride"},
            )

    def test_master_admin_search_and_verified_seed_records(self) -> None:
        rec154 = get_marriage_record_by_id(self.db, "154", actor=self.master_user)
        self.assertIsNotNone(rec154)
        self.assertEqual(rec154["Record ID"], "MAR-00154")
        self.assertEqual(rec154["Husband Name"], "Francisco Bautista")
        self.assertEqual(rec154["Wife Name"], "Leonida Yanguas")
        self.assertEqual(rec154["Scanned Image Drive ID"], "1S9AHmga3c_gDxoNZ7gbPDzb8CuGXOqta")

        matches = search_marriage_records(self.db, query="Melita Terora", actor=self.master_user)
        self.assertEqual(len(matches), 1)
        self.assertEqual(matches[0]["Record ID"], "MAR-00005")

        metrics = get_marriage_dashboard_metrics(self.db, actor=self.master_user)
        self.assertGreaterEqual(metrics["total_records"], 24)
        self.assertGreaterEqual(metrics["with_scanned_images"], 24)

    def test_new_marriage_registration_and_certificate_pdf_generation(self) -> None:
        reg_out = register_new_marriage_record(
            self.db,
            actor=self.master_user,
            form_data={
                "Husband Name": "Juan Dela Cruz",
                "Husband Age / Birthdate": "27 years old",
                "Husband Residence": "Caba, La Union",
                "Husband Father": "Pedro Dela Cruz",
                "Husband Mother": "Maria Santos",
                "Wife Name": "Clara Reyes",
                "Wife Age / Birthdate": "25 years old",
                "Wife Residence": "Caba, La Union",
                "Wife Father": "Jose Reyes",
                "Wife Mother": "Elena Gomez",
                "Parish": "Caba Parish",
                "Place of Marriage": "Church of the Temple of God, Caba, La Union",
                "Date of Marriage": "November 28, 2026",
                "Marriage License No.": "9988776",
                "Officiating Priest / Minister": "Msgr. Servant Troadio T. Abaquita",
                "Witnesses / Sponsors": "Sponsor One, Sponsor Two",
                "MC Number": "MC_450",
            },
        )
        self.assertTrue(reg_out["ok"])
        self.assertEqual(reg_out["record_id"], "MAR-00450")

        rec450 = get_marriage_record_by_id(self.db, "450", actor=self.master_user)
        self.assertIsNotNone(rec450)

        for orient in ("Portrait", "Landscape"):
            cert_res = generate_and_archive_marriage_certificate(
                db=self.db,
                actor=self.master_user,
                record=rec450,
                orientation=orient,
                output_dir=self.tmp_dir.name,
            )
            self.assertTrue(cert_res["ok"])
            self.assertTrue(cert_res["pdf_bytes"].startswith(b"%PDF-"))
            self.assertGreater(len(cert_res["pdf_bytes"]), 1000)

        certs = list_generated_marriage_certificates(self.db, actor=self.master_user)
        self.assertEqual(len(certs), 2)

        conf = confirm_marriage_image_link(
            db=self.db,
            record_id="450",
            drive_file_id="18l-85EWRs7YRJnjAuSsIbdANW7fgpV8b",
            filename="Juan Dela Cruz MC_450.jpeg",
            actor=self.master_user,
        )
        self.assertTrue(conf["ok"])


if __name__ == "__main__":
    unittest.main()
