from datetime import datetime, timedelta, timezone
import io
import os
import tempfile
import unittest
import openpyxl

from baptismal.authentication import AuthService
from baptismal.authorization import AuthorizationError, AuthorizationService
from baptismal.certificates import CertificateService
from baptismal.config import HISTORICAL_27_COLUMNS, OFFICIAL_NINE_PARISHES
from baptismal.database import DatabaseManager
from baptismal.historical_records import HistoricalRecordsService
from baptismal.image_archive import ImageArchiveService
from baptismal.migrations import run_migrations
from baptismal.registration import RegistrationService
from baptismal.temporary_access import TemporaryAccessService


class TestRoleBasedTenAccountsAndHongKongExclusion(unittest.TestCase):
    """Validates all 10 account types (1 Master Admin + 9 Parish Leader Admins) and strict Hong Kong exclusion."""

    def setUp(self):
        self.tmp = tempfile.NamedTemporaryFile(suffix=".db", delete=False)
        self.tmp.close()
        self.db = DatabaseManager(f"sqlite:///{self.tmp.name}")
        run_migrations(self.db)
        self.auth = AuthService(self.db)
        self.authz = AuthorizationService(self.db)
        self.hist = HistoricalRecordsService(self.db, authz=self.authz)
        self.img_arc = ImageArchiveService(self.db, authz=self.authz)
        self.reg = RegistrationService(self.db, authz=self.authz)
        self.temp = TemporaryAccessService(self.db, authz=self.authz)
        self.certs = CertificateService(self.db, authz=self.authz)

        # 1. Create Master Admin account
        self.master_id = self.auth.create_account(
            username="master_bap_admin",
            full_name="Master Admin",
            role="Master Admin",
            plain_password="MasterPassword#2026",
        )
        # 2. Provision all 9 official Parish Leader Admin accounts
        self.parish_admins = self.auth.provision_nine_parish_leader_admins(
            self.master_id, initial_password="ParishPassword#2026"
        )

    def tearDown(self):
        if os.path.exists(self.tmp.name):
            os.unlink(self.tmp.name)

    def test_ten_accounts_and_hong_kong_exclusion(self):
        # Verify exact 9 official parishes and 0 Hong Kong
        self.assertEqual(len(OFFICIAL_NINE_PARISHES), 9)
        self.assertEqual(len(self.parish_admins), 9)
        for p in OFFICIAL_NINE_PARISHES:
            self.assertNotIn("HONG KONG", p.upper())

        # Creating a Hong Kong Parish Admin account must fail
        with self.assertRaises(ValueError):
            self.auth.create_account(
                username="hk_bap_admin",
                full_name="HK Admin",
                role="Parish Leader Admin",
                assigned_parish="Hong Kong Parish",
                plain_password="HkPassword#2026",
            )

        # Import historical records including a Hong Kong row; verify source preserved but excluded from operational scope
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Historical Records"
        ws.append(HISTORICAL_27_COLUMNS)
        for idx, p_name in enumerate(OFFICIAL_NINE_PARISHES + ["Hong Kong Parish"], start=1):
            row_d = {c: "" for c in HISTORICAL_27_COLUMNS}
            row_d["Parish"] = p_name
            row_d["Name of the Child"] = f"Child {idx}"
            row_d["Year of Baptism"] = "1970"
            row_d["Book Number"] = "1"
            row_d["Page Number"] = str(idx)
            ws.append([row_d[c] for c in HISTORICAL_27_COLUMNS])
        buf = io.BytesIO()
        wb.save(buf)
        tot_imported = self.hist.import_sheet1_from_excel_bytes(buf.getvalue())
        self.assertEqual(tot_imported, 10)  # All 10 source rows preserved in DB

        # Operational search and dashboard stats must exclude Hong Kong (showing 9)
        rows_op, cnt_op = self.hist.search_records(actor_account_id=self.master_id)
        self.assertEqual(cnt_op, 9)
        stats_op = self.hist.get_dashboard_statistics(actor_account_id=self.master_id)
        self.assertEqual(stats_op["total_historical"], 9)

    def test_all_nine_parish_admins_isolation_and_permissions(self):
        created_record_ids = {}

        for pa in self.parish_admins:
            aid = pa["account_id"]
            parish = pa["assigned_parish"]

            # 1. Parish Admin CANNOT search historical records or access BOOK 1-9
            with self.assertRaises(AuthorizationError):
                self.hist.search_records(actor_account_id=aid)
            with self.assertRaises(AuthorizationError):
                self.img_arc.list_by_book("BOOK 1", actor_account_id=aid)

            # 2. Parish Admin registers a new baptism; even if payload tries to spoof another parish, server forces assigned_parish
            spoof_payload = {c: "" for c in HISTORICAL_27_COLUMNS}
            spoof_payload.update({
                "Parish": "Hong Kong Parish",  # Spoof attempt
                "Name of the Child": f"New Child of {parish}",
                "Date of Baptism": "2026-08-01",
                "Year of Baptism": "2026",
                "Name of the Father": f"Father {parish}",
                "Name of the Mother": f"Mother {parish}",
            })
            reg_res = self.reg.register_new_baptism(aid, spoof_payload, sync_to_drive_excel=False)
            self.assertEqual(reg_res["parish"], parish)
            self.assertEqual(reg_res["verification_status"], "Pending Verification")
            created_record_ids[parish] = reg_res["record_id"]

            # 3. Parish Admin can edit their own submission while Pending Verification
            spoof_payload["Sponsors"] = "Updated Sponsor"
            self.reg.update_own_submission(aid, reg_res["record_id"], spoof_payload)

            # 4. Parish Admin CANNOT verify/approve registrations or generate certificates independently
            with self.assertRaises(AuthorizationError):
                self.reg.verify_registration(aid, reg_res["record_id"], "Approved")
            with self.assertRaises(AuthorizationError):
                self.certs.generate_certificate(aid, "HIST-00001", upload_to_drive=False)

        # 5. Cross-parish isolation: each Parish Admin sees ONLY their own 1 submission (never the other 8 parishes)
        for pa in self.parish_admins:
            aid = pa["account_id"]
            parish = pa["assigned_parish"]
            my_subs = self.reg.list_submissions_for_actor(aid)
            self.assertEqual(len(my_subs), 1)
            self.assertEqual(my_subs[0]["parish"], parish)
            p_stats = self.reg.get_parish_dashboard_statistics(aid)
            self.assertEqual(p_stats["my_new_registrations"], 1)
            self.assertEqual(p_stats["pending_verification"], 1)

        # 6. Once Master Admin approves Lipay Norte's submission, Lipay Norte Admin can no longer edit it
        lipay_admin = next(p for p in self.parish_admins if p["assigned_parish"] == "Lipay Norte Parish")
        lipay_rec_id = created_record_ids["Lipay Norte Parish"]
        self.reg.verify_registration(self.master_id, lipay_rec_id, "Approved", "Verified by Master Admin")
        with self.assertRaises(AuthorizationError):
            self.reg.update_own_submission(lipay_admin["account_id"], lipay_rec_id, {
                "Parish": "Lipay Norte Parish",
                "Name of the Child": "Tamper Attempt",
                "Date of Baptism": "2026-08-01",
                "Year of Baptism": "2026",
            })

        # 7. Historical Record & Certified Certificate Request + Master Admin Release Gating
        malanas_admin = next(p for p in self.parish_admins if p["assigned_parish"] == "Malanas Parish")
        req_id = self.temp.submit_request(
            requester_account_id=malanas_admin["account_id"],
            requester_username=malanas_admin["username"],
            requested_person_name="New Child of Lipay Norte Parish",
            purpose="Marriage requirement",
            request_type="CERTIFIED_CERTIFICATE",
            parish="Malanas Parish",
        )
        # Master Admin generates an UNRELEASED certificate first -> Malanas Admin CANNOT download it
        unreleased_cert = self.certs.generate_certificate(
            self.master_id, lipay_rec_id, upload_to_drive=False, release_to_requester=False
        )
        with self.assertRaises(AuthorizationError):
            self.certs.download_released_certificate(malanas_admin["account_id"], unreleased_cert["certificate_id"])

        # Master Admin officially approves and releases certificate for Malanas Admin's request
        released_cert = self.certs.approve_and_release_certificate_for_request(
            admin_account_id=self.master_id,
            request_id=req_id,
            record_id=lipay_rec_id,
            upload_to_drive=False,
        )
        # Malanas Admin CAN download their released certificate
        pdf_bytes = self.certs.download_released_certificate(malanas_admin["account_id"], released_cert["certificate_id"])
        self.assertTrue(pdf_bytes.startswith(b"%PDF-"))

        # Lipay Norte Admin CANNOT download a certificate released to Malanas Admin
        with self.assertRaises(AuthorizationError):
            self.certs.download_released_certificate(lipay_admin["account_id"], released_cert["certificate_id"])


if __name__ == "__main__":
    unittest.main()
