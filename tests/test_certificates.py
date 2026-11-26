import os
import tempfile
import unittest
from baptismal.authentication import AuthService
from baptismal.authorization import AuthorizationService
from baptismal.certificates import CertificateService
from baptismal.config import HISTORICAL_27_COLUMNS
from baptismal.database import DatabaseManager
from baptismal.migrations import run_migrations
from baptismal.registration import RegistrationService


class TestCertificates(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.NamedTemporaryFile(suffix=".db", delete=False)
        self.tmp.close()
        self.db = DatabaseManager(f"sqlite:///{self.tmp.name}")
        run_migrations(self.db)
        self.auth = AuthService(self.db)
        self.authz = AuthorizationService(self.db)
        self.admin_id = self.auth.create_account("admin_c", "Admin Cert", "Master Admin", plain_password="AdminPass!2026")
        self.reg = RegistrationService(self.db, self.authz)
        self.certs = CertificateService(self.db, self.authz)

    def tearDown(self):
        if os.path.exists(self.tmp.name):
            os.unlink(self.tmp.name)

    def test_generate_portrait_and_landscape_and_archive(self):
        payload = {c: "" for c in HISTORICAL_27_COLUMNS}
        payload.update({
            "Parish": "TEST PARISH",
            "Name of the Child": "Synthetic Cert Person",
            "Date of Baptism": "2026-06-15",
            "Year of Baptism": "2026",
        })
        reg_res = self.reg.register_new_baptism(self.admin_id, payload, sync_to_drive_excel=False)
        rec_id = reg_res["record_id"]

        res_p = self.certs.generate_certificate(self.admin_id, rec_id, orientation="portrait", upload_to_drive=False)
        self.assertTrue(res_p["pdf_bytes"].startswith(b"%PDF-"))
        self.assertEqual(res_p["issuance_status"], "ORIGINAL")

        res_l = self.certs.generate_certificate(
            self.admin_id, rec_id, orientation="landscape", is_reissue=True, reissue_reason="Damaged copy", upload_to_drive=False
        )
        self.assertTrue(res_l["pdf_bytes"].startswith(b"%PDF-"))
        self.assertEqual(res_l["issuance_status"], "REISSUED")

        archived = self.certs.list_certificates(rec_id)
        self.assertEqual(len(archived), 2)


if __name__ == "__main__":
    unittest.main()
