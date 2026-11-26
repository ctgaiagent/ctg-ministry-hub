import os
import tempfile
import unittest
from baptismal.authentication import AuthService, hash_password_pbkdf2, verify_password_pbkdf2
from baptismal.database import DatabaseManager
from baptismal.migrations import run_migrations


class TestAuthentication(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.NamedTemporaryFile(suffix=".db", delete=False)
        self.tmp.close()
        self.db = DatabaseManager(f"sqlite:///{self.tmp.name}")
        run_migrations(self.db)
        self.auth = AuthService(self.db)

    def tearDown(self):
        if os.path.exists(self.tmp.name):
            os.unlink(self.tmp.name)

    def test_pbkdf2_hashing_and_verification(self):
        h = hash_password_pbkdf2("SyntheticPass#2026")
        self.assertTrue(h.startswith("pbkdf2_sha256$"))
        self.assertTrue(verify_password_pbkdf2("SyntheticPass#2026", h))
        self.assertFalse(verify_password_pbkdf2("WrongPass", h))

    def test_master_admin_auth_and_lockout(self):
        acc_id = self.auth.create_account("bap_admin_test", "Test Admin", "Master Admin", plain_password="StrongPassword!99")
        ok, user, _ = self.auth.authenticate("bap_admin_test", "StrongPassword!99")
        self.assertTrue(ok)
        self.assertEqual(user["account_id"], acc_id)
        valid_s, _ = self.auth.validate_session(user)
        self.assertTrue(valid_s)

        # 5 failed logins trigger lockout
        for _ in range(5):
            ok_bad, _, err_msg = self.auth.authenticate("bap_admin_test", "BadPass!")
            self.assertFalse(ok_bad)
            self.assertIn("Invalid credentials", err_msg)

        # Even valid password is rejected while locked
        ok_locked, _, _ = self.auth.authenticate("bap_admin_test", "StrongPassword!99")
        self.assertFalse(ok_locked)

        # Verify audit log recorded LOGIN_SUCCESS, LOGIN_FAILED, ACCOUNT_LOCKED
        ev_types = {r["event_type"] for r in self.auth.audit.list_events()}
        self.assertIn("LOGIN_SUCCESS", ev_types)
        self.assertIn("ACCOUNT_LOCKED", ev_types)


if __name__ == "__main__":
    unittest.main()
