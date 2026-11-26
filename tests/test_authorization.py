import os
import tempfile
import unittest
from baptismal.authentication import AuthService
from baptismal.authorization import AuthorizationError, AuthorizationService
from baptismal.database import DatabaseManager
from baptismal.migrations import run_migrations


class TestAuthorization(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.NamedTemporaryFile(suffix=".db", delete=False)
        self.tmp.close()
        self.db = DatabaseManager(f"sqlite:///{self.tmp.name}")
        run_migrations(self.db)
        self.auth = AuthService(self.db)
        self.authz = AuthorizationService(self.db)
        self.admin_id = self.auth.create_account("admin1", "Admin One", "Master Admin", plain_password="AdminPass!2026")
        self.req_id = self.auth.create_account("req1", "Requester One", "Temporary Requester", plain_password="ReqPass!2026")

    def tearDown(self):
        if os.path.exists(self.tmp.name):
            os.unlink(self.tmp.name)

    def test_unauthorized_user_denied_admin_and_records(self):
        self.assertTrue(self.authz.is_master_admin(self.admin_id))
        self.assertFalse(self.authz.is_master_admin(self.req_id))
        with self.assertRaises(AuthorizationError):
            self.authz.require_master_admin(self.req_id, action="register_new_baptism")
        with self.assertRaises(AuthorizationError):
            self.authz.require_record_permission(self.req_id, "HIST-00001", "view_record")

        ev_types = {r["event_type"] for r in self.authz.audit.list_events()}
        self.assertIn("ACCESS_DENIED", ev_types)


if __name__ == "__main__":
    unittest.main()
