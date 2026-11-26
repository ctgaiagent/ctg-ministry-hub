from datetime import datetime, timedelta, timezone
import os
import tempfile
import unittest
from baptismal.authentication import AuthService
from baptismal.authorization import AuthorizationService
from baptismal.database import DatabaseManager
from baptismal.migrations import run_migrations
from baptismal.temporary_access import TemporaryAccessService


class TestTemporaryAccess(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.NamedTemporaryFile(suffix=".db", delete=False)
        self.tmp.close()
        self.db = DatabaseManager(f"sqlite:///{self.tmp.name}")
        run_migrations(self.db)
        self.auth = AuthService(self.db)
        self.authz = AuthorizationService(self.db)
        self.temp = TemporaryAccessService(self.db, self.authz)
        self.admin_id = self.auth.create_account("admin_t", "Admin T", "Master Admin", plain_password="AdminPass!2026")
        self.req_acc = self.auth.create_account("req_t", "Req T", "Temporary Requester", plain_password="ReqPass!2026")

    def tearDown(self):
        if os.path.exists(self.tmp.name):
            os.unlink(self.tmp.name)

    def test_grant_expiration_and_immediate_revocation(self):
        req_id = self.temp.submit_request(self.req_acc, "req_t", "Synthetic Person", "School requirement")
        grant_id = self.temp.approve_request_and_grant_access(
            self.admin_id, req_id, "HIST-00010", ["view_record", "generate_certificate"], duration_hours=2.0
        )

        self.assertTrue(self.authz.has_temporary_record_permission(self.req_acc, "HIST-00010", "view_record"))
        self.assertFalse(self.authz.has_temporary_record_permission(self.req_acc, "HIST-00010", "download_certificate"))
        self.assertFalse(self.authz.has_temporary_record_permission(self.req_acc, "HIST-00099", "view_record"))

        # Immediate revocation
        self.temp.revoke_grant_immediately(self.admin_id, grant_id)
        self.assertFalse(self.authz.has_temporary_record_permission(self.req_acc, "HIST-00010", "view_record"))

        # Expiration check
        grant_id_2 = self.temp.approve_request_and_grant_access(
            self.admin_id, req_id, "HIST-00011", ["view_record"], duration_hours=1.0
        )
        past_iso = (datetime.now(timezone.utc) - timedelta(minutes=5)).isoformat()
        self.db.execute("UPDATE temporary_access_grants SET expires_at = %s WHERE grant_id = %s", (past_iso, grant_id_2))
        self.assertFalse(self.authz.has_temporary_record_permission(self.req_acc, "HIST-00011", "view_record"))


if __name__ == "__main__":
    unittest.main()
