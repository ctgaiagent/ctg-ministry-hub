import unittest
from unittest.mock import MagicMock
from baptismal.google_drive import GoogleDriveManager, GoogleDrivePermissionError


class TestGoogleDrivePermissions(unittest.TestCase):
    def test_viewer_only_workbook_is_blocked(self):
        mock_svc = MagicMock()
        mock_svc.files().get().execute.return_value = {
            "id": "mock_wb_id",
            "name": "OVERALL BAPTISMAL CERTIFICATE LIST.xlsx",
            "capabilities": {"canEdit": False, "canDownload": True},
            "headRevisionId": "rev_1",
        }
        mgr = GoogleDriveManager(drive_service=mock_svc, config={"workbook_file_id": "mock_wb_id"})
        with self.assertRaises(GoogleDrivePermissionError):
            mgr.verify_workbook_write_access("mock_wb_id")

    def test_editor_workbook_is_allowed(self):
        mock_svc = MagicMock()
        mock_svc.files().get().execute.return_value = {
            "id": "mock_wb_id",
            "name": "OVERALL BAPTISMAL CERTIFICATE LIST.xlsx",
            "capabilities": {"canEdit": True, "canDownload": True},
            "headRevisionId": "rev_2",
        }
        mgr = GoogleDriveManager(drive_service=mock_svc, config={"workbook_file_id": "mock_wb_id"})
        meta = mgr.verify_workbook_write_access("mock_wb_id")
        self.assertEqual(meta["headRevisionId"], "rev_2")


if __name__ == "__main__":
    unittest.main()
