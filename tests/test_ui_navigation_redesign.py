"""
tests/test_ui_navigation_redesign.py — Unit Tests for CTG UI/UX & Navigation Redesign
=====================================================================================
Verifies:
1. Role-based navigation groups for Master Admin, Diocese Admin, and Parish Leader.
2. Marriage Records & Certificates is strictly excluded from non-Master Admin navigation.
3. Every page_id maps cleanly to an existing route in PAGE_ID_TO_APP_MODE and back.
4. Design system tokens and HTML escaping helpers work without error.
"""

import unittest

from ui_components import (
    APP_MODE_TO_PAGE_ID,
    DESIGN_TOKENS,
    PAGE_ID_TO_APP_MODE,
    get_role_navigation_groups,
    render_status_badge,
)


class TestUINavigationRedesign(unittest.TestCase):
    def test_design_tokens_present(self) -> None:
        self.assertEqual(DESIGN_TOKENS["primary"], "#0B1730")
        self.assertEqual(DESIGN_TOKENS["secondary"], "#142747")
        self.assertEqual(DESIGN_TOKENS["accent_gold"], "#D4A640")
        self.assertEqual(DESIGN_TOKENS["background"], "#0B101B")
        badge = render_status_badge("<script>alert(1)</script>", "success")
        self.assertNotIn("<script>", badge)
        self.assertIn("&lt;script&gt;", badge)

    def test_master_admin_five_groups_and_marriage_access(self) -> None:
        groups = get_role_navigation_groups(user_role="Master Admin", login_role="Master Admin")
        self.assertEqual(len(groups), 5)
        all_ids = {it["page_id"] for g in groups for it in g["items"]}
        self.assertIn("marriage_portal", all_ids)
        self.assertIn("baptismal_portal", all_ids)
        self.assertIn("master_approval_queue", all_ids)
        self.assertIn("live_activity", all_ids)
        for pid in all_ids:
            self.assertIn(pid, PAGE_ID_TO_APP_MODE)
            mode = PAGE_ID_TO_APP_MODE[pid]
            self.assertEqual(APP_MODE_TO_PAGE_ID[mode], pid)

    def test_parish_leader_and_diocese_admin_rbac_navigation(self) -> None:
        for role in ("Parish Leader", "Diocese Admin"):
            groups = get_role_navigation_groups(user_role=role, login_role=role)
            self.assertEqual(len(groups), 5)
            all_ids = {it["page_id"] for g in groups for it in g["items"]}
            self.assertNotIn("marriage_portal", all_ids)
            self.assertNotIn("master_approval_queue", all_ids)
            self.assertNotIn("live_activity", all_ids)
            self.assertNotIn("notion_crm", all_ids)
            self.assertIn("baptismal_portal", all_ids)
            self.assertIn("parish_inbox", all_ids)
            self.assertIn("weekly_report", all_ids)


if __name__ == "__main__":
    unittest.main()
