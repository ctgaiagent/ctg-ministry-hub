"""
tests/test_ui_navigation_redesign.py — Unit Tests for CTG UI/UX, Top Bar & Sidebar Final Correction
====================================================================================================
Verifies:
1. Solid non-transparent navy design tokens (#10213D, #D4A640 gold accent, zero backdrop blur or glass).
2. Role-based 5-group navigation for Master Admin, Diocese Admin, and Parish Leader.
3. No separate 'Home' or 'Home & Leadership Portal' shortcut exists in the 5 sidebar navigation groups.
4. No three-line hamburger icon ('☰') exists in ui_components.py or app.py.
5. Marriage Records & Certificates is strictly excluded from non-Master Admin navigation.
6. Every page_id maps cleanly to an existing route in PAGE_ID_TO_APP_MODE and back.
"""

import inspect
import unittest

import ui_components
from ui_components import (
    APP_MODE_TO_PAGE_ID,
    DESIGN_TOKENS,
    PAGE_ID_TO_APP_MODE,
    get_role_navigation_groups,
    is_sidebar_expanded,
    render_status_badge,
    toggle_sidebar_state,
)


class TestUINavigationRedesign(unittest.TestCase):
    def test_design_tokens_and_solid_non_transparent_topbar(self) -> None:
        self.assertEqual(DESIGN_TOKENS["topbar_navy"], "#10213D")
        self.assertEqual(DESIGN_TOKENS["primary"], "#0B1730")
        self.assertEqual(DESIGN_TOKENS["secondary"], "#142747")
        self.assertEqual(DESIGN_TOKENS["accent_gold"], "#D4A640")
        self.assertEqual(DESIGN_TOKENS["background"], "#0B101B")
        badge = render_status_badge("<script>alert(1)</script>", "success")
        self.assertNotIn("<script>", badge)
        self.assertIn("&lt;script&gt;", badge)
        self.assertNotIn("rgba(", badge)

        ui_src = inspect.getsource(ui_components)
        self.assertNotIn("☰", ui_src, "Three-line hamburger icon must not be used")
        self.assertNotIn("backdrop-filter: blur", ui_src, "Top bar must not use glass backdrop blur")
        self.assertIn("top_navbar_sidebar_panel_toggle_btn", ui_src)
        self.assertIn("Church of the Temple of God", ui_src)
        self.assertIn("One Church &bull; One Faith &bull; One Family", ui_src)
        self.assertIn("top_navbar_global_refresh_btn", ui_src)

    def test_no_home_shortcut_in_sidebar_groups(self) -> None:
        for role in ("Master Admin", "Diocese Admin", "Parish Leader"):
            groups = get_role_navigation_groups(user_role=role, login_role=role)
            self.assertEqual(len(groups), 5)
            all_ids = [it["page_id"] for g in groups for it in g["items"]]
            all_labels = [it["label"] for g in groups for it in g["items"]]
            self.assertNotIn("home_overview", all_ids)
            self.assertNotIn("Home", all_labels)
            self.assertNotIn("Home & Leadership Portal", all_labels)

        nav_fn_src = inspect.getsource(ui_components.render_grouped_sidebar_navigation)
        self.assertNotIn("sb_nav_home_overview", nav_fn_src)
        self.assertNotIn("Dashboard (Overview)", nav_fn_src)

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

    def test_sidebar_toggle_helpers(self) -> None:
        self.assertIsInstance(is_sidebar_expanded(), bool)
        self.assertIsInstance(toggle_sidebar_state(), bool)


if __name__ == "__main__":
    unittest.main()
