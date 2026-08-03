"""기준 파일·페이지 수동 탐색 상태 시험."""
# Copyright (C) 2026 두부코드(DOOBOO_CODE)
# SPDX-License-Identifier: AGPL-3.0-only

import os
import unittest

from reference_navigation import ReferenceNavigationState


class ReferenceNavigationTests(unittest.TestCase):
    def test_suggestion_is_initial_position(self):
        state = ReferenceNavigationState.from_suggestion(
            ["a.pdf", "b.pdf"],
            [2, 7],
            {"pdf_path": "b.pdf", "page_index": 5},
        )
        self.assertEqual(state.current_path, os.path.abspath("b.pdf"))
        self.assertEqual(state.page_index, 5)

    def test_manual_navigation_has_no_three_page_limit(self):
        state = ReferenceNavigationState(("a.pdf",), (8,))
        for _index in range(7):
            state.next_page()
        self.assertEqual(state.page_index, 7)
        self.assertFalse(state.can_next_page)

    def test_file_navigation_opens_first_page(self):
        state = ReferenceNavigationState(("a.pdf", "b.pdf"), (4, 3), page_index=2)
        state.next_file()
        self.assertEqual(state.file_index, 1)
        self.assertEqual(state.page_index, 0)
        state.next_page().previous_file()
        self.assertEqual(state.file_index, 0)
        self.assertEqual(state.page_index, 0)

    def test_boundaries_do_not_wrap(self):
        state = ReferenceNavigationState(("a.pdf", "b.pdf"), (1, 1))
        state.previous_file().previous_page()
        self.assertEqual((state.file_index, state.page_index), (0, 0))
        state.next_file().next_file().next_page()
        self.assertEqual((state.file_index, state.page_index), (1, 0))

    def test_button_availability_tracks_position(self):
        state = ReferenceNavigationState(("a.pdf", "b.pdf"), (2, 3))
        self.assertFalse(state.can_previous_file)
        self.assertFalse(state.can_previous_page)
        self.assertTrue(state.can_next_file)
        self.assertTrue(state.can_next_page)
        state.next_page()
        self.assertFalse(state.can_next_page)
        state.next_file()
        self.assertTrue(state.can_previous_file)
        self.assertFalse(state.can_previous_page)


if __name__ == "__main__":
    unittest.main()
