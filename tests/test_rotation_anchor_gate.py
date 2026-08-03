"""회전 페이지 기준 앵커 게이팅과 회전 오프셋 회귀시험."""

import unittest

import fitz
import numpy as np

from engine_Drag import HybridTextExtractor
from page_preprocessing import PagePreprocessResult


class RotationAnchorGateTests(unittest.TestCase):
    def setUp(self):
        self.extractor = object.__new__(HybridTextExtractor)
        self.extractor.last_keyword_status = None
        self.document = fitz.open()
        self.page = self.document.new_page(width=500, height=500)

    def tearDown(self):
        self.document.close()

    @staticmethod
    def corrected_result():
        return PagePreprocessResult(
            status="corrected",
            processed_image=np.zeros((10, 10, 3), dtype=np.uint8),
            detected_orientation=90,
            orientation_correction=270,
            orientation_confidence=0.99,
            orientation_margin=0.80,
            orientation_applied=True,
        )

    def test_rotation_without_anchor_is_rejected_page_locally(self):
        result = self.corrected_result()
        self.extractor.last_preprocess_result = result
        rect = self.extractor.adjusted_rect(
            self.page,
            {
                "rect": fitz.Rect(120, 90, 180, 110),
                "keyword": "",
                "anchor_rect": None,
            },
        )
        self.assertTrue(rect.is_empty)
        self.assertEqual(result.status, "rejected")
        self.assertEqual(result.failure_reason, "anchor_required_for_rotation")

    def test_anchor_relative_offset_rotates_with_only_current_page(self):
        result = self.corrected_result()
        self.extractor.last_preprocess_result = result
        current_anchor = fitz.Rect(190, 190, 210, 210)
        self.extractor.find_keyword_candidates = (
            lambda _page, _origin, _keyword: [current_anchor]
        )
        rect = self.extractor.adjusted_rect(
            self.page,
            {
                "rect": fitz.Rect(130, 90, 170, 110),
                "keyword": "ANCHOR",
                "anchor_rect": fitz.Rect(90, 90, 110, 110),
                "tracking_anchor_rect": fitz.Rect(90, 90, 110, 110),
                "offset_x": 40,
                "offset_y": 0,
            },
        )
        self.assertEqual(result.reference_validation, "accepted")
        self.assertAlmostEqual(rect.x0, 190.0)
        self.assertAlmostEqual(rect.y0, 230.0)
        self.assertAlmostEqual(rect.x1, 210.0)
        self.assertAlmostEqual(rect.y1, 270.0)


if __name__ == "__main__":
    unittest.main()