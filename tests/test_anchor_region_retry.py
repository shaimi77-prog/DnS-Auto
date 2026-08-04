import unittest

import fitz

from engine_Drag import (
    HybridTextExtractor,
    OCR_MATCH,
    OCR_NO_MATCH,
    OCR_VALUE_RECOGNITION_DPI,
)


class AnchorRegionRetryTests(unittest.TestCase):
    def setUp(self):
        self.extractor = object.__new__(HybridTextExtractor)
        self.extractor._ocr_cache = {}
        self.extractor._ocr_statistics = HybridTextExtractor._new_ocr_statistics()
        self.extractor.last_keyword_status = OCR_NO_MATCH
        self.extractor.last_keyword_reason = None
        self.extractor.last_preprocess_result = None
        self.document = fitz.open()
        self.page = self.document.new_page(width=500, height=500)

    def tearDown(self):
        self.document.close()

    def test_retry_runs_one_210_dpi_inference_and_prefers_earliest_stage(self):
        calls = []
        # The closer candidate is in a later stage; the stage-1 candidate wins.
        stage_one = fitz.Rect(172, 197, 178, 203)
        later_but_closer = fitz.Rect(227, 197, 233, 203)

        def fake_detect(_page, clip, dpi):
            calls.append((fitz.Rect(clip), dpi))
            return {
                "boxes": (stage_one, later_but_closer),
                "texts": ("날짜,", "날짜,"),
                "scores": (0.91, 0.99),
                "width": 1,
                "height": 1,
                "source_width": 1,
                "source_height": 1,
                "transform_matrix": None,
            }

        self.extractor._ocr_detect = fake_detect
        self.extractor._ocr_box_to_pdf_rect = (
            lambda box, *_args, **_kwargs: fitz.Rect(box)
        )
        origin = fitz.Rect(180, 180, 220, 220)

        selected = self.extractor._retry_ocr_keyword_candidate(
            self.page, origin, "날짜,"
        )

        self.assertEqual(selected, stage_one)
        self.assertEqual(len(calls), 1)
        self.assertEqual(calls[0][1], OCR_VALUE_RECOGNITION_DPI)
        self.assertEqual(
            self.extractor._ocr_statistics["anchor_region_retry_count"], 1
        )

    def test_same_stage_prefers_nearest_candidate(self):
        farther = fitz.Rect(184, 184, 190, 190)
        nearer = fitz.Rect(197, 197, 203, 203)
        self.extractor._ocr_detect = lambda *_args, **_kwargs: {
            "boxes": (farther, nearer),
            "texts": ("날짜,", "날짜,"),
            "scores": (0.99, 0.80),
            "width": 1,
            "height": 1,
            "source_width": 1,
            "source_height": 1,
            "transform_matrix": None,
        }
        self.extractor._ocr_box_to_pdf_rect = (
            lambda box, *_args, **_kwargs: fitz.Rect(box)
        )

        selected = self.extractor._retry_ocr_keyword_candidate(
            self.page, fitz.Rect(180, 180, 220, 220), "날짜,"
        )

        self.assertEqual(selected, nearer)

    def test_adjusted_rect_uses_absolute_rect_when_retry_has_no_match(self):
        drag_rect = fitz.Rect(120, 120, 180, 160)
        mapping = {
            "rect": drag_rect,
            "keyword": "날짜,",
            "anchor_rect": fitz.Rect(90, 90, 110, 110),
            "tracking_anchor_rect": fitz.Rect(90, 90, 110, 110),
            "offset_x": 30,
            "offset_y": 30,
        }

        def no_page_match(*_args):
            self.extractor.last_keyword_status = OCR_NO_MATCH
            return []

        retries = []
        self.extractor.find_keyword_candidates = no_page_match
        self.extractor._retry_ocr_keyword_candidate = (
            lambda *_args: retries.append(True) or None
        )

        result = self.extractor.adjusted_rect(self.page, mapping)

        self.assertEqual(result, drag_rect)
        self.assertEqual(len(retries), 1)
        self.assertEqual(self.extractor.last_keyword_status, OCR_NO_MATCH)

    def test_adjusted_rect_uses_retry_match(self):
        mapping = {
            "rect": fitz.Rect(120, 120, 180, 160),
            "keyword": "날짜,",
            "anchor_rect": fitz.Rect(90, 90, 110, 110),
            "tracking_anchor_rect": fitz.Rect(90, 90, 110, 110),
            "offset_x": 30,
            "offset_y": 30,
        }
        retry_anchor = fitz.Rect(100, 70, 120, 90)

        def no_page_match(*_args):
            self.extractor.last_keyword_status = OCR_NO_MATCH
            return []

        self.extractor.find_keyword_candidates = no_page_match
        self.extractor._retry_ocr_keyword_candidate = lambda *_args: retry_anchor

        result = self.extractor.adjusted_rect(self.page, mapping)

        self.assertEqual(self.extractor.last_keyword_status, OCR_MATCH)
        self.assertEqual(result, fitz.Rect(130, 100, 190, 140))


if __name__ == "__main__":
    unittest.main()
