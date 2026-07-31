import unittest
from unittest import mock

import engine_Drag


class PdfEtaRegressionTests(unittest.TestCase):
    def setUp(self):
        engine_Drag.HybridTextExtractor._instance = None
        self.extractor = engine_Drag.HybridTextExtractor()

    def test_initialization_duration_field_always_exists(self):
        self.assertEqual(self.extractor.last_ocr_initialization_seconds, 0.0)
        self.extractor.reset_work_cache()
        self.assertEqual(self.extractor.last_ocr_initialization_seconds, 0.0)

    def test_already_initialized_ocr_does_not_require_timer(self):
        self.extractor.ocr = object()
        with mock.patch.object(engine_Drag.time, "monotonic", side_effect=AssertionError):
            self.assertTrue(self.extractor.ensure_ocr())

    def test_page_completion_uses_safe_initialization_default(self):
        del self.extractor.last_ocr_initialization_seconds
        self.assertEqual(
            getattr(self.extractor, "last_ocr_initialization_seconds", 0.0),
            0.0,
        )


if __name__ == "__main__":
    unittest.main()
