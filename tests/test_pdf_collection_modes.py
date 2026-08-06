import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import fitz

import engine_Drag as drag
import mcp_server


class PdfCollectionModeTests(unittest.TestCase):
    def setUp(self):
        drag.TEXT_EXTRACTOR.reset_work_cache()

    def tearDown(self):
        drag.TEXT_EXTRACTOR.reset_work_cache()

    @staticmethod
    def _mapping(rect, *, keyword="", anchor_rect=None):
        return {
            "rect": fitz.Rect(rect),
            "keyword": keyword,
            "anchor_rect": fitz.Rect(anchor_rect) if anchor_rect else None,
            "tracking_anchor_rect": fitz.Rect(anchor_rect) if anchor_rect else None,
            "offset_x": 0,
            "offset_y": 0,
        }

    def test_mode_default_and_validation(self):
        self.assertEqual(
            drag.validate_pdf_collection_mode(None), drag.PDF_MODE_STANDARD
        )
        with self.assertRaisesRegex(ValueError, "fast, standard, careful"):
            drag.validate_pdf_collection_mode("turbo")
        self.assertEqual(
            drag.validate_pdf_collection_mode("turbo", fallback=True),
            drag.PDF_MODE_STANDARD,
        )

    def test_fast_mixed_pages_skip_whole_row_without_any_ocr(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "mixed.pdf"
            document = fitz.open()
            for text in ("FIRST", None, "THIRD"):
                page = document.new_page(width=300, height=200)
                if text:
                    page.insert_text((30, 50), text, fontsize=12)
            document.save(path)
            document.close()

            failed = []
            summary = {}
            mapping = {"value": self._mapping((20, 25, 150, 70))}
            with patch.object(
                drag.TEXT_EXTRACTOR,
                "ensure_ocr",
                side_effect=AssertionError("fast mode must not initialize OCR"),
            ) as ensure:
                rows = drag._collect_pdf_rows(
                    [str(path)], ["value"], mapping, False, failed, summary,
                    pdf_collection_mode=drag.PDF_MODE_FAST,
                )

            self.assertEqual([row["value"] for row in rows], ["FIRST", "THIRD"])
            self.assertEqual(summary["processed_pages"], 2)
            self.assertEqual(summary["fast_skipped_page_count"], 1)
            self.assertEqual(summary["fast_skipped_field_count"], 1)
            self.assertEqual(summary["ocr_statistics"]["total_ocr_inference_count"], 0)
            self.assertEqual(failed, [])
            ensure.assert_not_called()

    def test_fast_missing_native_anchor_requires_ocr_without_calling_it(self):
        document = fitz.open()
        page = document.new_page(width=300, height=200)
        mapping = self._mapping(
            (100, 100, 180, 140), keyword="ANCHOR", anchor_rect=(20, 20, 80, 40)
        )
        with patch.object(
            drag.TEXT_EXTRACTOR,
            "ensure_ocr",
            side_effect=AssertionError("fast mode must not initialize OCR"),
        ) as ensure:
            value, details = drag.TEXT_EXTRACTOR.extract_text(
                page,
                mapping,
                pdf_collection_mode=drag.PDF_MODE_FAST,
                return_details=True,
            )
        document.close()
        self.assertEqual(value, "")
        self.assertEqual(details["mode"], "anchor_unavailable")
        self.assertTrue(details["requires_ocr"])
        ensure.assert_not_called()

    def test_fast_partial_native_row_and_rotated_page_are_excluded(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "partial-and-rotated.pdf"
            document = fitz.open()
            first = document.new_page(width=300, height=200)
            first.insert_text((30, 50), "ONLY-A", fontsize=12)
            second = document.new_page(width=300, height=200)
            second.insert_text((30, 50), "ROTATED", fontsize=12)
            second.set_rotation(90)
            document.save(path)
            document.close()
            mapping = {
                "a": self._mapping((20, 25, 150, 70)),
                "b": self._mapping((170, 25, 290, 70)),
            }
            summary = {}
            failed = []
            with patch.object(
                drag.TEXT_EXTRACTOR,
                "ensure_ocr",
                side_effect=AssertionError("fast mode must not initialize OCR"),
            ) as ensure:
                rows = drag._collect_pdf_rows(
                    [str(path)], ["a", "b"], mapping, False, failed, summary,
                    pdf_collection_mode=drag.PDF_MODE_FAST,
                )
            self.assertEqual(rows, [])
            self.assertEqual(summary["fast_skipped_page_count"], 2)
            self.assertEqual(len(summary["rotated_pages"]), 1)
            self.assertEqual(
                summary["rotated_pages"][0]["reason"],
                drag.ROTATED_PAGE_REQUIRES_OCR,
            )
            self.assertEqual(failed, [path.name])
            ensure.assert_not_called()

    def test_careful_passes_240_and_min256_to_value_ocr(self):
        document = fitz.open()
        page = document.new_page(width=300, height=200)
        mapping = self._mapping((20, 25, 150, 70))
        detected = {
            "boxes": (), "texts": (), "scores": (),
            "width": 100, "height": 50, "clip": (20, 25, 150, 70),
        }
        with patch.object(drag.TEXT_EXTRACTOR, "ensure_ocr", return_value=True), patch.object(
            drag.TEXT_EXTRACTOR, "_ocr_page_detect", return_value=detected
        ), patch.object(
            drag.TEXT_EXTRACTOR,
            "_ocr_value_detect",
            return_value={**detected, "quality": None},
        ) as value_ocr:
            drag.TEXT_EXTRACTOR.extract_text(
                page,
                mapping,
                pdf_collection_mode=drag.PDF_MODE_CAREFUL,
                return_details=True,
            )
        document.close()
        self.assertEqual(value_ocr.call_args.kwargs["dpi"], 240)
        self.assertEqual(value_ocr.call_args.kwargs["limit_side_len"], 256)

    def test_standard_and_careful_value_cache_keys_do_not_mix(self):
        class Result:
            boxes = ()
            txts = ()
            scores = ()

        document = fitz.open()
        page = document.new_page(width=300, height=200)
        rect = fitz.Rect(20, 25, 150, 70)
        with patch.object(drag.TEXT_EXTRACTOR, "ensure_ocr", return_value=True), patch.object(
            drag.TEXT_EXTRACTOR, "_run_value_ocr", return_value=Result()
        ) as run:
            drag.TEXT_EXTRACTOR._ocr_value_detect(page, rect, dpi=210, limit_side_len=224)
            drag.TEXT_EXTRACTOR._ocr_value_detect(page, rect, dpi=240, limit_side_len=256)
            drag.TEXT_EXTRACTOR._ocr_value_detect(page, rect, dpi=240, limit_side_len=256)
        document.close()
        self.assertEqual(run.call_count, 2)

    def test_mcp_schema_and_invalid_external_mode(self):
        tools = {item["name"]: item for item in mcp_server.TOOLS}
        mode = tools["start_pdf_merge"]["inputSchema"]["properties"][
            "pdf_collection_mode"
        ]
        self.assertEqual(mode["enum"], ["fast", "standard", "careful"])
        self.assertEqual(mode["default"], "standard")
        with self.assertRaises(ValueError):
            mcp_server.call("start_pdf_merge", {"pdf_collection_mode": "bad"})


if __name__ == "__main__":
    unittest.main()
