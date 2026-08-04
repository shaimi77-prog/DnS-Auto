"""Approved value-region OCR policy and quality-gate contracts."""
from types import SimpleNamespace
import os
import tempfile
import unittest
from unittest.mock import patch

import fitz
import numpy as np

from engine_Drag import (
    OCR_VALUE_DET_BOX_THRESHOLD,
    OCR_VALUE_DET_LIMIT_SIDE_LEN,
    OCR_VALUE_DET_UNCLIP_RATIO,
    OCR_VALUE_GATE_MAX_GRAY_STD,
    OCR_VALUE_GATE_MAX_LAPLACIAN_VARIANCE,
    OCR_VALUE_RECOGNITION_DPI,
    HybridTextExtractor,
    _apply_value_reinforcement,
    _reference_format_from_values,
)


class RecordingOCR:
    def __init__(self):
        self.calls = []
        self.text_det = SimpleNamespace(limit_type="max", limit_side_len=736)

    def __call__(self, _image, **kwargs):
        self.calls.append(kwargs)
        return SimpleNamespace(
            boxes=(((2, 2), (20, 2), (20, 10), (2, 10)),),
            txts=("safe",),
            scores=(0.99,),
        )


class ValueOcrPolicyTests(unittest.TestCase):
    def setUp(self):
        self.extractor = HybridTextExtractor()
        self.original_ocr = self.extractor.ocr
        self.extractor.ocr = RecordingOCR()
        self.extractor.reset_work_cache()

    def tearDown(self):
        self.extractor.ocr = self.original_ocr
        self.extractor.reset_work_cache()

    def test_approved_constants(self):
        self.assertEqual(OCR_VALUE_RECOGNITION_DPI, 210)
        self.assertEqual(OCR_VALUE_DET_LIMIT_SIDE_LEN, 224)
        self.assertEqual(OCR_VALUE_DET_BOX_THRESHOLD, 0.5)
        self.assertEqual(OCR_VALUE_DET_UNCLIP_RATIO, 1.2)

    def test_gate_requires_both_conditions_and_includes_boundary(self):
        gate = self.extractor._value_enhancement_gate
        self.assertTrue(gate({"gray_std": OCR_VALUE_GATE_MAX_GRAY_STD,
                              "laplacian_variance": OCR_VALUE_GATE_MAX_LAPLACIAN_VARIANCE}))
        self.assertFalse(gate({"gray_std": OCR_VALUE_GATE_MAX_GRAY_STD + 0.001,
                               "laplacian_variance": OCR_VALUE_GATE_MAX_LAPLACIAN_VARIANCE}))
        self.assertFalse(gate({"gray_std": OCR_VALUE_GATE_MAX_GRAY_STD,
                               "laplacian_variance": OCR_VALUE_GATE_MAX_LAPLACIAN_VARIANCE + 0.001}))

    def test_unsharp_preserves_shape_dtype_and_input(self):
        image = np.full((24, 80, 3), 220, dtype=np.uint8)
        image[8:16, 20:60] = 80
        original = image.copy()
        enhanced = self.extractor._enhance_value_image(image)
        self.assertEqual(enhanced.shape, image.shape)
        self.assertEqual(enhanced.dtype, image.dtype)
        np.testing.assert_array_equal(image, original)
        self.assertFalse(np.array_equal(enhanced, image))

    def test_value_policy_cache_and_ocr_options(self):
        document = fitz.open()
        try:
            page = document.new_page(width=120, height=60)
            page.draw_rect(fitz.Rect(5, 5, 100, 40), fill=(0.8, 0.8, 0.8))
            clip = fitz.Rect(0, 0, 110, 50)
            first = self.extractor._ocr_value_detect(page, clip)
            second = self.extractor._ocr_value_detect(page, clip)
            enhanced = self.extractor._ocr_value_detect(page, clip, enhance=True)
            self.assertIs(first, second)
            self.assertIsNotNone(enhanced)
            self.assertEqual(len(self.extractor.ocr.calls), 2)
            for call in self.extractor.ocr.calls:
                self.assertTrue(call["use_det"])
                self.assertTrue(call["use_rec"])
                self.assertFalse(call["use_cls"])
                self.assertEqual(call["box_thresh"], 0.5)
                self.assertEqual(call["unclip_ratio"], 1.2)
            self.assertEqual(self.extractor.ocr.text_det.limit_type, "max")
            self.assertEqual(self.extractor.ocr.text_det.limit_side_len, 736)
            stats = self.extractor.ocr_statistics()
            self.assertEqual(stats["value_primary_ocr_count"], 1)
            self.assertEqual(stats["value_enhancement_ocr_count"], 1)
            self.assertEqual(stats["total_ocr_inference_count"], 2)
        finally:
            document.close()

    def test_cell_format_is_one_value_across_multiple_lines(self):
        value = "홍길동\n서울지방검찰청\n사건번호 2026-123\n해당 없음"
        self.assertEqual(
            self.extractor.classify_value_format(value),
            "LETTERS_DIGITS",
        )

    def test_format_ignores_selected_punctuation_without_changing_value(self):
        cases = {
            "010-1234-5678": "DIGITS",
            "02.123.4567": "DIGITS",
            "1,234": "DIGITS",
            "---": "EMPTY",
            "()": "SYMBOLS",
            "문자()": "LETTERS_SYMBOLS",
            "123()": "DIGITS_SYMBOLS",
            "문자123()": "LETTERS_DIGITS_SYMBOLS",
        }
        for value, expected in cases.items():
            with self.subTest(value=value):
                self.assertEqual(
                    self.extractor.classify_value_format(value), expected
                )
                self.assertEqual(value, value)

    def test_reference_format_has_no_floor_but_needs_comparison_for_ocr(self):
        self.assertEqual(
            _reference_format_from_values(["native"], allow_single=True),
            "LETTERS",
        )
        self.assertIsNone(
            _reference_format_from_values(["123"], allow_single=False)
        )
        self.assertEqual(
            _reference_format_from_values(
                ["123", "456", "문자"], allow_single=False
            ),
            "DIGITS",
        )
        self.assertIsNone(
            _reference_format_from_values(
                ["123", "문자", "()"], allow_single=False
            )
        )

    def test_reinforcement_selection_contract(self):
        select = self.extractor.select_reinforced_value
        self.assertEqual(
            select("서울지방검찰청", "지방검찰청", "LETTERS")[0],
            "서울지방검찰청",
        )
        self.assertEqual(
            select("사건번호", "사건번호 2026-123", "LETTERS_DIGITS")[0],
            "사건번호 2026-123",
        )
        self.assertEqual(
            select("박진석", "박진수", "LETTERS")[0],
            "박진석",
        )
        self.assertEqual(select("7", "홍길동", "LETTERS")[0], "홍길동")
        self.assertEqual(select("문자", "123", "DIGITS")[0], "123")
        self.assertEqual(select("", "123", "DIGITS")[0], "123")
        self.assertEqual(select("", "문자", "DIGITS")[0], "")
        self.assertEqual(select("abc", "abc123", None)[0], "abc")
    def test_second_pass_rerenders_only_outlier_and_keeps_no_image(self):
        with tempfile.TemporaryDirectory() as temporary:
            pdf_path = os.path.join(temporary, "sample.pdf")
            document = fitz.open()
            document.new_page(width=120, height=60)
            document.save(pdf_path)
            document.close()
            records = []
            for value in ("사건번호 2026", "사건번호 2027", "사건번호"):
                records.append({
                    "pdf_path": pdf_path,
                    "filename": "sample.pdf",
                    "page_index": 0,
                    "values": {"헤더": value},
                    "details": {"헤더": {
                        "mode": "ocr",
                        "primary_text": value,
                        "enhanced_text": None,
                        "low_quality_gate": False,
                        "rect": (5, 5, 100, 40),
                    }},
                })
            detected = object()
            with patch.object(
                self.extractor, "_ocr_value_detect", return_value=detected
            ) as detect, patch.object(
                self.extractor,
                "_recognized_consensus_text",
                return_value="사건번호 2028",
            ):
                _apply_value_reinforcement(
                    records,
                    ["헤더"],
                    {"헤더": "LETTERS_DIGITS"},
                )
            self.assertEqual(detect.call_count, 1)
            self.assertTrue(detect.call_args.kwargs["enhance"])
            self.assertEqual(records[2]["values"]["헤더"], "사건번호 2028")
            for record in records:
                self.assertNotIn("image", record)
                self.assertNotIn("image", record["details"]["헤더"])

    def test_empty_with_ink_retries_384_then_736(self):
        with tempfile.TemporaryDirectory() as temporary:
            pdf_path = os.path.join(temporary, "empty.pdf")
            document = fitz.open()
            document.new_page(width=120, height=60)
            document.save(pdf_path)
            document.close()
            record = {
                "pdf_path": pdf_path,
                "filename": "empty.pdf",
                "page_index": 0,
                "values": {"헤더": ""},
                "details": {"헤더": {
                    "mode": "ocr", "primary_text": "",
                    "enhanced_text": "", "low_quality_gate": True,
                    "rect": (5, 5, 100, 40),
                }},
            }
            detections = [SimpleNamespace(limit=384), SimpleNamespace(limit=736)]
            with patch.object(
                self.extractor,
                "_value_region_has_character_ink",
                return_value=True,
            ), patch.object(
                self.extractor,
                "_ocr_value_detect",
                side_effect=detections,
            ) as detect, patch.object(
                self.extractor,
                "_recognized_region_text",
                side_effect=["", "123"],
            ):
                _apply_value_reinforcement(
                    [record], ["헤더"], {"헤더": "DIGITS"}
                )
            self.assertEqual(detect.call_count, 2)
            self.assertEqual(
                detect.call_args_list[0].kwargs["limit_side_len"], 384
            )
            self.assertEqual(
                detect.call_args_list[1].kwargs["limit_side_len"], 736
            )
            self.assertEqual(record["values"]["헤더"], "123")
    def test_reset_clears_cache_and_statistics(self):
        self.extractor._ocr_cache[("private",)] = {"texts": ("not logged",)}
        self.extractor._ocr_statistics["value_primary_ocr_count"] = 3
        self.extractor.reset_work_cache()
        self.assertEqual(self.extractor._ocr_cache, {})
        self.assertEqual(self.extractor.ocr_statistics()["value_primary_ocr_count"], 0)



if __name__ == "__main__":
    unittest.main()
