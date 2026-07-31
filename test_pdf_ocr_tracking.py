"""스캔 PDF 기준 선정, OCR 앵커 및 회전 제외 로직의 회귀시험."""
# Copyright (C) 2026 두부코드(DOOBOO_CODE)
# SPDX-License-Identifier: AGPL-3.0-only

import json
import tempfile
from pathlib import Path
from types import SimpleNamespace

import fitz

from engine_Drag import (
    OCR_MATCH,
    ROTATED_PAGE_EXCLUDED,
    HybridTextExtractor,
    _collect_pdf_rows,
    select_reference_page,
)


class FakeOCR:
    def __init__(self):
        self.calls = []

    def __call__(self, _image, use_det=None, **_kwargs):
        self.calls.append(bool(use_det))
        if not use_det:
            return SimpleNamespace(txts=("",), scores=(0.0,), boxes=None)
        return SimpleNamespace(
            txts=("기준 단어", "여러 줄 값"),
            scores=(0.95, 0.92),
            boxes=(
                ((20, 20), (120, 20), (120, 50), (20, 50)),
                ((140, 60), (280, 60), (280, 95), (140, 95)),
            ),
        )


class FailingOCR:
    def __call__(self, _image, **_kwargs):
        raise RuntimeError("시험용 OCR 실행 실패")


def create_native_pdf(path):
    document = fitz.open()
    page = document.new_page(width=595, height=842)
    for index in range(5):
        page.insert_text(
            (70, 100 + index * 35),
            f"REFERENCE TEXT BLOCK {index} WITH ENOUGH CHARACTERS",
            fontsize=12,
        )
    document.save(path)
    document.close()


def create_scanned_pdf(path, rotation=0):
    document = fitz.open()
    page = document.new_page(width=595, height=842)
    page.draw_rect(fitz.Rect(40, 40, 500, 400), color=(0, 0, 0))
    if rotation:
        page.set_rotation(rotation)
    document.save(path)
    document.close()


def main():
    checks = {}
    details = {}
    with tempfile.TemporaryDirectory(prefix="dns_auto_ocr_tracking_") as temp:
        root = Path(temp)
        scanned = root / "01_scanned.pdf"
        native = root / "02_native.pdf"
        rotated = root / "03_rotated.pdf"
        create_scanned_pdf(scanned)
        create_native_pdf(native)
        create_scanned_pdf(rotated, rotation=90)

        reference = select_reference_page([str(scanned), str(native)])
        checks["native_reference_preferred"] = (
            Path(reference["pdf_path"]) == native
            and reference["source_type"] == "native"
        )
        scanned_reference = select_reference_page([str(scanned)])
        checks["scanned_reference_fallback"] = (
            Path(scanned_reference["pdf_path"]) == scanned
            and scanned_reference["source_type"] == "scanned"
        )

        extractor = HybridTextExtractor()
        original_ocr = extractor.ocr
        original_reason = extractor.ocr_unavailable_reason
        original_logged = extractor._ocr_unavailable_logged
        original_disabled = extractor._ocr_disabled_for_work
        fake_ocr = FakeOCR()
        extractor.ocr = fake_ocr
        extractor.ocr_unavailable_reason = None
        extractor.reset_work_cache()
        try:
            merged = extractor._merge_ocr_candidates(
                [
                    {
                        "rect": fitz.Rect(10, 10, 100, 40),
                        "score": 0.90,
                        "step": 1,
                    },
                    {
                        "rect": fitz.Rect(12, 11, 102, 41),
                        "score": 0.95,
                        "step": 2,
                    },
                    {
                        "rect": fitz.Rect(180, 10, 260, 40),
                        "score": 0.91,
                        "step": 2,
                    },
                ]
            )
            checks["ocr_candidate_distance_iou_merge"] = (
                len(merged) == 2
                and any(item["score"] == 0.95 for item in merged)
            )
            keyword_rect = extractor._keyword_subrect(
                fitz.Rect(0, 0, 90, 20),
                "prefixkeyword",
                "keyword",
            )
            checks["keyword_subrect"] = (
                keyword_rect.x0 > 0
                and keyword_rect.x1 == 90
            )
            checks["ocr_line_break_preserved"] = (
                extractor._combine_positioned_text(
                    [
                        (fitz.Rect(0, 0, 50, 10), "first"),
                        (fitz.Rect(0, 20, 50, 30), "second"),
                    ]
                )
                == "first\nsecond"
            )
            page_detected = {
                "clip": (0, 0, 100, 100),
                "width": 100,
                "height": 100,
                "boxes": (
                    ((10, 40), (60, 40), (60, 50), (10, 50)),
                ),
                "texts": ("2026. 1. 30.",),
                "scores": (0.93,),
            }
            region_detected = {
                "clip": (0, 0, 100, 100),
                "width": 100,
                "height": 100,
                "boxes": (
                    ((10, 40), (30, 40), (30, 50), (10, 50)),
                    ((30, 40), (40, 40), (40, 50), (30, 50)),
                    ((40, 40), (60, 40), (60, 50), (40, 50)),
                    ((10, 10), (70, 10), (70, 20), (10, 20)),
                ),
                "texts": ("2026.", "1.", "30.", "phone: 000"),
                "scores": (0.99, 0.98, 0.99, 0.99),
            }
            checks["region_noise_excluded"] = (
                extractor._recognized_consensus_text(
                    page_detected,
                    region_detected,
                    fitz.Rect(0, 0, 100, 100),
                )
                == "2026. 1. 30."
            )
            with fitz.open(scanned) as document:
                page = document[0]
                origin = fitz.Rect(0, 0, 300, 200)
                candidates = extractor.find_keyword_candidates(
                    page,
                    origin,
                    "기준 단어",
                )
                checks["ocr_keyword_match"] = (
                    extractor.last_keyword_status == OCR_MATCH
                    and len(candidates) == 1
                )
                mapping = {
                    "rect": fitz.Rect(0, 0, 300, 160),
                    "keyword": "",
                    "anchor_rect": None,
                }
                value = extractor.extract_text(
                    page,
                    mapping,
                    force_ocr=True,
                )
                checks["value_detection_fallback"] = (
                    "여러 줄 값" in value
                    and False not in fake_ocr.calls
                    and fake_ocr.calls.count(True) == 2
                )

            failed_files = []
            summary = {}
            rotated_rows = _collect_pdf_rows(
                [str(rotated)],
                ["항목"],
                {
                    "항목": {
                        "rect": fitz.Rect(0, 0, 300, 160),
                        "keyword": "",
                        "anchor_rect": None,
                    }
                },
                True,
                failed_files,
                summary,
            )
            checks["rotated_page_excluded"] = (
                not rotated_rows
                and len(summary["rotated_pages"]) == 1
                and summary["rotated_pages"][0]["reason"]
                == ROTATED_PAGE_EXCLUDED
            )
            details["summary"] = summary
            extractor.ocr = FailingOCR()
            extractor.ocr_unavailable_reason = None
            extractor._ocr_unavailable_logged = False
            extractor.reset_work_cache()
            with fitz.open(scanned) as document:
                failed_candidates = extractor.find_keyword_candidates(
                    document[0],
                    fitz.Rect(0, 0, 300, 200),
                    "기준 단어",
                )
                checks["ocr_unavailable_distinguished"] = (
                    not failed_candidates
                    and extractor.last_keyword_status == "OCR_UNAVAILABLE"
                    and extractor.last_keyword_reason == "OCR_EXECUTION_ERROR"
                )
        finally:
            extractor.ocr = original_ocr
            extractor.ocr_unavailable_reason = original_reason
            extractor._ocr_unavailable_logged = original_logged
            extractor._ocr_disabled_for_work = original_disabled
            extractor.reset_work_cache()

    failures = [name for name, passed in checks.items() if not passed]
    report = {
        "checks": checks,
        "details": details,
        "all_passed": not failures,
    }
    print(json.dumps(report, ensure_ascii=False, indent=2))
    raise SystemExit(1 if failures else 0)


if __name__ == "__main__":
    main()
