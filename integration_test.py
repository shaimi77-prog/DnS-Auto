"""DnS Auto PDF/OCR 엔진의 자동 통합 점검."""

import json
import re
import time
from pathlib import Path

import fitz
from PIL import Image, ImageDraw, ImageFont

from engine_Drag import HybridTextExtractor, validate_anchor_keyword


ROOT = Path(__file__).resolve().parent
ARTIFACTS = ROOT / "test_artifacts"
ARTIFACTS.mkdir(exist_ok=True)


def create_native_pdf(path):
    doc = fitz.open()
    template = doc.new_page(width=595, height=842)
    template.insert_text((100, 120), "ANCHOR", fontsize=16)
    template.insert_text((205, 120), "VALUE-001", fontsize=16)

    shifted = doc.new_page(width=595, height=842)
    shifted.insert_text((125, 145), "ANCHOR", fontsize=16)
    shifted.insert_text((230, 145), "VALUE-002", fontsize=16)
    doc.save(path)
    doc.close()


def create_scanned_pdf(path):
    image = Image.new("RGB", (1200, 400), "white")
    draw = ImageDraw.Draw(image)
    font_path = Path(r"C:\Windows\Fonts\malgun.ttf")
    font = ImageFont.truetype(str(font_path), 64)
    draw.text((70, 80), "OCR TEST 12345", fill="black", font=font)
    draw.text((70, 190), "한글 인식 시험", fill="black", font=font)
    image_path = path.with_suffix(".png")
    image.save(image_path)

    doc = fitz.open()
    page = doc.new_page(width=600, height=200)
    page.insert_image(page.rect, filename=str(image_path))
    doc.save(path)
    doc.close()


def create_duplicate_anchor_pdf(path):
    doc = fitz.open()
    template = doc.new_page(width=595, height=842)
    template.insert_text((100, 120), "ANCHOR", fontsize=16)
    template.insert_text((250, 120), "ANCHOR", fontsize=16)

    shifted = doc.new_page(width=595, height=842)
    # 첫 페이지에서 두 번째로 가까웠던 x=100 앵커가 다음 페이지에서는
    # x=140으로 이동해 첫 번째 거리순 후보가 되는 역전 상황을 구성합니다.
    shifted.insert_text((40, 140), "ANCHOR", fontsize=16)
    shifted.insert_text((140, 140), "ANCHOR", fontsize=16)
    doc.save(path)
    doc.close()


def main():
    run_id = time.time_ns()
    native_path = ARTIFACTS / f"native_test_{run_id}.pdf"
    scanned_path = ARTIFACTS / f"scanned_test_{run_id}.pdf"
    duplicate_anchor_path = ARTIFACTS / f"duplicate_anchor_{run_id}.pdf"
    create_native_pdf(native_path)
    create_scanned_pdf(scanned_path)
    create_duplicate_anchor_pdf(duplicate_anchor_path)

    extractor = HybridTextExtractor()
    results = {}

    with fitz.open(native_path) as doc:
        drag_rect = fitz.Rect(190, 95, 310, 130)
        mapping = extractor.create_mapping(doc[0], drag_rect, "ANCHOR")
        results["native_template"] = extractor.extract_text(doc[0], mapping)
        results["native_shifted"] = extractor.extract_text(doc[1], mapping)
        adjusted = extractor.adjusted_rect(doc[1], mapping)
        results["anchor_found"] = mapping["anchor_rect"] is not None
        results["anchor_shift"] = [round(adjusted.x0 - drag_rect.x0), round(adjusted.y0 - drag_rect.y0)]

    with fitz.open(duplicate_anchor_path) as doc:
        duplicate_drag_rect = fitz.Rect(190, 95, 310, 130)
        candidates = extractor.find_keyword_candidates(
            doc[0],
            duplicate_drag_rect,
            "ANCHOR",
        )
        selected_mapping = extractor.create_mapping(
            doc[0],
            duplicate_drag_rect,
            "ANCHOR",
            candidates[1],
            1,
        )
        selected_adjusted = extractor.adjusted_rect(doc[1], selected_mapping)
        results["anchor_candidate_count"] = len(candidates)
        results["anchor_candidate_distance_order"] = [
            round(extractor._distance(duplicate_drag_rect, candidate), 2)
            for candidate in candidates
        ]
        results["tracked_anchor_x"] = round(
            selected_mapping["tracking_anchor_rect"].x0
        )
        results["selected_anchor_shift"] = [
            round(selected_adjusted.x0 - duplicate_drag_rect.x0),
            round(selected_adjusted.y0 - duplicate_drag_rect.y0),
        ]
        extractor.reset_mapping_tracking(selected_mapping)
        results["reset_anchor_x"] = round(
            selected_mapping["tracking_anchor_rect"].x0
        )

    spaced_keyword, spaced_keyword_valid = validate_anchor_keyword("  신청인 성명  ")
    long_keyword, long_keyword_valid = validate_anchor_keyword("12345678901")
    results["spaced_keyword"] = spaced_keyword
    results["spaced_keyword_valid"] = spaced_keyword_valid
    results["long_keyword"] = long_keyword
    results["long_keyword_valid"] = long_keyword_valid

    with fitz.open(scanned_path) as doc:
        english_mapping = extractor.create_mapping(doc[0], fitz.Rect(10, 20, 300, 90))
        korean_mapping = extractor.create_mapping(doc[0], fitz.Rect(10, 85, 300, 155))
        results["scanned_native_layer"] = doc[0].get_text("text").strip()
        results["scanned_ocr_english"] = extractor.extract_text(doc[0], english_mapping, force_ocr=True)
        results["scanned_ocr_korean"] = extractor.extract_text(doc[0], korean_mapping, force_ocr=True)
        tiny_mapping = extractor.create_mapping(doc[0], fitz.Rect(1, 1, 3, 3))
        results["tiny_rect"] = extractor.extract_text(doc[0], tiny_mapping, force_ocr=True)

    checks = {
        "native_template": "VALUE-001" in results["native_template"],
        "native_shifted": "VALUE-002" in results["native_shifted"],
        "anchor_found": results["anchor_found"],
        "anchor_shift": results["anchor_shift"] == [25, 25],
        "anchor_candidates_collected": results["anchor_candidate_count"] == 2,
        "anchor_candidates_distance_sorted": (
            results["anchor_candidate_distance_order"]
            == sorted(results["anchor_candidate_distance_order"])
        ),
        "coordinate_anchor_tracking": (
            results["tracked_anchor_x"] == 140
            and results["selected_anchor_shift"] == [40, 20]
        ),
        "tracking_reset_per_pdf": results["reset_anchor_x"] == 100,
        "keyword_middle_space_preserved": (
            results["spaced_keyword"] == "신청인 성명"
            and results["spaced_keyword_valid"]
        ),
        "keyword_over_10_rejected": (
            results["long_keyword"] == "12345678901"
            and not results["long_keyword_valid"]
        ),
        "scan_has_no_native_text": results["scanned_native_layer"] == "",
        # OCR 엔진·모델 버전에 따라 문자 단위 결과는 달라질 수 있다. 빈 결과가 아닌지와
        # 숫자 영역이 실질적으로 읽혔는지를 회귀 기준으로 삼는다.
        "ocr_english": (
            bool(results["scanned_ocr_english"].strip())
            and len(re.findall(r"\d", results["scanned_ocr_english"])) >= 3
        ),
        "ocr_korean": "한글" in results["scanned_ocr_korean"],
        "tiny_rect_guard": results["tiny_rect"] == "",
    }
    report = {"results": results, "checks": checks, "all_passed": all(checks.values())}
    report_path = ARTIFACTS / "integration_report_anchor_patch.json"
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=2))
    raise SystemExit(0 if report["all_passed"] else 1)


if __name__ == "__main__":
    main()
