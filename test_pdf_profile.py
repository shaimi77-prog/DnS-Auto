"""PDF 설정 그리드 공통 로직과 매핑 프로파일의 비대화형 회귀시험."""

import json
import tempfile
from pathlib import Path

import fitz
from openpyxl import Workbook

from engine_Drag import (
    EXCEL_MAX_ROW,
    SheetGroupSelector,
    _build_mapping_sets,
    _extract_headers,
    _mapping_from_json,
    _mapping_to_json,
    _preview_bounds,
)
from utils_profiles import (
    PDF_PROFILE_TYPE,
    SHEET_PROFILE_TYPE,
    read_profile,
    write_profile,
)


def main():
    checks = {}

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "통계"
    worksheet.merge_cells("A1:A2")
    worksheet["A1"] = "기관"
    worksheet["B1"] = "처리"
    worksheet["B2"] = "결과"
    headers = _extract_headers(worksheet, 1, 2)
    checks["merged_header_read"] = headers[:2] == ["기관", "처리_결과"]

    sets = _build_mapping_sets(
        {
            "4월": {"group": "그룹 1", "S": 2, "E": 3},
            "5월": {"group": "그룹 1", "S": 2, "E": 3},
            "통계": {"group": "개별", "S": 4, "E": 5},
        }
    )
    checks["group_shared"] = sets[0]["sheets"] == ["4월", "5월"]
    checks["individual_separate"] = sets[1]["group"] == "개별" and sets[1]["sheets"] == ["통계"]

    absolute_mapping = {
        "rect": fitz.Rect(10, 20, 30, 40),
        "keyword": "",
        "anchor_rect": None,
        "tracking_anchor_rect": None,
        "offset_x": 0,
        "offset_y": 0,
    }
    absolute_json = _mapping_to_json("절대좌표", 1, absolute_mapping)
    absolute_restored = _mapping_from_json(absolute_json)
    checks["anchorless_null_schema"] = (
        absolute_json["keyword"] is None
        and absolute_json["anchor_rect"] is None
        and absolute_json["anchor_offset"] is None
        and absolute_restored["anchor_rect"] is None
    )

    anchored_mapping = {
        "rect": fitz.Rect(100, 120, 180, 150),
        "keyword": "성 명",
        "anchor_rect": fitz.Rect(50, 120, 90, 145),
        "tracking_anchor_rect": fitz.Rect(50, 120, 90, 145),
        "offset_x": 50,
        "offset_y": 0,
    }
    anchored_json = _mapping_to_json("성명", 2, anchored_mapping)
    anchored_restored = _mapping_from_json(anchored_json)
    checks["anchor_roundtrip"] = (
        anchored_restored["keyword"] == "성 명"
        and anchored_restored["offset_x"] == 50
        and list(anchored_restored["anchor_rect"]) == [50.0, 120.0, 90.0, 145.0]
    )
    broken_anchor = dict(anchored_json)
    broken_anchor["anchor_offset"] = None
    try:
        _mapping_from_json(broken_anchor)
        checks["partial_anchor_rejected"] = False
    except ValueError:
        checks["partial_anchor_rejected"] = True

    checks["preview_clamps_start"] = _preview_bounds(5, 20) == (1, 30)
    checks["preview_shows_empty_tail"] = _preview_bounds(20, 25) == (10, 35)
    checks["preview_clamps_excel_end"] = _preview_bounds(
        EXCEL_MAX_ROW - 2,
        EXCEL_MAX_ROW,
    ) == (EXCEL_MAX_ROW - 12, EXCEL_MAX_ROW)

    profile = {
        "schema_version": 1,
        "profile_type": PDF_PROFILE_TYPE,
        "metadata": {"profile_name": "시험"},
        "mapping_sets": [
            {
                "group": "개별",
                "sheets": ["통계"],
                "header_start": 4,
                "header_end": 5,
                "pdf": {
                    "page_width": 595,
                    "page_height": 842,
                    "rotation": 0,
                },
                "fields": [absolute_json, anchored_json],
            }
        ],
    }
    validator = SheetGroupSelector.__new__(SheetGroupSelector)
    validator.workbook = workbook
    validator.template_path = "template.xlsx"
    fatal, _minor = validator._validate_profile_structure(profile)
    checks["valid_profile_structure"] = not fatal
    invalid_profile = json.loads(json.dumps(profile, ensure_ascii=False))
    invalid_profile["mapping_sets"][0]["group"] = "A"
    fatal, _minor = validator._validate_profile_structure(invalid_profile)
    checks["invalid_group_rejected"] = any("그룹 값" in reason for reason in fatal)
    with tempfile.TemporaryDirectory(prefix="dns_auto_profile_test_") as temp_dir:
        path = Path(temp_dir) / "시험.json"
        write_profile(profile, str(path))
        restored_profile = json.loads(path.read_text(encoding="utf-8"))
        checks["atomic_profile_write"] = restored_profile == profile
        checks["temp_file_cleaned"] = not Path(f"{path}.tmp").exists()

        legacy_profile = json.loads(json.dumps(profile, ensure_ascii=False))
        legacy_profile.pop("profile_type")
        legacy_path = Path(temp_dir) / "레거시_pdf.json"
        write_profile(legacy_profile, str(legacy_path))
        restored_legacy, legacy_detected = read_profile(
            str(legacy_path),
            PDF_PROFILE_TYPE,
            allow_legacy_pdf=True,
        )
        checks["legacy_pdf_accepted"] = (
            legacy_detected
            and restored_legacy == legacy_profile
        )
        try:
            read_profile(str(legacy_path), PDF_PROFILE_TYPE)
            checks["legacy_pdf_requires_opt_in"] = False
        except ValueError:
            checks["legacy_pdf_requires_opt_in"] = True

        missing_type_path = Path(temp_dir) / "유형_및_매핑없음.json"
        write_profile({"schema_version": 1}, str(missing_type_path))
        try:
            read_profile(
                str(missing_type_path),
                PDF_PROFILE_TYPE,
                allow_legacy_pdf=True,
            )
            checks["invalid_legacy_rejected"] = False
        except ValueError:
            checks["invalid_legacy_rejected"] = True

        sheet_profile = {
            "schema_version": 1,
            "profile_type": SHEET_PROFILE_TYPE,
            "sheet_configs": [],
        }
        sheet_path = Path(temp_dir) / "sheet.json"
        write_profile(sheet_profile, str(sheet_path))
        try:
            read_profile(str(sheet_path), PDF_PROFILE_TYPE)
            checks["sheet_profile_rejected_by_pdf"] = False
        except ValueError:
            checks["sheet_profile_rejected_by_pdf"] = True
        try:
            read_profile(str(path), SHEET_PROFILE_TYPE)
            checks["pdf_profile_rejected_by_sheet"] = False
        except ValueError:
            checks["pdf_profile_rejected_by_sheet"] = True

    workbook.close()
    failures = [name for name, passed in checks.items() if not passed]
    print(json.dumps({"checks": checks, "all_passed": not failures}, ensure_ascii=False, indent=2))
    raise SystemExit(1 if failures else 0)


if __name__ == "__main__":
    main()
