"""Sheet 설정 프로파일의 스키마, 유형 분리 및 안전 저장을 검증합니다."""

import json
import tempfile
from pathlib import Path

from engine_Sheet import MultiSheetSelector
from utils_profiles import (
    PDF_PROFILE_TYPE,
    SHEET_PROFILE_TYPE,
    profile_directory,
    read_profile,
    write_profile,
)


def main():
    checks = {}
    selector = MultiSheetSelector.__new__(MultiSheetSelector)
    selector.template_path = "기준양식.xlsx"
    selector.sheet_vars = {"총무과": {}, "보안과": {}}

    selected = {
        "총무과": {
            "S": 1,
            "E": 2,
            "mode": 1,
            "key_col": "",
            "protect": True,
        },
        "보안과": {
            "S": 3,
            "E": 4,
            "mode": 2,
            "key_col": "A",
            "protect": False,
        },
    }
    profile = selector._profile_document(selected)
    checks["sheet_type_written"] = profile["profile_type"] == SHEET_PROFILE_TYPE
    checks["sheet_settings_written"] = (
        profile["sheet_configs"][0]["sheet_name"] == "총무과"
        and profile["sheet_configs"][1]["key_col"] == "A"
        and profile["sheet_configs"][1]["protect"] is False
    )
    fatal, minor = selector._validate_profile(profile)
    checks["valid_sheet_profile"] = not fatal and not minor

    wrong_type = json.loads(json.dumps(profile, ensure_ascii=False))
    wrong_type["profile_type"] = PDF_PROFILE_TYPE
    fatal, _minor = selector._validate_profile(wrong_type)
    checks["wrong_type_rejected_by_validator"] = any(
        "Sheet 설정 프로파일이 아닙니다" in reason for reason in fatal
    )

    missing_sheet = json.loads(json.dumps(profile, ensure_ascii=False))
    missing_sheet["sheet_configs"][0]["sheet_name"] = "없는시트"
    fatal, _minor = selector._validate_profile(missing_sheet)
    checks["missing_sheet_rejected"] = any(
        "기준양식에 없습니다" in reason for reason in fatal
    )

    checks["profile_directories_separated"] = (
        Path(profile_directory(PDF_PROFILE_TYPE)).name == "pdf"
        and Path(profile_directory(SHEET_PROFILE_TYPE)).name == "sheet"
    )

    with tempfile.TemporaryDirectory(prefix="dns_auto_sheet_profile_") as temp_dir:
        sheet_path = Path(temp_dir) / "sheet.json"
        pdf_path = Path(temp_dir) / "pdf.json"
        write_profile(profile, str(sheet_path))
        loaded, legacy = read_profile(str(sheet_path), SHEET_PROFILE_TYPE)
        checks["sheet_roundtrip"] = loaded == profile and legacy is False
        checks["temporary_file_cleaned"] = not Path(f"{sheet_path}.tmp").exists()

        write_profile(
            {
                "schema_version": 1,
                "profile_type": PDF_PROFILE_TYPE,
                "mapping_sets": [],
            },
            str(pdf_path),
        )
        try:
            read_profile(str(pdf_path), SHEET_PROFILE_TYPE)
            checks["pdf_blocked_in_sheet_loader"] = False
        except ValueError as error:
            checks["pdf_blocked_in_sheet_loader"] = (
                "필요 유형: sheet_config" in str(error)
                and "선택 유형: pdf_mapping" in str(error)
            )

        legacy_pdf_path = Path(temp_dir) / "legacy_pdf.json"
        write_profile(
            {"schema_version": 1, "mapping_sets": [{"group": "개별"}]},
            str(legacy_pdf_path),
        )
        _legacy_profile, legacy = read_profile(
            str(legacy_pdf_path),
            PDF_PROFILE_TYPE,
            allow_legacy_pdf=True,
        )
        checks["legacy_pdf_accepted_with_flag"] = legacy is True

    failures = [name for name, passed in checks.items() if not passed]
    print(
        json.dumps(
            {"checks": checks, "all_passed": not failures},
            ensure_ascii=False,
            indent=2,
        )
    )
    raise SystemExit(1 if failures else 0)


if __name__ == "__main__":
    main()
