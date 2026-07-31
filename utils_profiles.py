"""DnS Auto 프로파일의 경로, 유형 검증 및 안전 저장을 관리합니다."""
# Copyright (C) 2026 두부코드(DOOBOO_CODE)
# SPDX-License-Identifier: AGPL-3.0-only

import json
import os
import sys


PROFILE_SCHEMA_VERSION = 1
PDF_PROFILE_TYPE = "pdf_mapping"
SHEET_PROFILE_TYPE = "sheet_config"
PROFILE_SUBDIRECTORIES = {
    PDF_PROFILE_TYPE: "pdf",
    SHEET_PROFILE_TYPE: "sheet",
}


def application_dir():
    """onefile EXE와 소스 실행 환경의 프로그램 기준 경로를 반환합니다."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def profiles_root():
    return os.path.join(application_dir(), "profiles")


def profile_directory(profile_type):
    try:
        subdirectory = PROFILE_SUBDIRECTORIES[profile_type]
    except KeyError as error:
        raise ValueError(f"지원하지 않는 프로파일 유형입니다: {profile_type}") from error
    return os.path.join(profiles_root(), subdirectory)


def prepare_profile_directory(profile_type):
    path = profile_directory(profile_type)
    os.makedirs(path, exist_ok=True)
    if not os.access(path, os.W_OK):
        raise PermissionError(f"쓰기 권한이 없습니다: {path}")
    return path


def read_profile(path, expected_type, allow_legacy_pdf=False):
    """JSON을 읽고 요청한 프로파일 유형인지 확인합니다."""
    with open(path, "r", encoding="utf-8") as profile_file:
        profile = json.load(profile_file)
    if not isinstance(profile, dict):
        raise ValueError("프로파일 최상위 구조가 객체가 아닙니다.")

    actual_type = profile.get("profile_type")
    legacy_pdf = (
        allow_legacy_pdf
        and expected_type == PDF_PROFILE_TYPE
        and actual_type is None
        and "mapping_sets" in profile
    )
    if actual_type != expected_type and not legacy_pdf:
        readable_actual = actual_type or "유형 정보 없음"
        raise ValueError(
            "선택한 파일은 현재 기능에서 사용할 수 없는 프로파일입니다.\n"
            f"필요 유형: {expected_type}\n"
            f"선택 유형: {readable_actual}"
        )
    return profile, legacy_pdf


def write_profile(profile, path):
    """임시 파일 검증 후 원자적으로 JSON 프로파일을 교체합니다."""
    temp_path = f"{path}.tmp"
    try:
        with open(temp_path, "w", encoding="utf-8") as profile_file:
            json.dump(profile, profile_file, ensure_ascii=False, indent=2)
            profile_file.flush()
            os.fsync(profile_file.fileno())
        with open(temp_path, "r", encoding="utf-8") as profile_file:
            json.load(profile_file)
        os.replace(temp_path, path)
    finally:
        if os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except OSError:
                pass
