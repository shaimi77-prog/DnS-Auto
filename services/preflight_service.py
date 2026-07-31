"""MCP 실행 전 파일·경로·프로필을 읽기 전용으로 점검합니다."""

from __future__ import annotations

import os
from pathlib import Path
from typing import Iterable


SUPPORTED_EXTENSIONS = {".pdf", ".xlsx", ".xlsm", ".xls", ".doc", ".docx", ".hwp", ".hwpx", ".json"}


def inspect_paths(paths: Iterable[str], allowed_roots: Iterable[str] | None = None) -> list[dict]:
    roots = [Path(root).resolve() for root in allowed_roots or ()]
    findings = []
    for raw_path in paths:
        path = Path(raw_path).expanduser()
        item = {"path": str(path), "ok": False, "issues": []}
        if not path.is_absolute():
            item["issues"].append("절대 경로가 아닙니다.")
        elif path.suffix.lower() not in SUPPORTED_EXTENSIONS:
            item["issues"].append(f"지원하지 않는 확장자입니다: {path.suffix}")
        elif roots and not any(_is_within(path, root) for root in roots):
            item["issues"].append("허용 입력 폴더 밖의 경로입니다.")
        elif not path.is_file():
            item["issues"].append("존재하는 일반 파일이 아닙니다.")
        else:
            item.update({"ok": True, "size_bytes": path.stat().st_size, "extension": path.suffix.lower()})
        findings.append(item)
    return findings


def _is_within(path: Path, root: Path) -> bool:
    try:
        path.resolve().relative_to(root)
        return True
    except ValueError:
        return False
