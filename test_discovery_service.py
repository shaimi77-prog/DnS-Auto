from __future__ import annotations

import json
import tempfile
from pathlib import Path

from openpyxl import Workbook
from services.discovery_service import discover_merge_plan


def workbook(path: Path, sheets=("자료",)) -> None:
    book = Workbook()
    book.active.title = sheets[0]
    for name in sheets[1:]:
        book.create_sheet(name)
    book.save(path)
    book.close()


def main() -> None:
    checks = {}
    with tempfile.TemporaryDirectory(prefix="dns_auto_discovery_") as temp:
        root = Path(temp)
        inputs = root / "inputs"
        nested = inputs / "부서A"
        nested.mkdir(parents=True)
        output = root / "outputs"
        profile_dir = root / "profiles" / "sheet"
        profile_dir.mkdir(parents=True)
        template = inputs / "취합양식.xlsx"
        source = nested / "원본1.xlsx"
        workbook(template)
        workbook(source)
        profile = {
            "schema_version": 1,
            "profile_type": "sheet_config",
            "metadata": {"profile_name": "월별취합", "template_file_name": template.name},
            "sheet_configs": [{"sheet_name": "자료", "start_row": 1, "end_row": 1}],
        }
        (profile_dir / "monthly.json").write_text(json.dumps(profile, ensure_ascii=False), encoding="utf-8")

        plan = discover_merge_plan(root, output)
        checks["recursive_source_search"] = plan["source_candidates"] == ["inputs\\부서A\\원본1.xlsx"]
        checks["operation_detected"] = plan["operation"] == "excel"
        checks["single_template_detected"] = plan["selected_template"] == "inputs\\취합양식.xlsx"
        checks["compatible_profile_requires_confirmation"] = plan["status"] == "needs_confirmation" and len(plan["compatible_profiles"]) == 1
        checks["default_output"] = plan["output_root"] == "outputs"

        explicit = discover_merge_plan(root, output, profile_name="월별취합")
        checks["explicit_profile_ready"] = explicit["status"] == "ready" and explicit["selected_profile"]["name"] == "monthly" and explicit["next_tool"] == "start_sheet_merge"

        interactive = discover_merge_plan(root, output, interactive=True)
        checks["interactive_ready"] = interactive["status"] == "ready" and interactive["next_tool"] == "start_interactive_sheet_merge"

        second = inputs / "다른양식.xlsx"
        workbook(second)
        ambiguous = discover_merge_plan(root, output, operation="excel")
        checks["multiple_templates_not_guessed"] = ambiguous["status"] == "needs_clarification" and len(ambiguous["template_candidates"]) == 2 and any("여러 개" in q for q in ambiguous["questions"])

        missing = discover_merge_plan(root, output, input_root=root / "없음")
        checks["missing_input_asks_path"] = missing["status"] == "needs_clarification" and bool(missing["questions"])

    failed = [name for name, ok in checks.items() if not ok]
    for name, ok in checks.items():
        print(f"{name}: {'PASS' if ok else 'FAIL'}")
    if failed:
        raise AssertionError("실패: " + ", ".join(failed))


if __name__ == "__main__":
    main()