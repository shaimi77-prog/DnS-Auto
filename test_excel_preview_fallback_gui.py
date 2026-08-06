"""미리보기 로딩 실패 시에도 엑셀 취합 설정을 완료할 수 있는지 검증합니다."""

import json
import tkinter as tk

import engine_Sheet


def main():
    checks = {}
    warnings = []
    root = tk.Tk()
    root.withdraw()
    original_finish = engine_Sheet.MultiSheetSelector._finish_workbook_load
    original_warning = engine_Sheet.messagebox.showwarning
    original_question = engine_Sheet.messagebox.askyesno

    def finish_and_continue(selector, result):
        original_finish(selector, result)
        checks["load_failed_as_expected"] = result[0] == "error"
        checks["failure_status_visible"] = (
            "미리보기 사용 불가" in selector.preview_status_var.get()
        )
        row = selector.sheet_vars["시험"]
        row["check"].set(True)
        selector._selection_changed("시험")
        row["e"].set("1")
        selector.top.after(50, selector.on_ok)

    engine_Sheet.MultiSheetSelector._finish_workbook_load = finish_and_continue
    engine_Sheet.messagebox.showwarning = (
        lambda title, message, **_kwargs: warnings.append(
            {"title": title, "message": message}
        )
    )
    engine_Sheet.messagebox.askyesno = lambda *_args, **_kwargs: False
    try:
        selector = engine_Sheet.MultiSheetSelector(
            root,
            r"C:\DnS_AI 출품자료\존재하지_않는_미리보기.xlsx",
            ["시험"],
        )
        checks["warning_shown_once"] = len(warnings) == 1
        checks["settings_completed_without_preview"] = (
            selector.selected_sheets.get("시험", {}).get("E") == 1
        )
        checks["closed_cleanly"] = selector.workbook is None
    finally:
        engine_Sheet.MultiSheetSelector._finish_workbook_load = original_finish
        engine_Sheet.messagebox.showwarning = original_warning
        engine_Sheet.messagebox.askyesno = original_question
        root.destroy()

    all_passed = all(checks.values())
    print(json.dumps(
        {"checks": checks, "warnings": warnings, "all_passed": all_passed},
        ensure_ascii=False,
        indent=2,
    ))
    raise SystemExit(0 if all_passed else 1)


if __name__ == "__main__":
    main()
