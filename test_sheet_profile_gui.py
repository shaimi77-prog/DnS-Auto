"""Sheet 설정 창에서 프로파일 저장·불러오기와 유형 차단을 검증합니다."""

import json
import tempfile
import tkinter as tk
from pathlib import Path

from openpyxl import Workbook

import engine_Sheet
from utils_profiles import PDF_PROFILE_TYPE, SHEET_PROFILE_TYPE, write_profile


def main():
    checks = {}
    errors = []
    questions = []
    root = tk.Tk()
    root.withdraw()
    original_finish = engine_Sheet.MultiSheetSelector._finish_workbook_load
    original_save_dialog = engine_Sheet.filedialog.asksaveasfilename
    original_open_dialog = engine_Sheet.filedialog.askopenfilename
    original_info = engine_Sheet.messagebox.showinfo
    original_error = engine_Sheet.messagebox.showerror
    original_question = engine_Sheet.messagebox.askyesno

    with tempfile.TemporaryDirectory(prefix="dns_auto_sheet_profile_gui_") as temp_dir:
        temp_root = Path(temp_dir)
        template_path = temp_root / "기준양식.xlsx"
        profile_path = temp_root / "sheet.json"
        wrong_path = temp_root / "pdf.json"
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "총무과"
        worksheet.append(["기관", "금액"])
        worksheet.append(["A", 1])
        workbook.save(template_path)
        workbook.close()

        write_profile(
            {
                "schema_version": 1,
                "profile_type": PDF_PROFILE_TYPE,
                "mapping_sets": [],
            },
            str(wrong_path),
        )

        def finish_and_test(selector, result):
            original_finish(selector, result)
            if result[0] != "ok":
                selector.on_cancel()
                return
            checks["save_button_removed"] = not hasattr(
                selector,
                "save_profile_button",
            )

            row = selector.sheet_vars["총무과"]
            row["check"].set(True)
            selector._selection_changed("총무과")
            row["e"].set("1")
            row["mode"].set(selector.MODE_OPTIONS[1])
            selector._mode_changed("총무과")
            row["key"].set("B")
            row["protect"].set(False)

            engine_Sheet.filedialog.asksaveasfilename = (
                lambda **_kwargs: str(profile_path)
            )
            engine_Sheet.messagebox.showinfo = lambda *_args, **_kwargs: None
            selector.save_profile()
            saved = json.loads(profile_path.read_text(encoding="utf-8"))
            checks["profile_saved"] = (
                saved["profile_type"] == SHEET_PROFILE_TYPE
                and saved["sheet_configs"][0]["key_col"] == "B"
                and saved["sheet_configs"][0]["protect"] is False
            )

            row["check"].set(False)
            selector._selection_changed("총무과")
            row["e"].set("")
            row["mode"].set(selector.MODE_OPTIONS[0])
            row["key"].set(selector.KEY_OPTIONS[0])
            row["protect"].set(True)
            engine_Sheet.filedialog.askopenfilename = (
                lambda **_kwargs: str(profile_path)
            )
            engine_Sheet.messagebox.askyesno = lambda *_args, **_kwargs: True
            selector.load_profile()
            checks["profile_restored"] = (
                row["check"].get()
                and row["e"].get() == "1"
                and row["mode"].get() == selector.MODE_OPTIONS[1]
                and row["key"].get() == "B"
                and row["protect"].get() is False
            )
            checks["profile_label_updated"] = "sheet" in selector.profile_label_var.get()

            engine_Sheet.filedialog.askopenfilename = (
                lambda **_kwargs: str(wrong_path)
            )
            engine_Sheet.messagebox.showerror = (
                lambda _title, message, **_kwargs: errors.append(message)
            )
            selector.load_profile()
            checks["wrong_profile_type_blocked"] = (
                bool(errors)
                and "필요 유형: sheet_config" in errors[-1]
                and row["key"].get() == "B"
            )
            engine_Sheet.messagebox.askyesno = (
                lambda _title, message, **_kwargs: questions.append(message) or True
            )
            selector.top.after(50, selector.on_ok)

        engine_Sheet.MultiSheetSelector._finish_workbook_load = finish_and_test
        try:
            selector = engine_Sheet.MultiSheetSelector(
                root,
                str(template_path),
                ["총무과"],
            )
            checks["selected_settings_returned"] = selector.selected_sheets == {
                "총무과": {
                    "S": 1,
                    "E": 1,
                    "mode": 2,
                    "key_col": "B",
                    "protect": False,
                }
            }
            checks["closed_cleanly"] = selector.workbook is None
            checks["save_prompt_after_completion"] = (
                len(questions) == 1
                and "프로파일로 저장하시겠습니까" in questions[0]
            )
        finally:
            engine_Sheet.MultiSheetSelector._finish_workbook_load = original_finish
            engine_Sheet.filedialog.asksaveasfilename = original_save_dialog
            engine_Sheet.filedialog.askopenfilename = original_open_dialog
            engine_Sheet.messagebox.showinfo = original_info
            engine_Sheet.messagebox.showerror = original_error
            engine_Sheet.messagebox.askyesno = original_question
            root.destroy()

    all_passed = all(checks.values())
    print(
        json.dumps(
            {
                "checks": checks,
                "errors": errors,
                "all_passed": all_passed,
            },
            ensure_ascii=False,
            indent=2,
        )
    )
    raise SystemExit(0 if all_passed else 1)


if __name__ == "__main__":
    main()
