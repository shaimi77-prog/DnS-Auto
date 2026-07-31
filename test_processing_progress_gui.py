"""Smoke test for the shared PDF/Excel processing progress dialog."""

import json
import tkinter as tk

from utils_progress import ProcessingProgressDialog


root = tk.Tk()
root.withdraw()
dialog = ProcessingProgressDialog(
    root,
    "진행 표시 시험",
    3,
    "페이지",
)
for index in range(1, 4):
    dialog.begin_unit(
        "시험.pdf",
        index,
        f"3페이지 중 {index}페이지 | OCR로 텍스트 추출 중",
    )
    dialog.complete_unit(index)

checks = {
    "completed": dialog.completed_units == 3,
    "progress_value": int(dialog.bar["value"]) == 3,
    "file_name": dialog.file_var.get() == "현재 파일: 시험.pdf",
    "detail": "3페이지 중 3페이지" in dialog.detail_var.get(),
    "ocr_activity": "OCR로 텍스트 추출 중" in dialog.detail_var.get(),
    "elapsed": "경과" in dialog.time_var.get(),
    "remaining": "예상 잔여" in dialog.time_var.get(),
}
dialog.close()
root.destroy()
print(json.dumps({"checks": checks, "all_passed": all(checks.values())}, ensure_ascii=False))
raise SystemExit(0 if all(checks.values()) else 1)
