"""Automated modal/lifecycle smoke test for conversion progress."""

import json
import tkinter as tk

from utils_converter import ProgressWindow


root = tk.Tk()
dialog = ProgressWindow(root, "변환 진행 시험", 2)
checks = {
    "transient": str(dialog.top.transient()) == str(root),
    "grabbed": dialog.top.grab_current() == dialog.top,
    "registered": getattr(root, "_dns_active_progress", None) is dialog,
}
dialog.update_progress(1, "시험.docx")
dialog.close()
checks["unregistered"] = not hasattr(root, "_dns_active_progress")
checks["closed"] = dialog._closed
root.destroy()
print(json.dumps({"checks": checks, "all_passed": all(checks.values())}, ensure_ascii=False))
raise SystemExit(0 if all(checks.values()) else 1)
