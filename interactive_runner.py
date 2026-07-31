"""Run the existing Tk GUI as an MCP-controlled interactive child process."""
from __future__ import annotations

import json
import os
import shutil
import sys
import traceback
from pathlib import Path
from tkinter import filedialog


def _output_candidates(folder: Path, mode: str) -> set[Path]:
    pattern = "DnS_Auto_Sheet다중취합_*" if mode == "sheet" else "DnS_Auto_Drag취합_*"
    return {path.resolve() for path in folder.glob(pattern) if path.is_file()}


def _unique_destination(output_root: Path, source: Path) -> Path:
    target = output_root / source.name
    counter = 2
    while target.exists():
        target = output_root / f"{source.stem}_{counter}{source.suffix}"
        counter += 1
    return target


def run_request(request_path: str) -> int:
    request_file = Path(request_path).resolve()
    request = json.loads(request_file.read_text(encoding="utf-8"))
    result_file = Path(request["result_path"]).resolve()
    mode = request["mode"]
    template = Path(request["template_path"]).resolve()
    output_root = Path(request["output_root"]).resolve()
    output_root.mkdir(parents=True, exist_ok=True)
    before = _output_candidates(template.parent, mode)
    result = {"state": "failed", "output_files": [], "failed_files": [], "message": "사용자가 대화형 설정을 취소했습니다.", "details": {"interactive": True, "cancelled": True}}

    import tkinter as tk
    import engine_Drag
    import engine_Sheet

    root = tk.Tk()
    root.withdraw()
    original_one = filedialog.askopenfilename
    original_many = filedialog.askopenfilenames
    original_startfile = getattr(os, "startfile", None)
    source_paths = tuple(request.get("source_paths", ()))
    pdfs_by_sheet = request.get("pdfs_by_sheet", {})

    def choose_one(*args, **kwargs):
        title = str(kwargs.get("title", ""))
        if "템플릿" in title or "양식" in title:
            return str(template)
        return original_one(*args, **kwargs)

    def choose_many(*args, **kwargs):
        title = str(kwargs.get("title", ""))
        if mode == "sheet" and source_paths:
            return source_paths
        if mode == "pdf":
            for sheet_name, paths in pdfs_by_sheet.items():
                if f"[{sheet_name}]" in title:
                    return tuple(paths)
        return original_many(*args, **kwargs)

    filedialog.askopenfilename = choose_one
    filedialog.askopenfilenames = choose_many
    if original_startfile is not None:
        os.startfile = lambda path: None
    try:
        if mode == "sheet":
            engine_Sheet.run_application(root)
        elif mode == "pdf":
            engine_Drag.run_application(root, force_ocr=bool(request.get("force_ocr")))
        else:
            raise ValueError(f"지원하지 않는 대화형 작업입니다: {mode}")
        created = sorted(_output_candidates(template.parent, mode) - before, key=lambda path: path.stat().st_mtime_ns)
        if created:
            outputs = []
            for source in created:
                target = _unique_destination(output_root, source)
                shutil.move(str(source), str(target))
                outputs.append(str(target.resolve()))
            result = {"state": "succeeded", "output_files": outputs, "failed_files": [], "message": "사용자 설정을 반영한 대화형 취합이 완료되었습니다.", "details": {"interactive": True, "cancelled": False}}
    except Exception as error:
        result = {"state": "failed", "output_files": [], "failed_files": [], "message": str(error), "details": {"interactive": True, "cancelled": False, "traceback": traceback.format_exc()}}
    finally:
        filedialog.askopenfilename = original_one
        filedialog.askopenfilenames = original_many
        if original_startfile is not None:
            os.startfile = original_startfile
        try:
            root.destroy()
        except Exception:
            pass
        result_file.write_text(json.dumps(result, ensure_ascii=False), encoding="utf-8")
    return 0 if result["state"] == "succeeded" else 2


if __name__ == "__main__":
    raise SystemExit(run_request(sys.argv[1]))