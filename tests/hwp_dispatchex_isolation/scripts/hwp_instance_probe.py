"""운영 코드를 건드리지 않는 HWP COM 격리시험 워커."""

from __future__ import annotations

import argparse
import ctypes
import json
import os
import time
from pathlib import Path

import fitz
import pythoncom
import win32api
import win32com.client as win32
import win32process


PROGID = "HWPFrame.HwpObject"


def create(method):
    if method == "ensure":
        return win32.gencache.EnsureDispatch(PROGID)
    if method == "dispatch":
        return win32.Dispatch(PROGID)
    return win32.DispatchEx(PROGID)


def identity(hwp):
    hwnd = int(hwp.XHwpWindows.Item(0).WindowHandle)
    _thread, pid = win32process.GetWindowThreadProcessId(hwnd)
    handle = win32api.OpenProcess(0x0400 | 0x0010, False, pid)
    try:
        path = win32process.GetModuleFileNameEx(handle, 0)
        created = win32process.GetProcessTimes(handle)["CreationTime"]
    finally:
        handle.Close()
    return {"hwnd": hwnd, "pid": pid, "process_path": path, "process_created": str(created)}


def register_security(hwp):
    try:
        return bool(hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule"))
    except Exception:
        return False


def show_window(hwp):
    try:
        hwnd = int(hwp.XHwpWindows.Item(0).WindowHandle)
        ctypes.windll.user32.ShowWindow(hwnd, 5)
        ctypes.windll.user32.SetForegroundWindow(hwnd)
        return True
    except Exception:
        return False


def export_pdf(hwp, source, target):
    started = time.perf_counter()
    opened = hwp.Open(str(source.resolve()), "", "force")
    hwp.HAction.GetDefault("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet)
    hwp.HParameterSet.HFileOpenSave.filename = str(target.resolve())
    hwp.HParameterSet.HFileOpenSave.Format = "PDF"
    executed = hwp.HAction.Execute(
        "FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet
    )
    if not target.is_file() or target.stat().st_size == 0:
        raise OSError("PDF_OUTPUT_MISSING")
    with fitz.open(target) as document:
        pages = document.page_count
    try:
        hwp.HAction.Run("FileClose")
    except Exception:
        pass
    return {
        "open_result": bool(opened),
        "execute_result": bool(executed),
        "output_bytes": target.stat().st_size,
        "page_count": pages,
        "conversion_seconds": round(time.perf_counter() - started, 3),
    }


def quit_object(hwp):
    started = time.perf_counter()
    hwp.Quit()
    return round(time.perf_counter() - started, 3)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--method", choices=("ensure", "dispatch", "dispatchex"), required=True)
    parser.add_argument("--mode", choices=("single", "double", "hold"), default="single")
    parser.add_argument("--fixture", type=Path, required=True)
    parser.add_argument("--output-dir", type=Path, required=True)
    parser.add_argument("--scenario-id", required=True)
    args = parser.parse_args()
    args.output_dir.mkdir(parents=True, exist_ok=True)
    result = {
        "scenario_id": args.scenario_id,
        "method": args.method,
        "mode": args.mode,
        "worker_pid": os.getpid(),
        "instances": [],
        "success": False,
    }
    objects = []
    pythoncom.CoInitialize()
    try:
        count = 2 if args.mode == "double" else 1
        for index in range(count):
            result["stage"] = "create"
            started = time.perf_counter()
            hwp = create(args.method)
            objects.append(hwp)
            result["stage"] = "identity"
            item = identity(hwp)
            item["create_seconds"] = round(time.perf_counter() - started, 3)
            item["window_shown"] = show_window(hwp)
            item["security_registered"] = register_security(hwp)
            result["stage"] = "export"
            item["conversion"] = export_pdf(
                hwp,
                args.fixture,
                args.output_dir / f"instance_{index + 1}.pdf",
            )
            result["instances"].append(item)
        if args.mode == "hold":
            result["stage"] = "holding"
            print(json.dumps(result, ensure_ascii=False), flush=True)
            command = input().strip()
            result["hold_command"] = command
            if command == "probe":
                result["alive_after_peer_quit"] = int(objects[0].XHwpWindows.Count) >= 0
                print(json.dumps(result, ensure_ascii=False), flush=True)
                command = input().strip()
                result["hold_command"] = command
        result["stage"] = "quit_first"
        result["instances"][0]["quit_seconds"] = quit_object(objects[0])
        objects[0] = None
        if count == 2:
            try:
                result["second_alive_after_first_quit"] = (
                    int(objects[1].XHwpWindows.Count) >= 0
                )
            except Exception:
                result["second_alive_after_first_quit"] = False
            result["instances"][1]["quit_seconds"] = quit_object(objects[1])
            objects[1] = None
        result["success"] = True
        result["stage"] = "complete"
    except Exception as error:
        result["error_type"] = type(error).__name__
        result["error_code"] = getattr(error, "hresult", None)
    finally:
        for hwp in objects:
            if hwp is not None:
                try:
                    hwp.Quit()
                except Exception:
                    pass
        pythoncom.CoUninitialize()
    print(json.dumps(result, ensure_ascii=False))


if __name__ == "__main__":
    main()
