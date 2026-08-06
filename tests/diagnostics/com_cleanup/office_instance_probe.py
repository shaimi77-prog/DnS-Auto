"""Word/Excel DispatchEx PID 격리시험 워커."""

from __future__ import annotations

import argparse
import json
import os
import sys
from pathlib import Path

import pythoncom
import win32com.client as win32


ROOT = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(ROOT))
from com_process_ownership import capture_processes, confirm_ownership  # noqa: E402


CONFIG = {
    "word": ("Word.Application", "WINWORD.EXE"),
    "excel": ("Excel.Application", "EXCEL.EXE"),
}


def create(application):
    progid, executable = CONFIG[application]
    before = capture_processes([executable])
    app = win32.DispatchEx(progid)
    app.Visible = False
    if application == "word":
        document = app.Documents.Add()
        try:
            hwnd = int(app.ActiveWindow.Hwnd)
        finally:
            document.Close(SaveChanges=False)
    else:
        hwnd = int(app.Hwnd)
    ownership = confirm_ownership(before, hwnd, [executable])
    return app, {
        "pid": ownership.process.pid if ownership.process else None,
        "ownership_status": ownership.status,
        "hwnd": hwnd,
        "version_available": bool(app.Version),
    }


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--application", choices=CONFIG, required=True)
    parser.add_argument("--mode", choices=("double", "hold"), required=True)
    parser.add_argument("--scenario-id", required=True)
    args = parser.parse_args()
    result = {
        "scenario_id": args.scenario_id,
        "application": args.application,
        "worker_pid": os.getpid(),
        "instances": [],
        "success": False,
    }
    apps = []
    pythoncom.CoInitialize()
    try:
        count = 2 if args.mode == "double" else 1
        for _index in range(count):
            app, identity = create(args.application)
            apps.append(app)
            result["instances"].append(identity)
        if args.mode == "hold":
            print(json.dumps(result), flush=True)
            command = input().strip()
            if command == "probe":
                result["alive_after_peer_quit"] = bool(apps[0].Version)
                print(json.dumps(result), flush=True)
                command = input().strip()
            result["command"] = command
        apps[0].Quit()
        apps[0] = None
        if count == 2:
            result["second_alive_after_first_quit"] = bool(apps[1].Version)
            apps[1].Quit()
            apps[1] = None
        result["success"] = True
    except Exception as error:
        result["error_type"] = type(error).__name__
        result["error_code"] = getattr(error, "hresult", None)
    finally:
        for app in apps:
            if app is not None:
                try:
                    app.Quit()
                except Exception:
                    pass
        pythoncom.CoUninitialize()
    print(json.dumps(result), flush=True)


if __name__ == "__main__":
    main()
