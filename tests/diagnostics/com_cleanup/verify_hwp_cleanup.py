"""시험용 HWP/HWPX에서 운영과 동일한 DispatchEx·FileClose·정리 계약을 검증한다."""

from __future__ import annotations

import ctypes
import json
import sys
from datetime import datetime
from pathlib import Path

import fitz
import pythoncom
import win32com.client as win32


ROOT = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(ROOT))
from com_process_ownership import Ownership, capture_processes, cleanup_com_session, confirm_ownership  # noqa: E402


def main():
    fixtures = [
        ROOT / "tests" / "hwp_dispatchex_isolation" / "fixtures" / "isolation_sample.hwp",
        ROOT / "tests" / "hwp_dispatchex_isolation" / "fixtures" / "isolation_sample.hwpx",
    ]
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    result_root = ROOT / "tests" / "results" / "com_cleanup" / f"hwp-{stamp}"
    result_root.mkdir(parents=True)
    before = capture_processes(["Hwp.exe"])
    outputs = []
    records = []
    hwp = None
    ownership = Ownership("unconfirmed")
    pythoncom.CoInitialize()
    try:
        hwp = win32.DispatchEx("HWPFrame.HwpObject")
        hwnd = int(hwp.XHwpWindows.Item(0).WindowHandle)
        ctypes.windll.user32.ShowWindow(hwnd, 5)
        ctypes.windll.user32.SetForegroundWindow(hwnd)
        ownership = confirm_ownership(before, hwnd, ["Hwp.exe"])
        try:
            security_registered = bool(hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule"))
        except Exception:
            security_registered = False
        for index, fixture in enumerate(fixtures, start=1):
            target = result_root / f"output_{index}.pdf"
            try:
                opened = bool(hwp.Open(str(fixture.resolve()), "", "force"))
                hwp.HAction.GetDefault("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet)
                hwp.HParameterSet.HFileOpenSave.filename = str(target.resolve())
                hwp.HParameterSet.HFileOpenSave.Format = "PDF"
                executed = bool(hwp.HAction.Execute("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet))
                with fitz.open(target) as document:
                    pages = document.page_count
                outputs.append(target)
                records.append({"fixture_type": fixture.suffix.lower(), "opened": opened, "executed": executed, "bytes": target.stat().st_size, "pages": pages})
            finally:
                hwp.HAction.Run("FileClose")
    finally:
        cleanup = cleanup_com_session(
            application="hwp",
            close_callbacks=[],
            quit_callback=(hwp.Quit if hwp is not None else None),
            ownership=ownership,
            co_uninitialize=pythoncom.CoUninitialize,
            allow_forced_cleanup=True,
        )
    after = capture_processes(["Hwp.exe"])
    payload = {
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "process_count_before": len(before),
        "process_count_after": len(after),
        "security_registered": security_registered,
        "outputs": records,
        "com_cleanup": cleanup,
    }
    (result_root / "summary.json").write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"result_root": str(result_root), **payload}, ensure_ascii=False))


if __name__ == "__main__":
    main()
