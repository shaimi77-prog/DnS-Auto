"""비업무용 HWPX fixture를 시험 전용 HWP fixture로 변환한다."""

from pathlib import Path
import ctypes

import pythoncom
import win32com.client as win32


ROOT = Path(__file__).resolve().parents[3]
source = ROOT / "tests" / "hwp_dispatchex_isolation" / "fixtures" / "isolation_sample.hwpx"
target = ROOT / "tests" / "hwp_dispatchex_isolation" / "fixtures" / "isolation_sample.hwp"

pythoncom.CoInitialize()
hwp = None
try:
    hwp = win32.DispatchEx("HWPFrame.HwpObject")
    hwnd = int(hwp.XHwpWindows.Item(0).WindowHandle)
    ctypes.windll.user32.ShowWindow(hwnd, 5)
    ctypes.windll.user32.SetForegroundWindow(hwnd)
    try:
        hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
    except Exception:
        pass
    hwp.Open(str(source.resolve()), "", "force")
    saved = hwp.SaveAs(str(target.resolve()), "HWP", "")
    if not target.is_file() or target.stat().st_size == 0:
        raise OSError("HWP_FIXTURE_NOT_CREATED")
    print("HWP_FIXTURE_OK", bool(saved), target.stat().st_size)
finally:
    if hwp is not None:
        hwp.Quit()
    pythoncom.CoUninitialize()
