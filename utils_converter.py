"""Word, HWP 및 구버전 Excel 파일의 일괄 변환 기능."""

from __future__ import annotations

import ctypes
import logging
import os
from pathlib import Path
import threading
import time
import tkinter as tk
from tkinter import filedialog, messagebox

import pythoncom
import win32com.client as win32

from config import PROGRAM_NAME
from core.models import JobState
from processing_cancellation import ProcessingCancellation
from services import conversion_service


class ProgressWindow:
    """모달 변환 상태창과 안전한 전체 취소 요청을 관리한다."""

    def __init__(self, parent, title, total, cancellation=None):
        self.parent = parent
        self.cancellation = cancellation or ProcessingCancellation()
        self.top = tk.Toplevel(parent)
        self.top.title(title)
        self.top.geometry("430x210")
        self.top.resizable(False, False)
        self.top.transient(parent)
        self.start_time = time.time()
        self.total = max(int(total), 1)
        self.current = 0
        self._closed = False
        self._prompt_open = False
        self._generation = 0
        self.top.protocol("WM_DELETE_WINDOW", self._on_close_request)

        self.title_var = tk.StringVar(value=f"현재 작업: {title}")
        self.file_var = tk.StringVar(value="현재 파일: 준비 중")
        self.status_var = tk.StringVar(value=f"0/{self.total} · 예상 잔여 시간 계산 중")
        tk.Label(self.top, textvariable=self.title_var, font=("맑은 고딕", 10, "bold")).pack(
            pady=(20, 8)
        )
        tk.Label(self.top, textvariable=self.file_var, font=("맑은 고딕", 10)).pack(pady=3)
        tk.Label(self.top, textvariable=self.status_var, font=("맑은 고딕", 10)).pack(pady=3)
        self.canvas = tk.Canvas(
            self.top, width=360, height=15, bg="#E0E0E0", highlightthickness=0
        )
        self.canvas.pack(pady=12)
        self.bar = self.canvas.create_rectangle(0, 0, 0, 15, fill="#2E7D32")
        self.top.update_idletasks()
        width, height = self.top.winfo_width(), self.top.winfo_height()
        screen_width, screen_height = (
            self.top.winfo_screenwidth(),
            self.top.winfo_screenheight(),
        )
        self.top.geometry(
            f"{width}x{height}+{(screen_width-width)//2}+{(screen_height-height)//2}"
        )
        self.top.grab_set()
        setattr(self.parent, "_dns_active_progress", self)
        self.top.update()

    @property
    def is_cancelled(self):
        return self.cancellation.should_cancel()

    def _on_close_request(self):
        if self._prompt_open or self.cancellation.save_started:
            return
        self._prompt_open = True
        try:
            answer = messagebox.askyesno(
                "전체 작업 취소",
                "전체 작업을 취소하시겠습니까?\n\n"
                "현재 처리 중인 파일 작업이 안전하게 끝난 뒤 취소됩니다.\n"
                "이번 작업에서 생성한 결과 파일은 저장되지 않습니다.",
                parent=self.top,
            )
            if answer and self.cancellation.request_cancel_all():
                self.title_var.set("전체 작업 취소를 요청했습니다.")
                self.file_var.set("현재 작업을 안전하게 마친 뒤 종료합니다…")
                self.status_var.set("")
                self.top.protocol("WM_DELETE_WINDOW", lambda: None)
        finally:
            self._prompt_open = False

    def update_progress(self, current_num, filename="", activity="변환 중"):
        if self._closed or self.is_cancelled:
            return
        try:
            if not self.top.winfo_exists():
                return
            self.current = int(current_num)
            self.file_var.set(f"현재 파일: {filename}")
            elapsed = time.time() - self.start_time
            if self.current > 0:
                remaining = elapsed / self.current * (self.total - self.current)
                remaining_text = (
                    "잠시만 기다려 주세요"
                    if remaining < 60
                    else f"약 {int(remaining // 60)}분 {int(remaining % 60)}초"
                )
            else:
                remaining_text = "계산 중"
            self.status_var.set(
                f"{self.total}개 파일 중 {self.current}번째 · {activity}\n"
                f"예상 잔여 시간 {remaining_text}"
            )
            self.canvas.coords(
                self.bar, 0, 0, 360 * min(self.current / self.total, 1), 15
            )
            self.top.update()
        except tk.TclError:
            self._closed = True

    def post(self, callback):
        token = self._generation

        def guarded():
            if self._closed or token != self._generation:
                return
            try:
                if self.top.winfo_exists():
                    callback()
            except tk.TclError:
                pass

        try:
            self.parent.after(0, guarded)
        except tk.TclError:
            pass

    def close(self):
        if getattr(self.parent, "_dns_active_progress", None) is self:
            delattr(self.parent, "_dns_active_progress")
        if self._closed:
            return
        self._closed = True
        self._generation += 1
        try:
            self.top.grab_release()
        except tk.TclError:
            pass
        try:
            self.top.destroy()
        except tk.TclError:
            pass


def _show_cancel_result(parent, result):
    failures = result.details.get("cleanup_failures", [])
    if failures:
        paths = "\n".join(path for path, _error in failures)
        messagebox.showerror(
            "취소 정리 실패",
            "작업은 취소되었으나 일부 결과 파일을 삭제하지 못했습니다.\n\n" + paths,
            parent=parent,
        )
    else:
        messagebox.showinfo(
            "취소됨",
            "전체 작업이 취소되었습니다.\n이번 작업의 결과 파일은 생성되지 않았습니다.",
            parent=parent,
        )


def _run_service_conversion(parent, paths, title, output_folder, service):
    if not paths:
        return
    output_dir = os.path.join(os.path.dirname(paths[0]), output_folder)
    cancellation = ProcessingCancellation()
    progress = ProgressWindow(parent, title, len(paths), cancellation)

    def report(event):
        progress.post(
            lambda: progress.update_progress(
                event.completed,
                os.path.basename(event.current_file or ""),
                event.message,
            )
        )

    def worker():
        try:
            result = service(
                paths,
                output_dir,
                report,
                cancellation=cancellation,
            )

            def finish():
                progress.close()
                if result.state == JobState.CANCELLED:
                    _show_cancel_result(parent, result)
                elif result.output_files:
                    messagebox.showinfo(
                        "변환 완료",
                        f"{len(result.output_files)}개 파일을 변환했습니다.\n\n"
                        f"저장 위치: {output_dir}",
                        parent=parent,
                    )
                    try:
                        os.startfile(output_dir)
                    except OSError:
                        pass
                else:
                    messagebox.showwarning("변환 결과", result.message, parent=parent)

            progress.post(finish)
        except Exception as error:
            progress.post(
                lambda: (
                    progress.close(),
                    messagebox.showerror("변환 오류", str(error), parent=parent),
                )
            )

    threading.Thread(
        target=worker, daemon=True, name="conversion-service-ui"
    ).start()


def convert_docx_to_pdf(parent_root):
    paths = filedialog.askopenfilenames(
        parent=parent_root,
        title="[DOCX -> PDF] 변환할 Word 파일을 선택하세요",
        filetypes=[("Word Files", "*.docx *.doc")],
    )
    _run_service_conversion(
        parent_root,
        paths,
        "DOC/DOCX → PDF",
        "변환완료_PDF",
        conversion_service.convert_docx_to_pdf,
    )


def convert_xls_to_xlsx(parent_root):
    paths = filedialog.askopenfilenames(
        parent=parent_root,
        title="[XLS -> XLSX] 변환할 구버전 Excel 파일을 선택하세요",
        filetypes=[("XLS Files", "*.xls")],
    )
    _run_service_conversion(
        parent_root,
        paths,
        "XLS → XLSX",
        "변환완료_XLSX",
        conversion_service.convert_xls_to_xlsx,
    )


def convert_hwp_to_pdf(parent_root):
    paths = filedialog.askopenfilenames(
        parent=parent_root,
        title="[HWP -> PDF] 변환할 한글 파일을 선택하세요",
        filetypes=[("HWP Files", "*.hwp *.hwpx")],
    )
    if not paths:
        return
    output_dir = os.path.join(os.path.dirname(paths[0]), "변환완료_PDF")
    os.makedirs(output_dir, exist_ok=True)
    cancellation = ProcessingCancellation()
    progress = ProgressWindow(parent_root, "HWP/HWPX → PDF", len(paths), cancellation)
    messagebox.showinfo(
        "안내",
        "한글 연동 중 보안 승인 팝업이 표시되면 [모두 허용]을 선택해 주세요.",
        parent=progress.top,
    )

    def worker():
        pythoncom.CoInitialize()
        hwp = None
        failed = []
        try:
            try:
                hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
            except Exception:
                hwp = win32.Dispatch("HWPFrame.HwpObject")
            try:
                hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
                hwnd = hwp.XHwpWindows.Item(0).WindowHandle
                if hwnd:
                    ctypes.windll.user32.ShowWindow(hwnd, 5)
                    ctypes.windll.user32.SetForegroundWindow(hwnd)
            except Exception:
                logging.warning("HWP 보안 모듈 등록 또는 창 제어 실패", exc_info=True)

            for index, source in enumerate(paths, start=1):
                if cancellation.should_cancel():
                    break
                filename = os.path.basename(source)
                target = Path(
                    conversion_service._unique_output_path(
                        Path(output_dir), Path(source), ".pdf"
                    )
                )
                cancellation.reserve_output(target)
                try:
                    hwp.Open(os.path.abspath(source), "", "force")
                    hwp.HAction.GetDefault(
                        "FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet
                    )
                    hwp.HParameterSet.HFileOpenSave.filename = str(target.resolve())
                    hwp.HParameterSet.HFileOpenSave.Format = "PDF"
                    executed = hwp.HAction.Execute(
                        "FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet
                    )
                    if executed is False or not target.is_file() or target.stat().st_size == 0:
                        raise OSError("HWP PDF 결과 파일이 생성되지 않았습니다.")
                except Exception:
                    target.unlink(missing_ok=True)
                    failed.append(filename)
                progress.post(
                    lambda i=index, name=filename: progress.update_progress(i, name)
                )
                if cancellation.should_cancel():
                    break
        except Exception as error:
            logging.exception("HWP 변환 실패")
            progress.post(
                lambda: messagebox.showerror(
                    "변환 오류", str(error), parent=parent_root
                )
            )
        finally:
            if hwp is not None:
                try:
                    hwp.Quit()
                except Exception:
                    logging.warning("HWP 종료 실패", exc_info=True)
            pythoncom.CoUninitialize()
            if cancellation.should_cancel():
                failures = cancellation.rollback_outputs()
                result = type(
                    "CancelResult",
                    (),
                    {"details": {"cleanup_failures": failures}},
                )()
                progress.post(
                    lambda: (progress.close(), _show_cancel_result(parent_root, result))
                )
            else:
                def finish():
                    progress.close()
                    if failed:
                        messagebox.showwarning(
                            PROGRAM_NAME,
                            f"{len(failed)}개 파일을 변환하지 못했습니다.",
                            parent=parent_root,
                        )
                    else:
                        messagebox.showinfo(
                            "변환 완료",
                            f"{len(paths)}개 파일을 PDF로 변환했습니다.\n\n{output_dir}",
                            parent=parent_root,
                        )
                        try:
                            os.startfile(output_dir)
                        except OSError:
                            pass
                progress.post(finish)

    threading.Thread(target=worker, daemon=True, name="hwp-pdf-converter").start()
