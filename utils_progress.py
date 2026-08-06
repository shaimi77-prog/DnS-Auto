"""Long-running collection progress dialog shared by PDF and Excel engines."""

from __future__ import annotations

import datetime
import time
import tkinter as tk
from tkinter import messagebox, ttk

from processing_cancellation import ProcessingCancellation
from processing_time import ProcessingTimeEstimator


class ProcessingProgressDialog:
    """Keep Tk responsive and coordinate a safe cancel-all request."""

    def __init__(
        self,
        parent,
        title,
        total_units,
        unit_name,
        planned_work=(),
        clock=time.monotonic,
        cancellation=None,
    ):
        self.parent = parent
        self.total_units = max(int(total_units), 1)
        self.unit_name = unit_name
        self.estimator = ProcessingTimeEstimator(planned_work, clock=clock)
        self.completed_units = 0
        self.cancellation = cancellation or ProcessingCancellation()
        self._cancel_prompt_open = False
        self._closed = False

        self.top = tk.Toplevel(parent)
        self.top.title(title)
        self.top.resizable(False, False)
        self.top.transient(parent)
        self.top.protocol("WM_DELETE_WINDOW", self._on_close_request)

        frame = ttk.Frame(self.top, padding=18)
        frame.pack(fill="both", expand=True)
        self.sheet_var = tk.StringVar(value="")
        self.file_var = tk.StringVar(value="처리를 준비하고 있습니다.")
        self.detail_var = tk.StringVar(value="")
        self.time_var = tk.StringVar(value="경과 00:00 | 예상 잔여 계산 중")
        ttk.Label(frame, textvariable=self.sheet_var, width=64, anchor="w").pack(fill="x")
        ttk.Label(
            frame,
            textvariable=self.file_var,
            font=("", 10, "bold"),
            width=64,
            anchor="w",
        ).pack(fill="x")
        ttk.Label(frame, textvariable=self.detail_var, width=64, anchor="w").pack(
            fill="x", pady=(8, 8)
        )
        self.bar = ttk.Progressbar(
            frame, mode="determinate", maximum=self.total_units, length=520
        )
        self.bar.pack(fill="x")
        ttk.Label(frame, textvariable=self.time_var, width=64, anchor="w").pack(
            fill="x", pady=(8, 0)
        )
        self.top.update_idletasks()
        x = parent.winfo_rootx() + max(
            (parent.winfo_width() - self.top.winfo_reqwidth()) // 2, 0
        )
        y = parent.winfo_rooty() + max(
            (parent.winfo_height() - self.top.winfo_reqheight()) // 2, 0
        )
        self.top.geometry(f"+{x}+{y}")
        self.top.grab_set()
        setattr(self.parent, "_dns_active_progress", self)
        self._refresh_events()

    def _on_close_request(self):
        if self._cancel_prompt_open or self.cancellation.save_started:
            return
        self._cancel_prompt_open = True
        try:
            answer = messagebox.askyesno(
                "전체 작업 취소",
                "전체 작업을 취소하시겠습니까?\n\n"
                "현재 처리 중인 파일 또는 페이지 작업이 안전하게 끝난 뒤 취소됩니다.\n"
                "이번 작업에서 생성한 결과 파일은 저장되지 않습니다.",
                parent=self.top,
            )
            if answer and self.cancellation.request_cancel_all():
                self.set_cancelling()
        finally:
            self._cancel_prompt_open = False

    def set_cancelling(self):
        self.sheet_var.set("")
        self.file_var.set("전체 작업 취소를 요청했습니다.")
        self.detail_var.set("현재 작업을 안전하게 마친 뒤 종료합니다…")
        self.time_var.set("")
        self.top.protocol("WM_DELETE_WINDOW", lambda: None)
        self._refresh_events()

    def enter_save_phase(self):
        if not self.cancellation.enter_save_phase():
            return False
        self.sheet_var.set("")
        self.file_var.set("결과 파일을 저장하고 있습니다…")
        self.detail_var.set("")
        self.top.protocol("WM_DELETE_WINDOW", lambda: None)
        self._refresh_events()
        return True

    @staticmethod
    def _format_seconds(seconds):
        return str(datetime.timedelta(seconds=max(int(seconds), 0)))

    def _refresh_time(self):
        metadata = self.estimator.metadata()
        self._set_time_metadata(metadata)

    def _set_time_metadata(self, metadata):
        remaining = metadata["estimated_remaining_seconds"]
        remaining_text = (
            f"약 {self._format_seconds(remaining)}"
            if remaining is not None
            else "계산 중"
        )
        self.time_var.set(
            f"경과 {self._format_seconds(metadata['elapsed_seconds'])} | "
            f"예상 잔여 {remaining_text}"
        )

    def update_from_event(self, event):
        """서비스가 계산한 진행률과 ETA를 수정 없이 표시한다."""
        if self._closed:
            return
        self.completed_units = min(max(int(event.completed), 0), self.total_units)
        self.bar["value"] = self.completed_units
        self.sheet_var.set(
            f"현재 시트: {event.current_sheet}" if event.current_sheet else ""
        )
        self.file_var.set(
            f"현재 파일: {event.current_file}" if event.current_file else event.message
        )
        self.detail_var.set(event.message)
        self._set_time_metadata(
            {
                "elapsed_seconds": event.elapsed_seconds,
                "estimated_remaining_seconds": event.estimated_remaining_seconds,
            }
        )
        self._refresh_events()

    def _refresh_events(self):
        if self._closed:
            return
        try:
            if self.top.winfo_exists():
                self.top.update_idletasks()
                self.top.update()
        except tk.TclError:
            self._closed = True

    def begin_unit(
        self,
        file_name,
        overall_index,
        detail="",
        work_type="unknown",
        ocr_weight=1,
        sheet_name="",
    ):
        if self.cancellation.should_cancel():
            self.set_cancelling()
            return
        current = min(max(int(overall_index), 1), self.total_units)
        self.sheet_var.set(f"현재 시트: {sheet_name}" if sheet_name else "")
        self.file_var.set(f"현재 파일: {file_name}")
        self.detail_var.set(
            f"전체 {self.total_units}{self.unit_name} 중 {current}{self.unit_name} 처리 중"
            + (f" · {detail}" if detail else "")
        )
        self.bar["value"] = max(current - 1, 0)
        self.estimator.begin(work_type, ocr_weight)
        self._refresh_time()
        self._refresh_events()

    def complete_unit(
        self,
        completed_units=None,
        work_type=None,
        ocr_weight=None,
        ocr_initialization_seconds=0,
    ):
        if completed_units is None:
            self.completed_units += 1
        else:
            self.completed_units = int(completed_units)
        self.completed_units = min(max(self.completed_units, 0), self.total_units)
        self.bar["value"] = self.completed_units
        self.estimator.complete(
            work_type=work_type,
            weight=ocr_weight,
            ocr_initialization_seconds=ocr_initialization_seconds,
        )
        if self.cancellation.should_cancel():
            self.set_cancelling()
            return
        self._refresh_time()
        self._refresh_events()

    def observe_ocr_work(self, ocr_weight, duration_seconds):
        self.estimator.observe(
            work_type="ocr",
            weight=ocr_weight,
            duration_seconds=duration_seconds,
        )
        self._refresh_time()
        self._refresh_events()
    def close(self):
        if getattr(self.parent, "_dns_active_progress", None) is self:
            delattr(self.parent, "_dns_active_progress")
        if self._closed:
            return
        self._closed = True
        try:
            self.top.grab_release()
        except tk.TclError:
            pass
        try:
            self.top.destroy()
        except tk.TclError:
            pass
