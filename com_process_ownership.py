"""DnS Auto가 생성한 COM 서버 프로세스만 식별하고 정리하는 공용 도구."""

from __future__ import annotations

import os
import gc
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Iterable


@dataclass(frozen=True)
class ProcessIdentity:
    pid: int
    created: str
    executable: str


@dataclass(frozen=True)
class Ownership:
    status: str
    process: ProcessIdentity | None = None


def _process_identity(pid: int) -> ProcessIdentity | None:
    try:
        import win32api
        import win32con
        import win32process

        handle = win32api.OpenProcess(
            win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_VM_READ,
            False,
            int(pid),
        )
        try:
            executable = win32process.GetModuleFileNameEx(handle, 0)
            created = str(win32process.GetProcessTimes(handle)["CreationTime"])
        finally:
            handle.Close()
        return ProcessIdentity(int(pid), created, str(Path(executable).resolve()))
    except Exception:
        return None


def capture_processes(executable_names: Iterable[str]) -> dict[int, ProcessIdentity]:
    """지정 실행 파일명의 접근 가능한 프로세스 스냅숏을 수집한다."""
    import win32process

    names = {name.casefold() for name in executable_names}
    result = {}
    for pid in win32process.EnumProcesses():
        identity = _process_identity(pid)
        if identity and Path(identity.executable).name.casefold() in names:
            result[identity.pid] = identity
    return result


def window_process_id(hwnd: int) -> int | None:
    try:
        import win32process

        _thread, pid = win32process.GetWindowThreadProcessId(int(hwnd))
        return int(pid) or None
    except Exception:
        return None


def confirm_ownership(
    before: dict[int, ProcessIdentity],
    hwnd: int | None,
    executable_names: Iterable[str],
) -> Ownership:
    """신규 PID·시작시간·경로·창 핸들이 일치할 때만 소유권을 확인한다."""
    if not hwnd:
        return Ownership("unconfirmed")
    pid = window_process_id(hwnd)
    identity = _process_identity(pid) if pid else None
    names = {name.casefold() for name in executable_names}
    if (
        identity is None
        or identity.pid in before
        or Path(identity.executable).name.casefold() not in names
    ):
        return Ownership("unconfirmed")
    return Ownership("confirmed", identity)


def process_is_same(identity: ProcessIdentity) -> bool:
    current = _process_identity(identity.pid)
    return current == identity


def wait_for_exit(identity: ProcessIdentity, timeout_seconds: float) -> bool:
    deadline = time.monotonic() + max(float(timeout_seconds), 0.0)
    while process_is_same(identity):
        if time.monotonic() >= deadline:
            return False
        time.sleep(0.1)
    return True


def terminate_confirmed_process(ownership: Ownership) -> bool:
    """완전히 확인된 동일 PID·시작시간·경로에만 제한 종료를 적용한다."""
    if ownership.status != "confirmed" or ownership.process is None:
        return False
    if not process_is_same(ownership.process):
        return True
    try:
        import win32api
        import win32con

        handle = win32api.OpenProcess(
            win32con.PROCESS_TERMINATE | win32con.SYNCHRONIZE,
            False,
            ownership.process.pid,
        )
        try:
            win32api.TerminateProcess(handle, 1)
        finally:
            handle.Close()
        return wait_for_exit(ownership.process, 5.0)
    except Exception:
        return False


def detached_quit_callback(application_object):
    """Quit 호출 직후 마지막 COM 참조를 놓을 수 있는 일회용 콜백."""
    holder = [application_object]

    def quit_and_release():
        application = holder.pop()
        try:
            application.Quit()
        finally:
            del application
            gc.collect()

    return quit_and_release


def cleanup_com_session(
    *,
    application: str,
    close_callbacks: Iterable[Callable[[], None]],
    quit_callback: Callable[[], None] | None,
    ownership: Ownership,
    co_uninitialize: Callable[[], None],
    allow_forced_cleanup: bool = False,
    normal_exit_timeout: float = 5.0,
    additional_exit_timeout: float = 30.0,
) -> dict:
    """각 정리 단계를 독립 실행하고 개인정보 없는 결과만 반환한다."""
    started = time.monotonic()
    close_failed = False
    close_attempted = False
    quit_status = "not_created" if quit_callback is None else "completed"
    error_stage = None
    for callback in close_callbacks:
        close_attempted = True
        try:
            callback()
        except Exception:
            close_failed = True
            error_stage = error_stage or "close"
    if quit_callback is not None:
        try:
            quit_callback()
        except Exception:
            quit_status = "failed"
            error_stage = error_stage or "quit"
        finally:
            quit_callback = None
            gc.collect()
    try:
        co_uninitialize()
    except Exception:
        error_stage = error_stage or "co_uninitialize"

    process_exit_status = "unknown"
    forced_status = "not_needed"
    if ownership.status == "confirmed" and ownership.process is not None:
        exited = wait_for_exit(ownership.process, normal_exit_timeout)
        if not exited:
            exited = wait_for_exit(ownership.process, additional_exit_timeout)
        if exited:
            process_exit_status = "exited"
        elif allow_forced_cleanup:
            forced_status = (
                "completed" if terminate_confirmed_process(ownership) else "failed"
            )
            process_exit_status = "exited" if forced_status == "completed" else "remained"
        else:
            process_exit_status = "remained"
            forced_status = "not_allowed"

    return {
        "application": application,
        "ownership_status": ownership.status,
        "close_status": (
            "failed" if close_failed else "completed" if close_attempted else "not_created"
        ),
        "quit_status": quit_status,
        "process_exit_status": process_exit_status,
        "forced_cleanup_status": forced_status,
        "cleanup_error_stage": error_stage,
        "cleanup_elapsed_seconds": round(time.monotonic() - started, 3),
    }
