"""Word·한글·Excel 변환을 GUI 없이 호출하기 위한 서비스 계층."""

from __future__ import annotations

from pathlib import Path
from typing import Callable, Iterable
from collections import Counter
import zipfile

from core.models import JobResult, JobState, ProgressEvent
from processing_cancellation import ProcessingCancellation
from processing_time import ProcessingTimeEstimator
from com_process_ownership import (
    Ownership,
    capture_processes,
    cleanup_com_session,
    confirm_ownership,
    detached_quit_callback,
)


ProgressReporter = Callable[[ProgressEvent], None]


def _safe_snapshot(names):
    try:
        return capture_processes(names)
    except Exception:
        return {}


def _safe_ownership(before, hwnd, names):
    try:
        return confirm_ownership(before, hwnd, names)
    except Exception:
        return Ownership("unconfirmed")


def _validate_docx_container(source: Path) -> None:
    """암호화되었거나 손상된 OOXML을 Word의 숨은 암호창 없이 차단한다."""
    if source.suffix.lower() == ".docx" and not zipfile.is_zipfile(source):
        raise ValueError("암호화되었거나 손상된 DOCX 파일입니다.")


def _unique_output_path(output_dir: Path, source_path: Path, extension: str) -> Path:
    candidate = output_dir / f"(변환완료){source_path.stem}{extension}"
    counter = 1
    while candidate.exists():
        candidate = output_dir / f"(변환완료){source_path.stem} ({counter}){extension}"
        counter += 1
    return candidate


def _cancelled_result(cancellation: ProcessingCancellation, estimator=None) -> JobResult:
    failures = cancellation.rollback_outputs()
    return JobResult(
        JobState.CANCELLED,
        message=(
            "변환이 취소되었습니다."
            if not failures
            else "변환은 취소되었으나 일부 결과 파일 정리에 실패했습니다."
        ),
        details={
            "cleanup_failures": failures,
            **({"timing": estimator.summary()} if estimator is not None else {}),
        },
    )


def convert_docx_to_pdf(
    paths: Iterable[str],
    output_dir: str,
    report: ProgressReporter | None = None,
    cancellation: ProcessingCancellation | None = None,
) -> JobResult:
    return _convert_with_word(
        [Path(path) for path in paths],
        Path(output_dir),
        report,
        cancellation or ProcessingCancellation(),
    )


def _convert_with_word(
    paths: list[Path],
    output_dir: Path,
    report: ProgressReporter | None,
    cancellation: ProcessingCancellation,
) -> JobResult:
    import pythoncom
    import win32com.client as win32

    output_dir.mkdir(parents=True, exist_ok=True)
    outputs, failed = [], []
    failure_stages = Counter()
    estimator = ProcessingTimeEstimator(
        [("docx_to_pdf", 1)] * len(paths), minimum_samples=1
    )
    estimator.start()
    word = None
    word_ownership = Ownership("unconfirmed")
    cleanup_details = None
    word_before = _safe_snapshot(["WINWORD.EXE"])
    pythoncom.CoInitialize()
    try:
        word = win32.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        try:
            probe = word.Documents.Add()
            try:
                word_ownership = _safe_ownership(
                    word_before, int(word.ActiveWindow.Hwnd), ["WINWORD.EXE"]
                )
            finally:
                probe.Close(SaveChanges=False)
        except Exception:
            word_ownership = Ownership("unconfirmed")
        for index, source in enumerate(paths, start=1):
            if cancellation.should_cancel():
                return _cancelled_result(cancellation, estimator)
            if report:
                report(
                    ProgressEvent(
                        index - 1,
                        len(paths),
                        "DOCX를 PDF로 변환 중",
                        str(source),
                        activity="conversion",
                        **estimator.metadata(),
                    )
                )
            estimator.begin("docx_to_pdf")
            document = None
            target = _unique_output_path(output_dir, source, ".pdf")
            cancellation.reserve_output(target)
            stage = "document_preflight"
            try:
                _validate_docx_container(source)
                stage = "document_open"
                document = word.Documents.Open(
                    FileName=str(source.resolve()),
                    ConfirmConversions=False,
                    ReadOnly=True,
                    AddToRecentFiles=False,
                )
                stage = "pdf_export"
                document.ExportAsFixedFormat(
                    OutputFileName=str(target),
                    ExportFormat=17,
                    OpenAfterExport=False,
                )
                stage = "output_validation"
                if not target.is_file() or target.stat().st_size == 0:
                    raise OSError("PDF 결과 파일이 생성되지 않았습니다.")
                outputs.append(str(target))
            except Exception:
                target.unlink(missing_ok=True)
                failed.append(str(source))
                failure_stages[stage] += 1
            finally:
                if document is not None:
                    try:
                        document.Close(SaveChanges=False)
                    except Exception:
                        failure_stages["document_close"] += 1
            estimator.complete(
                successful=str(source) not in failed
                and not cancellation.should_cancel()
            )
            sampled = str(source) not in failed and not cancellation.should_cancel()
            if report:
                report(
                    ProgressEvent(
                        index,
                        len(paths),
                        "DOCX 변환 완료",
                        str(source),
                        activity="conversion",
                        **estimator.metadata(),
                    )
                )
            if cancellation.should_cancel():
                if sampled:
                    estimator.discard_last_sample("docx_to_pdf")
                return _cancelled_result(cancellation, estimator)
    finally:
        quit_callback = detached_quit_callback(word) if word is not None else None
        word = None
        cleanup_details = cleanup_com_session(
            application="word",
            close_callbacks=[],
            quit_callback=quit_callback,
            ownership=word_ownership,
            co_uninitialize=pythoncom.CoUninitialize,
            allow_forced_cleanup=True,
        )
    return JobResult(
        JobState.SUCCEEDED if not failed else JobState.FAILED,
        outputs,
        failed,
        "DOCX 변환 완료",
        details={
            "timing": estimator.summary(),
            "failure_stage_counts": dict(sorted(failure_stages.items())),
            "com_cleanup": cleanup_details,
        },
    )


def convert_hwp_to_pdf(
    paths: Iterable[str],
    output_dir: str,
    report: ProgressReporter | None = None,
    cancellation: ProcessingCancellation | None = None,
) -> JobResult:
    return JobResult(
        JobState.NEEDS_USER_ACTION,
        message=(
            "현재 PC의 한글 파일 경로 보안 모듈이 무인 실행을 허용하지 않습니다. "
            "한글 보안 설정을 확인한 뒤 별도 격리 프로세스·시간 제한으로 다시 "
            "검증해야 합니다."
        ),
    )


def convert_xls_to_xlsx(
    paths: Iterable[str],
    output_dir: str,
    report: ProgressReporter | None = None,
    cancellation: ProcessingCancellation | None = None,
) -> JobResult:
    sources = [Path(path) for path in paths]
    destination = Path(output_dir)
    destination.mkdir(parents=True, exist_ok=True)
    outputs, failed = [], []
    failure_stages = Counter()
    cancellation = cancellation or ProcessingCancellation()
    estimator = ProcessingTimeEstimator(
        [("xls_to_xlsx", 1)] * len(sources), minimum_samples=1
    )
    estimator.start()
    excel = None
    pythoncom = None
    excel_ownership = Ownership("unconfirmed")
    cleanup_details = None
    excel_before = _safe_snapshot(["EXCEL.EXE"])
    try:
        try:
            import pythoncom as _pythoncom
            import win32com.client as win32

            pythoncom = _pythoncom
            pythoncom.CoInitialize()
            excel = win32.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            excel_ownership = _safe_ownership(
                excel_before, int(excel.Hwnd), ["EXCEL.EXE"]
            )
        except Exception:
            if pythoncom is not None:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass
                pythoncom = None
            excel = None

        for index, source in enumerate(sources, start=1):
            if cancellation.should_cancel():
                return _cancelled_result(cancellation, estimator)
            if report:
                report(
                    ProgressEvent(
                        index - 1,
                        len(sources),
                        "XLS를 XLSX로 변환 중",
                        str(source),
                        activity="conversion",
                        **estimator.metadata(),
                    )
                )
            estimator.begin("xls_to_xlsx")
            target = _unique_output_path(destination, source, ".xlsx")
            cancellation.reserve_output(target)
            workbook = None
            stage = "workbook_open"
            try:
                if excel is not None:
                    workbook = excel.Workbooks.Open(
                        str(source.resolve()),
                        Password="dummy_password_to_force_error",
                    )
                    stage = "workbook_save"
                    workbook.SaveAs(str(target), 51)
                else:
                    stage = "fallback_open"
                    import xlrd
                    from openpyxl import Workbook

                    source_book = xlrd.open_workbook(str(source))
                    target_book = Workbook()
                    target_book.remove(target_book.active)
                    for sheet_index in range(source_book.nsheets):
                        source_sheet = source_book.sheet_by_index(sheet_index)
                        target_sheet = target_book.create_sheet(source_sheet.name)
                        for row_index in range(source_sheet.nrows):
                            target_sheet.append(source_sheet.row_values(row_index))
                    stage = "fallback_save"
                    target_book.save(target)
                stage = "output_validation"
                if not target.is_file() or target.stat().st_size == 0:
                    raise OSError("XLSX 결과 파일이 생성되지 않았습니다.")
                outputs.append(str(target))
            except Exception:
                target.unlink(missing_ok=True)
                failed.append(str(source))
                failure_stages[stage] += 1
            finally:
                if workbook is not None:
                    try:
                        workbook.Close(SaveChanges=False)
                    except Exception:
                        pass
                    workbook = None
            estimator.complete(
                successful=str(source) not in failed
                and not cancellation.should_cancel()
            )
            sampled = str(source) not in failed and not cancellation.should_cancel()
            if report:
                report(
                    ProgressEvent(
                        index,
                        len(sources),
                        "XLS 변환 완료",
                        str(source),
                        activity="conversion",
                        **estimator.metadata(),
                    )
                )
            if cancellation.should_cancel():
                if sampled:
                    estimator.discard_last_sample("xls_to_xlsx")
                return _cancelled_result(cancellation, estimator)
    finally:
        if pythoncom is not None:
            quit_callback = detached_quit_callback(excel) if excel is not None else None
            excel = None
            cleanup_details = cleanup_com_session(
                application="excel",
                close_callbacks=[],
                quit_callback=quit_callback,
                ownership=excel_ownership,
                co_uninitialize=pythoncom.CoUninitialize,
                allow_forced_cleanup=True,
            )
    return JobResult(
        JobState.SUCCEEDED if not failed else JobState.FAILED,
        outputs,
        failed,
        "XLS 변환 완료",
        details={
            "timing": estimator.summary(),
            "com_cleanup": cleanup_details,
            "failure_stage_counts": dict(sorted(failure_stages.items())),
        },
    )
