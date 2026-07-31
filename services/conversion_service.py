"""Word·한글·Excel 변환을 GUI 없이 호출하기 위한 서비스 계층."""

from __future__ import annotations

from pathlib import Path
from typing import Callable, Iterable

from core.models import JobResult, JobState, ProgressEvent
from processing_cancellation import ProcessingCancellation


ProgressReporter = Callable[[ProgressEvent], None]


def _unique_output_path(output_dir: Path, source_path: Path, extension: str) -> Path:
    candidate = output_dir / f"(변환완료){source_path.stem}{extension}"
    counter = 1
    while candidate.exists():
        candidate = output_dir / f"(변환완료){source_path.stem} ({counter}){extension}"
        counter += 1
    return candidate


def _cancelled_result(cancellation: ProcessingCancellation) -> JobResult:
    failures = cancellation.rollback_outputs()
    return JobResult(
        JobState.CANCELLED,
        message=(
            "변환이 취소되었습니다."
            if not failures
            else "변환은 취소되었으나 일부 결과 파일 정리에 실패했습니다."
        ),
        details={"cleanup_failures": failures},
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
    word = None
    pythoncom.CoInitialize()
    try:
        word = win32.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        for index, source in enumerate(paths, start=1):
            if cancellation.should_cancel():
                return _cancelled_result(cancellation)
            if report:
                report(
                    ProgressEvent(
                        index - 1,
                        len(paths),
                        "DOCX를 PDF로 변환 중",
                        str(source),
                        activity="conversion",
                    )
                )
            document = None
            target = _unique_output_path(output_dir, source, ".pdf")
            cancellation.reserve_output(target)
            try:
                document = word.Documents.Open(
                    FileName=str(source.resolve()),
                    ConfirmConversions=False,
                    ReadOnly=True,
                    AddToRecentFiles=False,
                    PasswordDocument="dummy_password_to_force_error_and_skip",
                )
                document.ExportAsFixedFormat(
                    OutputFileName=str(target),
                    ExportFormat=17,
                    OpenAfterExport=False,
                )
                if not target.is_file() or target.stat().st_size == 0:
                    raise OSError("PDF 결과 파일이 생성되지 않았습니다.")
                outputs.append(str(target))
            except Exception:
                target.unlink(missing_ok=True)
                failed.append(str(source))
            finally:
                if document is not None:
                    document.Close(SaveChanges=False)
            if report:
                report(
                    ProgressEvent(
                        index,
                        len(paths),
                        "DOCX 변환 완료",
                        str(source),
                        activity="conversion",
                    )
                )
            if cancellation.should_cancel():
                return _cancelled_result(cancellation)
    finally:
        if word is not None:
            word.Quit()
        pythoncom.CoUninitialize()
    return JobResult(
        JobState.SUCCEEDED if not failed else JobState.FAILED,
        outputs,
        failed,
        "DOCX 변환 완료",
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
    cancellation = cancellation or ProcessingCancellation()
    excel = None
    pythoncom = None
    try:
        try:
            import pythoncom as _pythoncom
            import win32com.client as win32

            pythoncom = _pythoncom
            pythoncom.CoInitialize()
            excel = win32.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
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
                return _cancelled_result(cancellation)
            if report:
                report(
                    ProgressEvent(
                        index - 1,
                        len(sources),
                        "XLS를 XLSX로 변환 중",
                        str(source),
                        activity="conversion",
                    )
                )
            target = _unique_output_path(destination, source, ".xlsx")
            cancellation.reserve_output(target)
            workbook = None
            try:
                if excel is not None:
                    workbook = excel.Workbooks.Open(
                        str(source.resolve()),
                        Password="dummy_password_to_force_error",
                    )
                    workbook.SaveAs(str(target), 51)
                else:
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
                    target_book.save(target)
                if not target.is_file() or target.stat().st_size == 0:
                    raise OSError("XLSX 결과 파일이 생성되지 않았습니다.")
                outputs.append(str(target))
            except Exception:
                target.unlink(missing_ok=True)
                failed.append(str(source))
            finally:
                if workbook is not None:
                    workbook.Close(SaveChanges=False)
            if report:
                report(
                    ProgressEvent(
                        index,
                        len(sources),
                        "XLS 변환 완료",
                        str(source),
                        activity="conversion",
                    )
                )
            if cancellation.should_cancel():
                return _cancelled_result(cancellation)
    finally:
        if excel is not None:
            excel.Quit()
        if pythoncom is not None:
            pythoncom.CoUninitialize()
    return JobResult(
        JobState.SUCCEEDED if not failed else JobState.FAILED,
        outputs,
        failed,
        "XLS 변환 완료",
    )
