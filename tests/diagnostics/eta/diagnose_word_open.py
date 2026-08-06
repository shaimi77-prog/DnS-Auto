"""문서 원문/파일명을 출력하지 않고 Word 열기·PDF 내보내기 실패 경계를 진단한다."""

from pathlib import Path
import tempfile

import pythoncom
import win32com.client as win32


ROOT = Path(__file__).resolve().parents[3]
source = next(
    (ROOT / "tests" / "test_files" / "PDF 취합 테스트" / "docx 변환하기").glob("*.docx")
)


def attempt(label, include_password):
    app = document = None
    target = Path(tempfile.gettempdir()) / f"dns-auto-word-diagnostic-{label}.pdf"
    target.unlink(missing_ok=True)
    pythoncom.CoInitialize()
    try:
        app = win32.DispatchEx("Word.Application")
        app.Visible = False
        app.DisplayAlerts = 0
        kwargs = dict(
            FileName=str(source.resolve()),
            ConfirmConversions=False,
            ReadOnly=True,
            AddToRecentFiles=False,
        )
        if include_password:
            kwargs["PasswordDocument"] = "dummy_password_to_force_error_and_skip"
        document = app.Documents.Open(**kwargs)
        document.ExportAsFixedFormat(
            OutputFileName=str(target), ExportFormat=17, OpenAfterExport=False
        )
        print(label, "OK", target.is_file(), target.stat().st_size if target.exists() else 0)
    except Exception as error:
        print(label, "FAIL", type(error).__name__, repr(error))
    finally:
        if document is not None:
            document.Close(SaveChanges=False)
        if app is not None:
            app.Quit()
        pythoncom.CoUninitialize()
        target.unlink(missing_ok=True)


attempt("current_arguments", True)
attempt("without_password", False)
