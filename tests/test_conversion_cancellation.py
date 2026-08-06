import sys
import tempfile
import types
import unittest
import zipfile
from pathlib import Path
from unittest.mock import patch


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from processing_cancellation import ProcessingCancellation  # noqa: E402
from services import conversion_service  # noqa: E402


def _write_docx_container(path):
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("[Content_Types].xml", "<Types/>")


class _Document:
    def ExportAsFixedFormat(self, OutputFileName, **_kwargs):
        Path(OutputFileName).write_bytes(b"pdf")

    def Close(self, **_kwargs):
        pass


class _Word:
    def __init__(self):
        self.open_kwargs = []
        self.Documents = types.SimpleNamespace(Open=self._open)
        self.Visible = False
        self.DisplayAlerts = 0
        self.quit_called = False

    def _open(self, **kwargs):
        self.open_kwargs.append(kwargs)
        return _Document()

    def Quit(self):
        self.quit_called = True


class ConversionCancellationTests(unittest.TestCase):
    def test_word_cancel_rolls_back_completed_output_and_quits_com(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source1 = root / "one.docx"
            source2 = root / "two.docx"
            _write_docx_container(source1)
            _write_docx_container(source2)
            output = root / "outputs"
            cancellation = ProcessingCancellation()
            word = _Word()
            pythoncom = types.SimpleNamespace(CoInitialize=lambda: None, CoUninitialize=lambda: None)
            client = types.SimpleNamespace(DispatchEx=lambda _name: word)
            win32com = types.ModuleType("win32com")
            win32com.client = client

            def report(event):
                if event.completed == 1:
                    cancellation.request_cancel_all()

            with patch.dict(
                sys.modules,
                {"pythoncom": pythoncom, "win32com": win32com, "win32com.client": client},
            ):
                result = conversion_service.convert_docx_to_pdf(
                    [str(source1), str(source2)],
                    str(output),
                    report,
                    cancellation,
                )
            self.assertEqual(result.state.value, "cancelled")
            self.assertEqual(list(output.glob("*.pdf")), [])
            self.assertTrue(word.quit_called)
            self.assertEqual(result.details["timing"]["sample_counts"], {})
            self.assertNotIn("PasswordDocument", word.open_kwargs[0])

    def test_word_events_share_eta_contract_and_timing_summary(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            sources = [root / "one.docx", root / "two.docx"]
            for source in sources:
                _write_docx_container(source)
            word = _Word()
            pythoncom = types.SimpleNamespace(CoInitialize=lambda: None, CoUninitialize=lambda: None)
            client = types.SimpleNamespace(DispatchEx=lambda _name: word)
            win32com = types.ModuleType("win32com")
            win32com.client = client
            events = []
            with patch.dict(
                sys.modules,
                {"pythoncom": pythoncom, "win32com": win32com, "win32com.client": client},
            ):
                result = conversion_service.convert_docx_to_pdf(
                    [str(source) for source in sources], str(root / "outputs"), events.append
                )
            first_complete = next(event for event in events if event.completed == 1)
            self.assertEqual(first_complete.estimate_status, "available")
            self.assertEqual(result.details["timing"]["sample_counts"], {"docx_to_pdf": 2})

    def test_invalid_docx_is_rejected_before_word_open(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "encrypted-or-damaged.docx"
            source.write_bytes(b"not-an-ooxml-zip")
            word = _Word()
            pythoncom = types.SimpleNamespace(CoInitialize=lambda: None, CoUninitialize=lambda: None)
            client = types.SimpleNamespace(DispatchEx=lambda _name: word)
            win32com = types.ModuleType("win32com")
            win32com.client = client
            with patch.dict(
                sys.modules,
                {"pythoncom": pythoncom, "win32com": win32com, "win32com.client": client},
            ):
                result = conversion_service.convert_docx_to_pdf(
                    [str(source)], str(root / "outputs")
                )
            self.assertEqual(result.state.value, "failed")
            self.assertEqual(word.open_kwargs, [])
            self.assertEqual(
                result.details["failure_stage_counts"], {"document_preflight": 1}
            )


if __name__ == "__main__":
    unittest.main()
