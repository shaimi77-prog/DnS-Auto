import sys
import tempfile
import types
import unittest
from pathlib import Path
from unittest.mock import patch


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from processing_cancellation import ProcessingCancellation  # noqa: E402
from services import conversion_service  # noqa: E402


class _Document:
    def ExportAsFixedFormat(self, OutputFileName, **_kwargs):
        Path(OutputFileName).write_bytes(b"pdf")

    def Close(self, **_kwargs):
        pass


class _Word:
    def __init__(self):
        self.Documents = types.SimpleNamespace(Open=lambda **_kwargs: _Document())
        self.Visible = False
        self.DisplayAlerts = 0
        self.quit_called = False

    def Quit(self):
        self.quit_called = True


class ConversionCancellationTests(unittest.TestCase):
    def test_word_cancel_rolls_back_completed_output_and_quits_com(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source1 = root / "one.docx"
            source2 = root / "two.docx"
            source1.write_bytes(b"docx")
            source2.write_bytes(b"docx")
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


if __name__ == "__main__":
    unittest.main()
