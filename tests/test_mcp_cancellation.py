import time
import unittest
import json
import tempfile
from pathlib import Path
from openpyxl import Workbook
from services import pdf_service
from core.job_manager import JobManager
from core.models import JobResult, JobState, ProgressEvent
from processing_cancellation import ProcessingCancellation
class McpCancellationTests(unittest.TestCase):
    def test_manager_cancels_cooperative_job_without_output(self):
        manager = JobManager(); cancellation = ProcessingCancellation()
        def worker(report):
            for index in range(100):
                if cancellation.should_cancel(): return JobResult(JobState.CANCELLED, message="cancelled")
                report(ProgressEvent(index, 100, "working")); time.sleep(0.002)
            return JobResult(JobState.SUCCEEDED, output_files=["unexpected.xlsx"])
        job_id = manager.start(worker, cancellation=cancellation); deadline = time.monotonic() + 1
        while manager.get(job_id).progress is None and time.monotonic() < deadline: time.sleep(0.001)
        self.assertTrue(manager.cancel(job_id))
        while manager.get(job_id).result is None and time.monotonic() < deadline: time.sleep(0.001)
        job = manager.get(job_id); self.assertEqual(job.state, JobState.CANCELLED); self.assertEqual(job.result.output_files, []); self.assertFalse(manager.cancel(job_id))
    def test_pdf_cancel_before_processing_creates_no_output(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            template = root / "template.xlsx"
            book = Workbook(); book.active.title = "Data"; book.save(template)
            profile = root / "profile.json"
            profile.write_text(json.dumps({"mapping_sets": [{"sheets": ["Data"], "header_start": 1, "header_end": 1, "fields": []}]}), encoding="utf-8")
            pdf = root / "input.pdf"; pdf.write_bytes(b"not opened after cancellation")
            cancellation = ProcessingCancellation(); cancellation.request_cancel_all()
            result = pdf_service.merge_pdfs(str(template), {"Data": [str(pdf)]}, str(profile), str(root / "outputs"), cancellation=cancellation)
            self.assertEqual(result.state, JobState.CANCELLED)
            self.assertEqual(list((root / "outputs").glob("*.xlsx")), [])
            self.assertEqual(pdf_service.drag_engine.TEXT_EXTRACTOR._ocr_cache, {})
            self.assertEqual(pdf_service.drag_engine.TEXT_EXTRACTOR.ocr_statistics()["total_ocr_inference_count"], 0)

    def test_uncancellable_job_rejects_request(self):
        manager = JobManager(); job_id = manager.start(lambda report: JobResult(JobState.SUCCEEDED)); deadline = time.monotonic() + 1
        while manager.get(job_id).result is None and time.monotonic() < deadline: time.sleep(0.001)
        self.assertFalse(manager.cancel(job_id))
if __name__ == "__main__": unittest.main()
