import unittest

from processing_time import ProcessingTimeEstimator
from core.models import ProgressEvent


class Clock:
    def __init__(self):
        self.now = 0.0

    def __call__(self):
        return self.now

    def advance(self, seconds):
        self.now += seconds


class ProcessingTimeEstimatorTests(unittest.TestCase):
    def test_calculating_until_two_matching_samples(self):
        clock = Clock()
        estimator = ProcessingTimeEstimator(
            [("native_text", 1)] * 4, clock=clock
        )
        estimator.begin("native_text")
        clock.advance(2)
        estimator.complete(work_type="native_text")
        self.assertEqual(estimator.metadata()["estimate_status"], "calculating")
        estimator.begin("native_text")
        clock.advance(4)
        estimator.complete(work_type="native_text")
        self.assertEqual(estimator.metadata()["estimated_remaining_seconds"], 7)

    def test_ocr_initialization_and_weight_are_removed_from_unit_average(self):
        clock = Clock()
        estimator = ProcessingTimeEstimator(
            [("ocr", 3), ("ocr", 3), ("ocr", 6)], clock=clock
        )
        estimator.begin("ocr", 3)
        clock.advance(15)
        estimator.complete(
            work_type="ocr", weight=3, ocr_initialization_seconds=6
        )
        estimator.begin("ocr", 3)
        clock.advance(9)
        estimator.complete(work_type="ocr", weight=3)
        self.assertEqual(estimator.metadata()["estimated_remaining_seconds"], 18)

    def test_skipped_does_not_pollute_samples_and_completion_is_zero(self):
        clock = Clock()
        estimator = ProcessingTimeEstimator(
            [("skipped", 1), ("native_text", 1), ("native_text", 1)],
            clock=clock,
        )
        for kind, seconds in (("skipped", 50), ("native_text", 2), ("native_text", 2)):
            estimator.begin(kind)
            clock.advance(seconds)
            estimator.complete(work_type=kind)
        self.assertEqual(estimator.metadata()["estimated_remaining_seconds"], 0)

    def test_extra_ocr_observation_updates_eta_without_completing_page(self):
        clock = Clock()
        estimator = ProcessingTimeEstimator(
            [("ocr", 1), ("ocr", 1), ("ocr", 3)], clock=clock
        )
        estimator.begin("ocr", 1)
        clock.advance(2)
        estimator.complete(work_type="ocr", weight=1)
        estimator.observe(work_type="ocr", weight=2, duration_seconds=4)
        self.assertEqual(estimator.completed, 1)
        self.assertEqual(estimator.metadata()["estimated_remaining_seconds"], 8)
    def test_progress_event_exposes_mcp_estimate_contract(self):
        event = ProgressEvent(
            4, 20, "OCR", activity="ocr", elapsed_seconds=35,
            estimated_remaining_seconds=130, estimate_status="available"
        )
        self.assertEqual(event.activity, "ocr")
        self.assertEqual(event.estimated_remaining_seconds, 130)

    def test_arbitrary_types_are_isolated_and_available_after_one_sample(self):
        clock = Clock()
        estimator = ProcessingTimeEstimator(
            [("docx_to_pdf", 1), ("xls_to_xlsx", 1), ("docx_to_pdf", 1)],
            clock=clock,
            minimum_samples=1,
        )
        estimator.begin("docx_to_pdf")
        clock.advance(3)
        estimator.complete()
        self.assertEqual(estimator.metadata()["estimate_status"], "calculating")
        estimator.begin("xls_to_xlsx")
        clock.advance(7)
        estimator.complete()
        self.assertEqual(estimator.metadata()["estimated_remaining_seconds"], 3)

    def test_failed_unit_advances_without_becoming_a_sample(self):
        clock = Clock()
        estimator = ProcessingTimeEstimator(
            [("xlsx_merge", 1), ("xlsx_merge", 1)],
            clock=clock,
            minimum_samples=1,
        )
        estimator.begin("xlsx_merge")
        clock.advance(50)
        estimator.complete(successful=False)
        self.assertEqual(estimator.completed, 1)
        self.assertEqual(estimator.summary()["sample_counts"], {})
        self.assertEqual(estimator.metadata()["estimate_status"], "calculating")


if __name__ == "__main__":
    unittest.main()
