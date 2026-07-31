import ast
import sys
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from core.models import JobState, ProgressEvent  # noqa: E402


def top_level_functions(path):
    tree = ast.parse(path.read_text(encoding="utf-8-sig"))
    return [node.name for node in tree.body if isinstance(node, ast.FunctionDef)]


class ActivePathAndGuardTests(unittest.TestCase):
    def test_converter_entry_points_are_unique(self):
        names = top_level_functions(ROOT / "utils_converter.py")
        for name in (
            "convert_docx_to_pdf",
            "convert_hwp_to_pdf",
            "convert_xls_to_xlsx",
        ):
            self.assertEqual(names.count(name), 1)

    def test_excel_gui_entry_point_is_unique(self):
        names = top_level_functions(ROOT / "engine_Sheet.py")
        self.assertEqual(names.count("run_application"), 1)

    def test_pdf_save_is_guarded_by_progress_save_phase(self):
        source = (ROOT / "engine_Drag.py").read_text(encoding="utf-8")
        save_index = source.index("wb_write.save(output_path)")
        guard_index = source.rindex("progress.enter_save_phase()", 0, save_index)
        self.assertLess(guard_index, save_index)

    def test_excel_service_exposes_cancellation_contract(self):
        tree = ast.parse(
            (ROOT / "services" / "sheet_service.py").read_text(encoding="utf-8")
        )
        merge = next(
            node
            for node in tree.body
            if isinstance(node, ast.FunctionDef) and node.name == "merge_workbooks"
        )
        self.assertIn("cancellation", [arg.arg for arg in merge.args.args])

    def test_progress_event_keeps_estimator_metadata_and_sheet(self):
        event = ProgressEvent(
            1,
            2,
            "처리 중",
            "A.xlsx",
            current_sheet="4월",
            activity="excel",
            elapsed_seconds=3,
            estimated_remaining_seconds=4,
            estimate_status="available",
        )
        self.assertEqual(event.current_sheet, "4월")
        self.assertEqual(event.estimated_remaining_seconds, 4)

    def test_cancelled_is_distinct_job_state(self):
        self.assertEqual(JobState.CANCELLED.value, "cancelled")


if __name__ == "__main__":
    unittest.main()
