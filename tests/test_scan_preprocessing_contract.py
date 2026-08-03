"""스캔 전처리 GUI·MCP·배포 고지 연결의 정적 회귀시험."""

from pathlib import Path
import unittest


SOURCE = Path(__file__).resolve().parents[1]


class ScanPreprocessingContractTests(unittest.TestCase):
    def test_reference_navigation_labels_and_buttons_exist(self):
        engine = (SOURCE / "engine_Drag.py").read_text(encoding="utf-8")
        for phrase in (
            "이전 파일",
            "이전 페이지",
            "다음 페이지",
            "다음 파일",
            "판독 유형:",
            "OCR 필요",
        ):
            self.assertIn(phrase, engine)

    def test_scanned_reference_confirmation_can_reselect(self):
        engine = (SOURCE / "engine_Drag.py").read_text(encoding="utf-8")
        self.assertIn('ocr_confirmation_state = {}', engine)
        self.assertIn("ReferenceReselectionRequested", engine)
        self.assertIn("다른 페이지 선택", engine)

    def test_preprocessing_is_page_local_and_native_profiles_are_unchanged(self):
        engine = (SOURCE / "engine_Drag.py").read_text(encoding="utf-8")
        self.assertIn("fitz.Rect(clip) == page.rect", engine)
        self.assertIn("orientation_enabled=True", engine)
        self.assertNotIn('"skipped" if int(page.rotation) % 360', engine)
        self.assertIn("inverse_transform_box", engine)
        self.assertNotIn('field["rect_ratio"]', engine)

    def test_mcp_returns_deskew_metadata_from_shared_engine(self):
        service = (SOURCE / "services" / "pdf_service.py").read_text(
            encoding="utf-8"
        )
        self.assertIn('"deskew_pages": []', service)
        self.assertIn("drag_engine._collect_pdf_rows", service)

    def test_opencv_is_direct_and_notice_is_not_duplicated(self):
        requirements = (SOURCE / "requirements.txt").read_text(encoding="utf-8")
        notices = (SOURCE / "THIRD_PARTY_NOTICES.md").read_text(
            encoding="utf-8"
        )
        self.assertIn("opencv-python>=5.0,<6", requirements)
        rows = [
            line for line in notices.splitlines()
            if line.startswith("| OpenCV") or line.startswith("| opencv-python")
        ]
        self.assertEqual(len(rows), 1)


if __name__ == "__main__":
    unittest.main()
