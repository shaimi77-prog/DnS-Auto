import tkinter as tk
import unittest
from unittest.mock import Mock, patch

import engine_Drag
from DnS_Auto_Main import DnSAIMainApp


class PdfModeGuiTests(unittest.TestCase):
    def setUp(self):
        self.root = tk.Tk()
        self.root.withdraw()
        self.app = DnSAIMainApp(self.root)

    def tearDown(self):
        if self.root.winfo_exists():
            self.root.destroy()

    def test_dashboard_exposes_shared_hover_description_without_mode_state(self):
        self.assertTrue(hasattr(self.app, "dashboard_description"))
        self.assertFalse(hasattr(self.app, "pdf_collection_mode_var"))
        self.assertFalse(hasattr(self.app, "force_ocr_var"))

    def test_auxiliary_mode_buttons_stay_inside_pdf_container(self):
        for button in (self.app.fast_mode_button, self.app.careful_mode_button):
            placement = button.place_info()
            self.assertGreaterEqual(float(placement["relx"]), 0.0)
            self.assertGreaterEqual(float(placement["rely"]), 0.0)
            self.assertLessEqual(
                float(placement["relx"]) + float(placement["relwidth"]), 1.0
            )
            self.assertLessEqual(
                float(placement["rely"]) + float(placement["relheight"]), 1.0
            )

    def test_standard_button_covers_entire_pdf_container(self):
        placement = self.app.standard_mode_button.place_info()
        self.assertEqual(float(placement["relx"]), 0.0)
        self.assertEqual(float(placement["rely"]), 0.0)
        self.assertEqual(float(placement["relwidth"]), 1.0)
        self.assertEqual(float(placement["relheight"]), 1.0)

    def test_large_buttons_are_flat_until_pressed(self):
        self.assertEqual(self.app.standard_mode_button.cget("relief"), "flat")
        self.assertEqual(self.app.excel_merge_button.cget("relief"), "flat")

    def test_all_mode_titles_and_subtitles_share_typography(self):
        title_fonts = {
            self.app.fast_title_label.cget("font"),
            self.app.standard_title_label.cget("font"),
            self.app.careful_title_label.cget("font"),
        }
        subtitle_fonts = {
            self.app.fast_subtitle_label.cget("font"),
            self.app.standard_subtitle_label.cget("font"),
            self.app.careful_subtitle_label.cget("font"),
        }
        self.assertEqual(len(title_fonts), 1)
        self.assertEqual(len(subtitle_fonts), 1)
        self.assertNotEqual(title_fonts, subtitle_fonts)

    def test_pressing_standard_text_does_not_execute_until_release(self):
        with patch.object(self.app, "run_drag_engine") as run:
            self.app.pdf_title_label.event_generate("<ButtonPress-1>")
        run.assert_not_called()
        self.app.standard_mode_button.configure(relief="flat")

    def test_gui_passes_selected_mode_and_never_forces_ocr(self):
        bridge = Mock()
        bridge.run_pdf_application.return_value = False
        with patch("DnS_Auto_Main.runtime_module", return_value=bridge):
            self.app.run_drag_engine(engine_Drag.PDF_MODE_CAREFUL)
        bridge.run_pdf_application.assert_called_once_with(
            self.root,
            force_ocr=False,
            pdf_collection_mode=engine_Drag.PDF_MODE_CAREFUL,
        )

    def test_fast_confirmation_cancel_stops_before_file_selection(self):
        bridge = Mock()
        with patch("DnS_Auto_Main.messagebox.askyesno", return_value=False), patch(
            "DnS_Auto_Main.runtime_module", return_value=bridge
        ):
            self.app.run_drag_engine(engine_Drag.PDF_MODE_FAST)
        bridge.run_pdf_application.assert_not_called()

    def test_standard_is_direct_default_without_popup(self):
        bridge = Mock()
        bridge.run_pdf_application.return_value = False
        with patch("DnS_Auto_Main.runtime_module", return_value=bridge):
            self.app.run_drag_engine()
        bridge.run_pdf_application.assert_called_once_with(
            self.root,
            force_ocr=False,
            pdf_collection_mode=engine_Drag.PDF_MODE_STANDARD,
        )


if __name__ == "__main__":
    unittest.main()
