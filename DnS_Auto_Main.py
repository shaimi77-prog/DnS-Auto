"""DnS Auto의 메인 화면과 기능별 실행 진입점."""

import logging
import tkinter as tk
from tkinter import font as tkfont

import engine_Drag
import engine_Sheet
import ui_mcp_bridge
import utils_converter
from config import PROGRAM_NAME, setup_detailed_logging


class DnSAIMainApp:
    """취합 및 파일 변환 기능을 제공하는 메인 화면."""

    def __init__(self, root):
        self.root = root
        self.root.protocol("WM_DELETE_WINDOW", self._on_root_close)
        self.root.title(f"{PROGRAM_NAME} - 메인 대시보드")

        # 화면 중앙 배치 및 크기 설정
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        w, h = 600, 480
        x = (sw - w) // 2
        y = (sh - h) // 2
        self.root.geometry(f"{w}x{h}+{x}+{y}")
        self.root.configure(bg="#F5F5F5")

        self.setup_ui()

    def _on_root_close(self):
        progress = getattr(self.root, "_dns_active_progress", None)
        if progress is not None:
            progress._on_close_request()
            return
        self.root.destroy()

    def setup_ui(self):
        title_font = tkfont.Font(family="맑은 고딕", size=24, weight="bold")
        desc_font = tkfont.Font(family="맑은 고딕", size=10)
        btn_font = tkfont.Font(family="맑은 고딕", size=12, weight="bold")

        # 상단 타이틀 영역
        top_frame = tk.Frame(self.root, bg="#F5F5F5", pady=30)
        top_frame.pack(fill=tk.X)

        tk.Label(top_frame, text="DnS Auto", font=title_font, fg="#1565C0", bg="#F5F5F5").pack()
        tk.Label(top_frame, text="원하시는 취합 작업의 종류를 선택해 주세요.", font=desc_font, fg="#555555", bg="#F5F5F5").pack(pady=(10, 0))

        # 중앙 버튼 영역
        btn_frame = tk.Frame(self.root, bg="#F5F5F5")
        btn_frame.pack(expand=True, fill=tk.BOTH, padx=40, pady=20)

        # PDF 영역 지정 취합
        pdf_btn = tk.Button(btn_frame, text="PDF 파일 취합\n(Drag)", justify="center", font=btn_font, bg="#E3F2FD", fg="#0D47A1",
                            activebackground="#BBDEFB", relief="flat", bd=1, cursor="hand2",
                            command=self.run_drag_engine)
        pdf_btn.pack(side=tk.LEFT, expand=True, fill=tk.BOTH, padx=10, pady=10)

        self.force_ocr_var = tk.BooleanVar(value=False)
        tk.Checkbutton(
            self.root,
            text="PDF 취합 시 강제 OCR 사용",
            variable=self.force_ocr_var,
            bg="#F5F5F5",
            activebackground="#F5F5F5",
            fg="#455A64",
            font=("맑은 고딕", 9)
        ).pack(pady=(0, 4))

        # Excel 시트 취합
        excel_btn = tk.Button(btn_frame, text="엑셀 파일 취합\n(Sheet)", justify="center", font=btn_font, bg="#E8F5E9", fg="#1B5E20",
                              activebackground="#C8E6C9", relief="flat", bd=1, cursor="hand2",
                              command=self.run_sheet_engine)
        excel_btn.pack(side=tk.RIGHT, expand=True, fill=tk.BOTH, padx=10, pady=10)

        # 파일 변환 기능
        util_frame = tk.Frame(self.root, bg="#ECEFF1", pady=10, padx=40)
        util_frame.pack(fill=tk.X, pady=(0, 20))

        tk.Label(util_frame, text="🛠️ 파일 변환 유틸리티", font=("맑은 고딕", 10, "bold"), bg="#ECEFF1", fg="#37474F").pack(anchor="w", pady=(0, 10))

        util_btn_frame = tk.Frame(util_frame, bg="#ECEFF1")
        util_btn_frame.pack(fill=tk.X)

        docx_btn = tk.Button(util_btn_frame, text="DOCX → PDF 일괄 변환", font=("맑은 고딕", 9), bg="#FFFFFF", fg="#1565C0",
                             activebackground="#BBDEFB", relief="ridge", bd=1, cursor="hand2",
                             command=self.run_docx_converter)
        docx_btn.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 3))

        hwp_btn = tk.Button(util_btn_frame, text="HWP, HWPX → PDF 일괄 변환", font=("맑은 고딕", 9), bg="#FFFFFF", fg="#D84315",
                            activebackground="#FFCCBC", relief="ridge", bd=1, cursor="hand2",
                            command=self.run_hwp_converter)
        hwp_btn.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=3)

        xls_btn = tk.Button(util_btn_frame, text="XLS → XLSX 일괄 변환", font=("맑은 고딕", 9), bg="#FFFFFF", fg="#2E7D32",
                            activebackground="#C8E6C9", relief="ridge", bd=1, cursor="hand2",
                            command=self.run_xls_converter)
        xls_btn.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(3, 0))

        # 저작권 표시
        bottom_frame = tk.Frame(self.root, bg="#E0E0E0", height=30)
        bottom_frame.pack(side=tk.BOTTOM, fill=tk.X)
        tk.Label(bottom_frame, text="© 두부코드(DOOBOO_CODE)", font=("맑은 고딕", 9), bg="#E0E0E0", fg="#757575").pack(pady=5)

    def run_drag_engine(self):
        logging.info("메인 UI -> [PDF 파일 취합 (Drag)] 버튼 클릭됨")
        success = ui_mcp_bridge.run_pdf_application(self.root, force_ocr=self.force_ocr_var.get())
        self.check_exit(success)

    def run_sheet_engine(self):
        logging.info("메인 UI -> [엑셀 파일 취합 (Sheet)] 버튼 클릭됨")
        success = engine_Sheet.run_application(self.root)
        self.check_exit(success)

    def run_hwp_converter(self):
        logging.info("메인 UI -> [HWP -> PDF 변환] 버튼 클릭됨")
        utils_converter.convert_hwp_to_pdf(self.root)

    def run_docx_converter(self):
        logging.info("메인 UI -> [DOCX -> PDF 변환] 버튼 클릭됨")
        utils_converter.convert_docx_to_pdf(self.root)

    def run_xls_converter(self):
        logging.info("메인 UI -> [XLS -> XLSX 변환] 버튼 클릭됨")
        utils_converter.convert_xls_to_xlsx(self.root)

    def check_exit(self, success):
        """취합 작업이 완료된 경우 다음 작업 진행 여부를 확인합니다."""
        if success:
            from tkinter import messagebox

            # 파일 탐색기 뒤에 확인 창이 가려지지 않도록 메인 창을 전면으로 이동합니다.
            self.root.attributes('-topmost', True)
            self.root.attributes('-topmost', False)
            self.root.lift()
            self.root.focus_force()

            if messagebox.askokcancel("작업 완료", "데이터 취합 작업이 완료되었습니다.\n\n프로그램을 종료하시겠습니까?\n(새로운 작업을 계속하려면 '취소'를 클릭하세요)", parent=self.root):
                logging.info("사용자 선택: 프로그램 종료")
                self.root.destroy()
            else:
                logging.info("사용자 선택: 새로운 작업 계속")


def main():
    """로그를 초기화하고 메인 화면을 실행합니다."""
    setup_detailed_logging()
    logging.info(f"=== {PROGRAM_NAME} 메인 UI 시작 ===")

    root = tk.Tk()
    app = DnSAIMainApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
