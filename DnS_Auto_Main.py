"""DnS Auto의 메인 화면과 기능별 실행 진입점."""

import logging
import importlib
import queue
import threading
import tkinter as tk
from tkinter import font as tkfont, messagebox, ttk

from config import PROGRAM_NAME, setup_detailed_logging


PDF_MODE_FAST = "fast"
PDF_MODE_STANDARD = "standard"
PDF_MODE_CAREFUL = "careful"
_RUNTIME_MODULES = {}


def load_runtime_modules(progress_callback=None):
    """대시보드를 막지 않고 무거운 기능 모듈을 한 번만 준비합니다."""
    modules = (
        ("engine_Drag", "PDF 기능을 준비하고 있습니다..."),
        ("engine_Sheet", "Excel 기능을 준비하고 있습니다..."),
        ("ui_mcp_bridge", "취합 연결 기능을 준비하고 있습니다..."),
        ("utils_converter", "문서 변환 기능을 준비하고 있습니다..."),
    )
    for name, message in modules:
        if progress_callback:
            progress_callback(message)
        if name not in _RUNTIME_MODULES:
            logging.info("시작 모듈 준비 시작: %s", name)
            _RUNTIME_MODULES[name] = importlib.import_module(name)
            logging.info("시작 모듈 준비 완료: %s", name)
    return _RUNTIME_MODULES


def runtime_module(name):
    if name not in _RUNTIME_MODULES:
        load_runtime_modules()
    return _RUNTIME_MODULES[name]


class DnSAIMainApp:
    """취합 및 파일 변환 기능을 제공하는 메인 화면."""

    def __init__(self, root):
        self.root = root
        self.root.protocol("WM_DELETE_WINDOW", self._on_root_close)
        self.root.title(f"{PROGRAM_NAME} - 메인 대시보드")

        # 화면 중앙 배치 및 크기 설정
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        w, h = 700, 620
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

        default_description = "원하는 취합 방식을 선택하세요. 일반 PDF 작업에는 표준 모드를 권장합니다."
        self.dashboard_description = tk.StringVar(value=default_description)

        def bind_description(widget, text):
            widget.bind("<Enter>", lambda _event: self.dashboard_description.set(text))
            widget.bind("<Leave>", lambda _event: self.dashboard_description.set(default_description))

        btn_frame.grid_columnconfigure(0, weight=1, uniform="collection")
        btn_frame.grid_columnconfigure(1, weight=1, uniform="collection")
        btn_frame.grid_rowconfigure(0, weight=1)

        pdf_group = tk.Frame(btn_frame, bg="#DDECF7")
        pdf_group.grid(row=0, column=0, sticky="nsew", padx=(10, 10), pady=10)

        standard_description = "PDF에 포함된 문자를 우선 추출하고, 이미지 문자는 OCR로 인식합니다."
        standard_button = tk.Button(
            pdf_group, text="", bg="#DDECF7", activebackground="#D6E8F5",
            relief="flat", overrelief="flat", bd=2, cursor="hand2",
            command=lambda: self.run_drag_engine(PDF_MODE_STANDARD),
        )
        standard_button.place(relx=0, rely=0, relwidth=1, relheight=1)
        bind_description(standard_button, standard_description)

        pdf_title = tk.Label(
            pdf_group, text="PDF 파일 취합\n(Drag)", bg="#DDECF7", fg="#0D47A1",
            font=("맑은 고딕", 12, "bold"), justify="center", cursor="hand2",
        )
        pdf_title.place(relx=0.5, rely=0.48, anchor="center")
        standard_title = tk.Label(
            pdf_group, text="표준 모드", bg="#DDECF7", fg="#0D47A1",
            font=("맑은 고딕", 10, "bold"), justify="center", cursor="hand2",
        )
        standard_title.place(relx=0.5, rely=0.69, anchor="center")
        standard_subtitle = tk.Label(
            pdf_group, text="(기  본  값)", bg="#DDECF7", fg="#0D47A1",
            font=("맑은 고딕", 8, "bold"), justify="center", cursor="hand2",
        )
        standard_subtitle.place(relx=0.5, rely=0.77, anchor="center")
        def pointer_inside(widget):
            x = widget.winfo_pointerx()
            y = widget.winfo_pointery()
            return (
                widget.winfo_rootx() <= x < widget.winfo_rootx() + widget.winfo_width()
                and widget.winfo_rooty() <= y < widget.winfo_rooty() + widget.winfo_height()
            )

        def press_standard_label(_event):
            standard_button.configure(relief="sunken")

        def release_standard_label(_event):
            standard_button.configure(relief="flat")
            if not pointer_inside(pdf_group):
                return
            if pointer_inside(fast_button) or pointer_inside(careful_button):
                return
            self.run_drag_engine(PDF_MODE_STANDARD)

        for label in (pdf_title, standard_title, standard_subtitle):
            label.bind("<ButtonPress-1>", press_standard_label)
            label.bind("<ButtonRelease-1>", release_standard_label)
            bind_description(label, standard_description)

        fast_description = "PDF에 포함된 문자만 취합하고, 이미지 문자는 건너뜁니다."
        fast_button = tk.Button(
            pdf_group, text="",
            bg="#E8F2F9", activebackground="#DFECF5", fg="#0D47A1",
            relief="raised", bd=2, cursor="hand2",
            command=lambda: self.run_drag_engine(PDF_MODE_FAST),
        )
        fast_button.place(relx=0.025, rely=0.025, relwidth=0.46, relheight=0.22)
        bind_description(fast_button, fast_description)
        fast_title = tk.Label(
            pdf_group, text="신속 모드", bg="#E8F2F9", fg="#0D47A1",
            font=("맑은 고딕", 10, "bold"), cursor="hand2",
        )
        fast_title.place(relx=0.255, rely=0.095, anchor="center")
        fast_subtitle = tk.Label(
            pdf_group, text="(OCR 건너뛰기)", bg="#E8F2F9", fg="#0D47A1",
            font=("맑은 고딕", 8, "bold"), cursor="hand2",
        )
        fast_subtitle.place(relx=0.255, rely=0.17, anchor="center")

        careful_description = "PDF에 포함된 문자를 우선 추출하고, 이미지 문자는 표준보다 조금 더 크게 분석합니다."
        careful_button = tk.Button(
            pdf_group, text="",
            bg="#D5E8F5", activebackground="#CBDFEE", fg="#0D47A1",
            relief="raised", bd=2, cursor="hand2",
            command=lambda: self.run_drag_engine(PDF_MODE_CAREFUL),
        )
        careful_button.place(relx=0.515, rely=0.025, relwidth=0.46, relheight=0.22)
        bind_description(careful_button, careful_description)
        careful_title = tk.Label(
            pdf_group, text="신중 모드", bg="#D5E8F5", fg="#0D47A1",
            font=("맑은 고딕", 10, "bold"), cursor="hand2",
        )
        careful_title.place(relx=0.745, rely=0.095, anchor="center")
        careful_subtitle = tk.Label(
            pdf_group, text="(이미지 문자 확대 분석)", bg="#D5E8F5", fg="#0D47A1",
            font=("맑은 고딕", 8, "bold"), cursor="hand2",
        )
        careful_subtitle.place(relx=0.745, rely=0.17, anchor="center")

        def bind_mode_labels(labels, button, mode, description):
            def press(_event):
                button.configure(relief="sunken")

            def release(_event):
                button.configure(relief="raised")
                if pointer_inside(button):
                    self.run_drag_engine(mode)

            for label in labels:
                label.bind("<ButtonPress-1>", press)
                label.bind("<ButtonRelease-1>", release)
                bind_description(label, description)

        bind_mode_labels(
            (fast_title, fast_subtitle), fast_button,
            PDF_MODE_FAST, fast_description,
        )
        bind_mode_labels(
            (careful_title, careful_subtitle), careful_button,
            PDF_MODE_CAREFUL, careful_description,
        )
        self.pdf_mode_container = pdf_group
        self.fast_mode_button = fast_button
        self.careful_mode_button = careful_button
        self.standard_mode_button = standard_button
        self.pdf_title_label = pdf_title
        self.standard_title_label = standard_title
        self.standard_subtitle_label = standard_subtitle
        self.fast_title_label = fast_title
        self.fast_subtitle_label = fast_subtitle
        self.careful_title_label = careful_title
        self.careful_subtitle_label = careful_subtitle

        excel_description = "여러 Excel 파일의 시트 데이터를 설정된 기준에 따라 하나의 통합 문서로 취합합니다."
        excel_button = tk.Button(
            btn_frame, text="엑셀 파일 취합\n(Sheet)", justify="center",
            font=("맑은 고딕", 12, "bold"), bg="#E8F5E9", fg="#1B5E20",
            activebackground="#DDEFE0", relief="flat", overrelief="flat",
            bd=2, cursor="hand2",
            command=self.run_sheet_engine,
        )
        excel_button.grid(row=0, column=1, sticky="nsew", padx=(10, 10), pady=10)
        bind_description(excel_button, excel_description)
        self.excel_merge_button = excel_button

        tk.Label(
            self.root, textvariable=self.dashboard_description,
            bg="#FFFFFF", fg="#000000", font=("맑은 고딕", 9, "bold"),
            height=2, anchor="center", justify="center", wraplength=610,
        ).pack(fill=tk.X, padx=35, pady=(0, 8))

        # 파일 변환 기능
        util_frame = tk.Frame(self.root, bg="#ECEFF1", pady=10, padx=40)
        util_frame.pack(fill=tk.X, pady=(0, 20))

        tk.Label(util_frame, text="🛠️ 파일 변환 유틸리티", font=("맑은 고딕", 10, "bold"), bg="#ECEFF1", fg="#37474F").pack(anchor="w", pady=(0, 10))

        util_btn_frame = tk.Frame(util_frame, bg="#ECEFF1")
        util_btn_frame.pack(fill=tk.X)

        docx_btn = tk.Button(util_btn_frame, text="DOCX → PDF 일괄 변환", font=("맑은 고딕", 9), bg="#FFFFFF", fg="#D84315",
                             activebackground="#FFCCBC", relief="ridge", bd=1, cursor="hand2",
                             command=self.run_docx_converter)
        docx_btn.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 3))

        hwp_btn = tk.Button(util_btn_frame, text="HWP, HWPX → PDF 일괄 변환", font=("맑은 고딕", 9), bg="#FFFFFF", fg="#1565C0",
                            activebackground="#BBDEFB", relief="ridge", bd=1, cursor="hand2",
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

    def run_drag_engine(self, mode=PDF_MODE_STANDARD):
        logging.info("메인 UI -> [PDF 파일 취합 (Drag)] 버튼 클릭됨")
        if mode == PDF_MODE_FAST and not messagebox.askyesno(
            PROGRAM_NAME,
            "신속 모드는 이미지 문자를 인식하지 않습니다. 매핑된 필드 중 "
            "하나라도 PDF 문자로 추출되지 않으면 해당 페이지 행 전체를 "
            "제외합니다. 빈 행은 만들지 않고 다음 정상 결과를 이어서 "
            "적재합니다. 계속하시겠습니까?",
            parent=self.root,
        ):
            return
        success = runtime_module("ui_mcp_bridge").run_pdf_application(
            self.root,
            force_ocr=False,
            pdf_collection_mode=mode,
        )
        self.check_exit(success)

    def run_sheet_engine(self):
        logging.info("메인 UI -> [엑셀 파일 취합 (Sheet)] 버튼 클릭됨")
        success = runtime_module("engine_Sheet").run_application(self.root)
        self.check_exit(success)

    def run_hwp_converter(self):
        logging.info("메인 UI -> [HWP -> PDF 변환] 버튼 클릭됨")
        runtime_module("utils_converter").convert_hwp_to_pdf(self.root)

    def run_docx_converter(self):
        logging.info("메인 UI -> [DOCX -> PDF 변환] 버튼 클릭됨")
        runtime_module("utils_converter").convert_docx_to_pdf(self.root)

    def run_xls_converter(self):
        logging.info("메인 UI -> [XLS -> XLSX 변환] 버튼 클릭됨")
        runtime_module("utils_converter").convert_xls_to_xlsx(self.root)

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


class StartupLoader:
    """가벼운 화면을 먼저 표시하고 기능 모듈은 작업 스레드에서 준비합니다."""

    def __init__(self, root):
        self.root = root
        self.events = queue.Queue()
        self.root.title(f"{PROGRAM_NAME} - 시작 중")
        self.root.configure(bg="#F7F9FC")
        self.root.resizable(False, False)
        width, height = 520, 220
        x = (self.root.winfo_screenwidth() - width) // 2
        y = (self.root.winfo_screenheight() - height) // 2
        self.root.geometry(f"{width}x{height}+{x}+{y}")

        tk.Label(
            root, text="DnS Auto", bg="#F7F9FC", fg="#1565C0",
            font=("맑은 고딕", 25, "bold"),
        ).pack(pady=(38, 8))
        self.status = tk.StringVar(value="PDF·Excel·변환 기능을 준비하고 있습니다.")
        tk.Label(
            root, textvariable=self.status, bg="#F7F9FC", fg="#455A64",
            font=("맑은 고딕", 10),
        ).pack(pady=(0, 18))
        self.progress = ttk.Progressbar(root, mode="indeterminate", length=380)
        self.progress.pack()
        self.progress.start(12)

        threading.Thread(target=self._load, daemon=True).start()
        self.root.after(50, self._poll)

    def _load(self):
        try:
            load_runtime_modules(lambda message: self.events.put(("progress", message)))
            self.events.put(("ready", None))
        except Exception as error:
            logging.exception("시작 모듈 준비 실패")
            self.events.put(("error", str(error)))

    def _poll(self):
        try:
            while True:
                event, value = self.events.get_nowait()
                if event == "progress":
                    self.status.set(value)
                elif event == "ready":
                    self._show_dashboard()
                    return
                else:
                    self.progress.stop()
                    messagebox.showerror(
                        PROGRAM_NAME,
                        f"프로그램 기능을 준비하지 못했습니다.\n\n{value}",
                        parent=self.root,
                    )
                    self.root.destroy()
                    return
        except queue.Empty:
            self.root.after(50, self._poll)

    def _show_dashboard(self):
        self.root.withdraw()
        self.progress.stop()
        for child in self.root.winfo_children():
            child.destroy()
        self.root.resizable(True, True)
        DnSAIMainApp(self.root)
        self.root.update_idletasks()
        self.root.deiconify()
        logging.info("기능 모듈 준비 완료: 메인 대시보드 표시")


def main():
    """로그를 초기화하고 메인 화면을 실행합니다."""
    setup_detailed_logging()
    logging.info(f"=== {PROGRAM_NAME} 메인 UI 시작 ===")

    root = tk.Tk()
    StartupLoader(root)
    root.mainloop()


if __name__ == "__main__":
    main()
