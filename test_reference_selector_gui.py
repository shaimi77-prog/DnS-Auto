"""기준 파일·페이지 탐색 창의 실제 Tk 동작 스모크 테스트."""

import json
from pathlib import Path
import tempfile
import tkinter as tk

import fitz

from engine_Drag import ReferencePageSelector, select_reference_page


def descendants(widget):
    for child in widget.winfo_children():
        yield child
        yield from descendants(child)


def make_pdf(path, pages, prefix):
    document = fitz.open()
    for index in range(pages):
        page = document.new_page()
        page.insert_text((72, 72), f"{prefix} PAGE {index + 1} " + "TEXT " * 20)
    document.save(path)
    document.close()


def main():
    checks = {}
    with tempfile.TemporaryDirectory() as temporary:
        first = Path(temporary) / "first.pdf"
        second = Path(temporary) / "second.pdf"
        make_pdf(first, 5, "FIRST")
        make_pdf(second, 2, "SECOND")
        suggestion = select_reference_page([str(first), str(second)])
        root = tk.Tk()
        root.withdraw()

        def exercise():
            top = next(child for child in root.winfo_children() if isinstance(child, tk.Toplevel))
            widgets = list(descendants(top))
            buttons = {
                widget.cget("text"): widget
                for widget in widgets
                if isinstance(widget, tk.Button)
            }
            labels = [widget for widget in widgets if isinstance(widget, tk.Label)]
            buttons["다음 페이지"].invoke()
            buttons["다음 페이지"].invoke()
            buttons["다음 페이지"].invoke()
            checks["manual_beyond_three_pages"] = any(
                "4페이지" in label.cget("text") for label in labels
            )
            buttons["다음 파일"].invoke()
            checks["next_file_opens_first_page"] = any(
                "second.pdf / 1페이지" in label.cget("text") for label in labels
            )
            checks["previous_page_disabled_at_first"] = (
                str(buttons["이전 페이지"].cget("state")) == str(tk.DISABLED)
            )
            buttons["이 페이지를 기준으로 선택"].invoke()

        root.after(250, exercise)
        selector = ReferencePageSelector(
            [str(first), str(second)], suggestion, parent_root=root
        )
        checks["selected_second_first_page"] = (
            selector.result is not None
            and Path(selector.result["pdf_path"]).name == "second.pdf"
            and selector.result["page_index"] == 0
        )
        root.destroy()
    print(json.dumps({"checks": checks, "all_passed": all(checks.values())}, ensure_ascii=False))
    raise SystemExit(0 if all(checks.values()) else 1)


if __name__ == "__main__":
    main()