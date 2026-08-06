"""실제 Word/Excel 취소 후 출력 롤백과 COM 잔존 여부를 검증한다."""

from __future__ import annotations

import json
import sys
import time
from datetime import datetime
from pathlib import Path


ROOT = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(ROOT))
from com_process_ownership import capture_processes  # noqa: E402
from processing_cancellation import ProcessingCancellation  # noqa: E402
from services.conversion_service import convert_docx_to_pdf, convert_xls_to_xlsx  # noqa: E402


def count(name):
    return len(capture_processes([name]))


def run(service, sources, output, executable):
    cancellation = ProcessingCancellation()

    def report(event):
        if event.completed == 1:
            cancellation.request_cancel_all()

    before = count(executable)
    result = service([str(path) for path in sources[:2]], str(output), report, cancellation)
    time.sleep(1)
    return {
        "state": result.state.value,
        "output_count_after_rollback": len(list(output.glob("*"))),
        "process_count_before": before,
        "process_count_after": count(executable),
    }


stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
root = ROOT / "tests" / "results" / "com_cleanup" / f"cancel-{stamp}"
docx = sorted((ROOT / "tests" / "test_files" / "PDF 취합 테스트" / "docx 변환하기").glob("*.docx"))
xls = sorted((ROOT / "tests" / "test_files" / "엑셀 취합 테스트" / "6. xls변환하기").glob("*.xls"))
payload = {
    "docx": run(convert_docx_to_pdf, docx, root / "docx", "WINWORD.EXE"),
    "xls": run(convert_xls_to_xlsx, xls, root / "xls", "EXCEL.EXE"),
}
root.mkdir(parents=True, exist_ok=True)
(root / "summary.json").write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
print(json.dumps({"result_root": str(root), **payload}, ensure_ascii=False))
