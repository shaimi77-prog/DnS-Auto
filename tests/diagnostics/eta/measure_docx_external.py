"""샌드박스 밖 Word COM에서 ETA 패치 전후 DOCX 변환을 검증한다."""

from __future__ import annotations

import importlib.util
import json
import sys
import time
from datetime import datetime
from pathlib import Path

import fitz


ROOT = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(ROOT))
from services import conversion_service as patched  # noqa: E402


def baseline_module():
    path = ROOT / "백업" / "20260805-001" / "services" / "conversion_service.py"
    spec = importlib.util.spec_from_file_location("eta_baseline_docx_service", path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def pdf_pages(path):
    with fitz.open(path) as document:
        return document.page_count


def run(module, sources, destination):
    events = []
    started = time.perf_counter()
    result = module.convert_docx_to_pdf(
        [str(path) for path in sources], str(destination), events.append
    )
    elapsed = time.perf_counter() - started
    outputs = [Path(path) for path in result.output_files]
    available = [event for event in events if event.estimated_remaining_seconds is not None]
    return result, {
        "elapsed_seconds": round(elapsed, 3),
        "state": result.state.value,
        "input_count": len(sources),
        "output_count": len(outputs),
        "failed_count": len(result.failed_files),
        "nonempty_output_count": sum(path.stat().st_size > 0 for path in outputs),
        "page_counts": sorted(pdf_pages(path) for path in outputs),
        "total_pages": sum(pdf_pages(path) for path in outputs),
        "eta_available_event_count": len(available),
        "first_eta_completed": available[0].completed if available else None,
        "timing": result.details.get("timing"),
    }


def main():
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    root = ROOT / "tests" / "results" / "eta" / f"docx-external-{stamp}"
    sources = sorted(
        (ROOT / "tests" / "test_files" / "PDF 취합 테스트" / "docx 변환하기").glob("*.docx")
    )
    report = {"created_at": datetime.now().isoformat(timespec="seconds"), "runs": {}}
    results = {}
    for label, module in (("baseline", baseline_module()), ("patched", patched)):
        destination = root / label
        destination.mkdir(parents=True)
        results[label], report["runs"][label] = run(module, sources, destination)
    report["validation"] = {
        "patched_all_outputs_nonempty": report["runs"]["patched"]["output_count"]
        == report["runs"]["patched"]["nonempty_output_count"],
        "patched_page_count_entries": len(report["runs"]["patched"]["page_counts"]),
    }
    path = root / "measurement.json"
    path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"result_root": str(root), **report}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
