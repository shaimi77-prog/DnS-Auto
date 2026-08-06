"""공식 DOCX 자료의 ETA 정확도를 파일명 없이 측정한다."""

from __future__ import annotations

import json
import statistics
import sys
import time
from datetime import datetime
from pathlib import Path

import fitz


ROOT = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(ROOT))
from services.conversion_service import convert_docx_to_pdf  # noqa: E402


def main():
    sources = sorted(
        (ROOT / "tests" / "test_files" / "PDF 취합 테스트" / "docx 변환하기").glob("*.docx")
    )
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    result_root = ROOT / "tests" / "results" / "eta" / f"docx-eta-{stamp}"
    output_dir = result_root / "outputs"
    output_dir.mkdir(parents=True)
    observations = []
    started = time.perf_counter()

    def report(event):
        observations.append(
            {
                "at": time.perf_counter(),
                "completed": event.completed,
                "status": event.estimate_status,
                "estimate": event.estimated_remaining_seconds,
                "elapsed": event.elapsed_seconds,
            }
        )

    result = convert_docx_to_pdf(
        [str(path) for path in sources], str(output_dir), report
    )
    finished = time.perf_counter()
    estimates = []
    seen_completed = set()
    for item in observations:
        if item["estimate"] is None or item["completed"] >= len(sources):
            continue
        if item["completed"] in seen_completed:
            continue
        seen_completed.add(item["completed"])
        actual = max(finished - item["at"], 0.0)
        error = abs(item["estimate"] - actual)
        estimates.append(
            {
                "completed": item["completed"],
                "estimated_remaining_seconds": item["estimate"],
                "actual_remaining_seconds": round(actual, 3),
                "absolute_error_seconds": round(error, 3),
            }
        )
    outputs = [Path(path) for path in result.output_files]
    page_counts = []
    for path in outputs:
        with fitz.open(path) as document:
            page_counts.append(document.page_count)
    errors = [item["absolute_error_seconds"] for item in estimates]
    payload = {
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "elapsed_seconds": round(finished - started, 3),
        "state": result.state.value,
        "input_count": len(sources),
        "output_count": len(outputs),
        "failed_count": len(result.failed_files),
        "nonempty_output_count": sum(path.stat().st_size > 0 for path in outputs),
        "total_pages": sum(page_counts),
        "first_event_status": observations[0]["status"] if observations else None,
        "first_eta_completed": estimates[0]["completed"] if estimates else None,
        "eta_available_event_count": sum(
            item["estimate"] is not None for item in observations
        ),
        "eta_accuracy_point_count": len(estimates),
        "mean_absolute_error_seconds": round(statistics.mean(errors), 3) if errors else None,
        "median_absolute_error_seconds": round(statistics.median(errors), 3) if errors else None,
        "maximum_absolute_error_seconds": round(max(errors), 3) if errors else None,
        "observations": estimates,
        "timing": result.details.get("timing"),
    }
    report_path = result_root / "measurement.json"
    report_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"result_root": str(result_root), **payload}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
