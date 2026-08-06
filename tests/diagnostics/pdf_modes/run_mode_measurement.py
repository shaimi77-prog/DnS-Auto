"""Run one privacy-safe PDF mode measurement against public demo fixtures."""

from __future__ import annotations

import argparse
import importlib.util
import json
import sys
import time
from pathlib import Path


ROOT = Path(__file__).resolve().parents[3]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))


def load_baseline_service():
    backup = ROOT / "백업" / "20260805-003"
    engine_spec = importlib.util.spec_from_file_location(
        "engine_Drag", backup / "engine_Drag.py"
    )
    engine = importlib.util.module_from_spec(engine_spec)
    sys.modules["engine_Drag"] = engine
    engine_spec.loader.exec_module(engine)
    # The backed-up source remains immutable; only its runtime resource root
    # points at the real model directory for an equivalent baseline run.
    engine.__file__ = str(ROOT / "engine_Drag.py")
    service_spec = importlib.util.spec_from_file_location(
        "baseline_pdf_service", backup / "services" / "pdf_service.py"
    )
    service = importlib.util.module_from_spec(service_spec)
    service_spec.loader.exec_module(service)
    return service


def prepare_inputs(run_root: Path, dataset: str):
    root_fixture = (
        ROOT / "tests" / "test_files" / "PDF 취합 테스트"
        / "여러 페이지_여러 시트"
    )
    fixture = (
        root_fixture / "GitHub_시연용_더미데이터"
        if dataset == "demo" else root_fixture
    )
    template = (
        fixture / "취합양식_시연용.xlsx"
        if dataset == "demo" else fixture / "취합양식(2월~5월).xlsx"
    )
    source_profile = (
        root_fixture / "patched_header_reinforcement_2to5_20260804"
        / "combined_profile.json"
    )
    profile = json.loads(source_profile.read_text(encoding="utf-8"))
    months = ("4월", "5월") if dataset == "demo" else ("2월", "3월", "4월", "5월")
    if dataset == "demo":
        profile["mapping_sets"] = [
            item for item in profile["mapping_sets"]
            if set(item.get("sheets", ())) == {"4월", "5월"}
        ]
    profile_path = run_root / "demo-profile.json"
    profile_path.write_text(
        json.dumps(profile, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    pdfs = {
        month: [str(path) for path in sorted((fixture / month).glob("*.pdf"))]
        for month in months
    }
    return template, profile_path, pdfs


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--variant", choices=("baseline", "patched"), required=True)
    parser.add_argument("--mode", choices=("fast", "standard", "careful"), required=True)
    parser.add_argument("--dataset", choices=("demo", "full"), default="demo")
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    args.output.mkdir(parents=True, exist_ok=False)
    template, profile, pdfs = prepare_inputs(args.output, args.dataset)
    if args.variant == "baseline":
        service = load_baseline_service()
    else:
        from services import pdf_service as service
    started = time.monotonic()
    kwargs = {}
    if args.variant == "patched":
        kwargs["pdf_collection_mode"] = args.mode
    result = service.merge_pdfs(
        str(template), pdfs, str(profile), str(args.output), False, **kwargs
    )
    elapsed = time.monotonic() - started
    details = result.details or {}
    stats = details.get("ocr_statistics", {})
    report = {
        "variant": args.variant,
        "mode": args.mode,
        "dataset": args.dataset,
        "state": result.state.value,
        "elapsed_seconds": round(elapsed, 3),
        "output_files": [Path(path).name for path in result.output_files],
        "failed_file_count": len(result.failed_files),
        "processed_pages": details.get("processed_pages", 0),
        "fast_skipped_page_count": details.get("fast_skipped_page_count", 0),
        "fast_skipped_file_count": details.get("fast_skipped_file_count", 0),
        "fast_skipped_field_count": details.get("fast_skipped_field_count", 0),
        "experimental": bool(details.get("experimental", False)),
        "ocr_statistics": stats,
    }
    (args.output / "measurement.json").write_text(
        json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(json.dumps(report, ensure_ascii=False))
    if not result.output_files:
        raise SystemExit(2)


if __name__ == "__main__":
    main()
