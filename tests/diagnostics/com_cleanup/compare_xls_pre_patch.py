"""COM 종료 패치 직전 백업과 현재 XLS 결과를 같은 COM 환경에서 비교한다."""

import importlib.util
import json
import sys
from datetime import datetime
from pathlib import Path


ROOT = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(ROOT))
from tests.diagnostics.eta.measure_conversion_eta import compare_xlsx  # noqa: E402


path = ROOT / "백업" / "20260805-002" / "services" / "conversion_service.py"
spec = importlib.util.spec_from_file_location("pre_com_patch_conversion", path)
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)
stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
root = ROOT / "tests" / "results" / "com_cleanup" / f"xls-prepatch-{stamp}"
source_dir = ROOT / "tests" / "test_files" / "엑셀 취합 테스트" / "6. xls변환하기"
sources = sorted(source_dir.glob("*.xls"))
result = module.convert_xls_to_xlsx([str(item) for item in sources], str(root / "outputs"))
current = ROOT / "tests" / "results" / "com_cleanup" / "actual-20260805-092903" / "xls"
equal, differences = compare_xlsx(root / "outputs", current)
payload = {
    "state": result.state.value,
    "input_count": len(sources),
    "output_count": len(result.output_files),
    "failed_count": len(result.failed_files),
    "cell_content_equal_to_current": equal,
    "different_workbook_count": len(differences),
}
(root / "summary.json").write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
print(json.dumps({"result_root": str(root), **payload}, ensure_ascii=False))
