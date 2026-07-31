import json
import subprocess
import time
import zipfile
from pathlib import Path

source_root = Path(__file__).resolve().parents[1]
project_root = source_root.parent
archive = project_root / "release" / "DnS_Auto_Portable.zip"
destination = project_root / f"대화형_이동테스트_{time.time_ns()}"
destination.mkdir()
with zipfile.ZipFile(archive) as package:
    package.extractall(destination)
bundle = destination / "DnS Auto"
request = bundle / "interactive-entry.request.json"
result = bundle / "interactive-entry.result.json"
payload = {"mode": "entry_smoke", "template_path": str(bundle / "inputs" / "dummy.xlsx"), "output_root": str(bundle / "outputs"), "result_path": str(result)}
request.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
completed = subprocess.run([str(bundle / "DnS Auto MCP.exe"), "--interactive-request", str(request)], cwd=Path.home(), capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=30)
assert completed.returncode == 2, (completed.returncode, completed.stderr)
assert result.is_file(), completed.stderr
data = json.loads(result.read_text(encoding="utf-8"))
assert data["state"] == "failed"
assert data["details"]["interactive"] is True
assert "지원하지 않는 대화형 작업" in data["message"]
print(json.dumps({"bundle": str(bundle), "interactive_entry": "loaded", "tk_child_process": "started", "result_contract": data["state"]}, ensure_ascii=False))