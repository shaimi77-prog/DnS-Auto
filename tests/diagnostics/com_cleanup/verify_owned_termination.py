"""시험이 생성한 전용 Python PID만 대상으로 제한 종료 보완을 검증한다."""

import json
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(ROOT))
from com_process_ownership import Ownership, _process_identity, terminate_confirmed_process


process = subprocess.Popen([sys.executable, "-c", "import time; time.sleep(120)"])
try:
    identity = _process_identity(process.pid)
    result = terminate_confirmed_process(Ownership("confirmed", identity))
    process.wait(timeout=10)
    print(json.dumps({"ownership_confirmed": identity is not None, "terminated": result, "process_exited": process.poll() is not None}))
finally:
    if process.poll() is None:
        process.terminate()
        process.wait(timeout=10)
