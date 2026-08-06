"""HWP DispatchEx 반복·별도 워커 격리시험 오케스트레이터."""

from __future__ import annotations

import json
import subprocess
import sys
from datetime import datetime
from pathlib import Path


ROOT = Path(__file__).resolve().parents[3]
WORKER = Path(__file__).with_name("hwp_instance_probe.py")
FIXTURE = ROOT / "tests" / "hwp_dispatchex_isolation" / "fixtures" / "isolation_sample.hwpx"


def command(method, mode, output, scenario):
    return [
        sys.executable,
        "-B",
        str(WORKER),
        "--method",
        method,
        "--mode",
        mode,
        "--fixture",
        str(FIXTURE),
        "--output-dir",
        str(output),
        "--scenario-id",
        scenario,
    ]


def run_once(method, mode, output, scenario):
    completed = subprocess.run(
        command(method, mode, output, scenario),
        text=True,
        encoding="utf-8",
        capture_output=True,
        timeout=180,
        check=False,
    )
    lines = [line for line in completed.stdout.splitlines() if line.strip().startswith("{")]
    return json.loads(lines[-1]) if lines else {
        "scenario_id": scenario,
        "success": False,
        "error_type": "NO_STRUCTURED_RESULT",
        "returncode": completed.returncode,
    }


def run_separate_workers(output, scenario):
    processes = []
    ready = []
    try:
        for suffix in ("a", "b"):
            process = subprocess.Popen(
                command("dispatchex", "hold", output / suffix, f"{scenario}_{suffix}"),
                stdin=subprocess.PIPE,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                encoding="utf-8",
            )
            processes.append(process)
            ready.append(json.loads(process.stdout.readline()))
        processes[0].stdin.write("quit\n")
        processes[0].stdin.flush()
        final_a = json.loads(processes[0].stdout.readline())
        processes[0].wait(timeout=30)
        processes[1].stdin.write("probe\n")
        processes[1].stdin.flush()
        probed_b = json.loads(processes[1].stdout.readline())
        processes[1].stdin.write("quit\n")
        processes[1].stdin.flush()
        final_b = json.loads(processes[1].stdout.readline())
        processes[1].wait(timeout=30)
        pids = [item["instances"][0]["pid"] for item in ready]
        return {
            "scenario_id": scenario,
            "success": final_a.get("success") and final_b.get("success"),
            "pids": pids,
            "pid_separated": len(set(pids)) == 2,
            "second_alive_after_first_quit": probed_b.get("alive_after_peer_quit", False),
            "instances": [item["instances"][0] for item in ready],
        }
    except Exception as error:
        return {"scenario_id": scenario, "success": False, "error_type": type(error).__name__}
    finally:
        for process in processes:
            if process.poll() is None and process.stdin:
                try:
                    process.stdin.write("quit\n")
                    process.stdin.flush()
                    process.wait(timeout=30)
                except Exception:
                    pass


def sanitize(result):
    for instance in result.get("instances", []):
        instance.pop("process_path", None)
    return result


def main():
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    root = ROOT / "tests" / "hwp_dispatchex_isolation" / "results" / f"run-{stamp}"
    root.mkdir(parents=True)
    results = []
    for repetition in range(1, 4):
        for method in ("ensure", "dispatch", "dispatchex"):
            results.append(sanitize(run_once(method, "single", root / f"{method}-{repetition}", f"single_{method}_{repetition}")))
        results.append(sanitize(run_once("dispatchex", "double", root / f"double-{repetition}", f"double_dispatchex_{repetition}")))
        results.append(sanitize(run_separate_workers(root / f"workers-{repetition}", f"separate_workers_{repetition}")))
    payload = {
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "result_count": len(results),
        "success_count": sum(bool(item.get("success")) for item in results),
        "results": results,
    }
    path = root / "isolation_summary.json"
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"result_root": str(root), **payload}, ensure_ascii=False))


if __name__ == "__main__":
    main()
