"""Word/Excel DispatchEx 이중·별도 워커 반복시험."""

from __future__ import annotations

import json
import subprocess
import sys
from datetime import datetime
from pathlib import Path


ROOT = Path(__file__).resolve().parents[3]
WORKER = Path(__file__).with_name("office_instance_probe.py")


def command(application, mode, scenario):
    return [sys.executable, "-B", str(WORKER), "--application", application, "--mode", mode, "--scenario-id", scenario]


def run_double(application, scenario):
    completed = subprocess.run(command(application, "double", scenario), capture_output=True, text=True, encoding="utf-8", timeout=120)
    lines = [line for line in completed.stdout.splitlines() if line.startswith("{")]
    return json.loads(lines[-1])


def run_workers(application, scenario):
    processes, ready = [], []
    try:
        for suffix in ("a", "b"):
            process = subprocess.Popen(command(application, "hold", f"{scenario}_{suffix}"), stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, encoding="utf-8")
            processes.append(process)
            ready.append(json.loads(process.stdout.readline()))
        processes[0].stdin.write("quit\n"); processes[0].stdin.flush()
        final_a = json.loads(processes[0].stdout.readline()); processes[0].wait(timeout=30)
        processes[1].stdin.write("probe\n"); processes[1].stdin.flush()
        probe_b = json.loads(processes[1].stdout.readline())
        processes[1].stdin.write("quit\n"); processes[1].stdin.flush()
        final_b = json.loads(processes[1].stdout.readline()); processes[1].wait(timeout=30)
        pids = [item["instances"][0]["pid"] for item in ready]
        return {
            "scenario_id": scenario,
            "application": application,
            "success": final_a.get("success") and final_b.get("success"),
            "pids": pids,
            "pid_separated": None not in pids and len(set(pids)) == 2,
            "ownership_confirmed": all(item["instances"][0]["ownership_status"] == "confirmed" for item in ready),
            "second_alive_after_first_quit": probe_b.get("alive_after_peer_quit", False),
        }
    finally:
        for process in processes:
            if process.poll() is None:
                try:
                    process.stdin.write("quit\n"); process.stdin.flush(); process.wait(timeout=30)
                except Exception:
                    pass


def main():
    results = []
    for application in ("word", "excel"):
        for repetition in range(1, 4):
            results.append(run_double(application, f"{application}_double_{repetition}"))
            results.append(run_workers(application, f"{application}_workers_{repetition}"))
    payload = {
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "result_count": len(results),
        "success_count": sum(bool(item.get("success")) for item in results),
        "results": results,
    }
    root = ROOT / "tests" / "results" / "com_cleanup" / datetime.now().strftime("office-isolation-%Y%m%d-%H%M%S")
    root.mkdir(parents=True)
    (root / "summary.json").write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"result_root": str(root), **payload}, ensure_ascii=False))


if __name__ == "__main__":
    main()
