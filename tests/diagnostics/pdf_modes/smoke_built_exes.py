import argparse
import json
import subprocess
import time
from pathlib import Path

import win32gui


def check_gui(executable: Path) -> str:
    existing = set()

    def remember(handle, _extra):
        if "DnS Auto" in win32gui.GetWindowText(handle):
            existing.add(handle)

    win32gui.EnumWindows(remember, None)
    process = subprocess.Popen([str(executable)], cwd=executable.parent)
    try:
        deadline = time.monotonic() + 90
        while time.monotonic() < deadline:
            found = []

            def collect(handle, _extra):
                title = win32gui.GetWindowText(handle)
                if handle not in existing and "DnS Auto" in title:
                    found.append(title)

            win32gui.EnumWindows(collect, None)
            if found:
                return found[0]
            if process.poll() is not None:
                break
            time.sleep(0.1)
        raise AssertionError(f"GUI window not found; exit={process.poll()}")
    finally:
        subprocess.run(
            ["taskkill", "/PID", str(process.pid), "/T", "/F"],
            capture_output=True,
            check=False,
        )
        try:
            process.wait(timeout=10)
        except subprocess.TimeoutExpired:
            process.kill()


def check_mcp(executable: Path) -> dict:
    process = subprocess.Popen(
        [str(executable)],
        cwd=executable.parent,
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        encoding="utf-8",
    )

    def request(identifier, method, params=None):
        payload = {"jsonrpc": "2.0", "id": identifier, "method": method}
        if params is not None:
            payload["params"] = params
        process.stdin.write(json.dumps(payload, ensure_ascii=False) + "\n")
        process.stdin.flush()
        line = process.stdout.readline()
        if not line:
            raise RuntimeError(process.stderr.read())
        return json.loads(line)

    try:
        initialized = request(1, "initialize", {"protocolVersion": "2024-11-05"})
        assert initialized["result"]["serverInfo"]["name"] == "dns-auto-mcp"
        tools = request(2, "tools/list")["result"]["tools"]
        pdf_tools = [tool for tool in tools if "pdf" in tool["name"]]
        schemas = json.dumps(pdf_tools, ensure_ascii=False)
        for mode in ("fast", "standard", "careful"):
            assert mode in schemas, f"missing PDF mode in MCP schema: {mode}"
        return {
            "server": initialized["result"]["serverInfo"],
            "pdf_tools": [tool["name"] for tool in pdf_tools],
            "modes": ["fast", "standard", "careful"],
        }
    finally:
        process.terminate()
        process.wait(timeout=10)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--gui", type=Path)
    parser.add_argument("--mcp", type=Path)
    args = parser.parse_args()
    if not args.gui and not args.mcp:
        parser.error("at least one of --gui or --mcp is required")
    result = {}
    if args.gui:
        result["gui_window"] = check_gui(args.gui.resolve())
    if args.mcp:
        result["mcp"] = check_mcp(args.mcp.resolve())
    print(json.dumps(result, ensure_ascii=False))


if __name__ == "__main__":
    main()
