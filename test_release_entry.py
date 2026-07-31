from __future__ import annotations

import json
import subprocess
import time
from pathlib import Path

import win32gui

ROOT = Path(__file__).resolve().parent.parent
BUNDLE = ROOT / "release" / "DnS Auto"
GUI = BUNDLE / "DnS Auto.exe"
MCP = BUNDLE / "DnS Auto MCP.exe"

existing: set[int] = set()
def remember(handle, _extra):
    if "DnS Auto" in win32gui.GetWindowText(handle):
        existing.add(handle)
win32gui.EnumWindows(remember, None)
process = subprocess.Popen([str(GUI)], cwd=ROOT)
title = ""
try:
    deadline = time.monotonic() + 30
    while time.monotonic() < deadline:
        found: list[str] = []
        def collect(handle, _extra):
            text = win32gui.GetWindowText(handle)
            if handle not in existing and "DnS Auto" in text:
                found.append(text)
        win32gui.EnumWindows(collect, None)
        if found:
            title = found[0]
            break
        if process.poll() is not None:
            break
        time.sleep(0.1)
    assert title, f"GUI 창을 찾지 못했습니다: exit={process.poll()}"
finally:
    process.terminate()
    try:
        process.wait(timeout=10)
    except subprocess.TimeoutExpired:
        process.kill()
        process.wait(timeout=5)

mcp = subprocess.Popen([str(MCP)], cwd=ROOT, stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, encoding="utf-8")
def request(identifier, method, params=None):
    payload = {"jsonrpc": "2.0", "id": identifier, "method": method}
    if params is not None:
        payload["params"] = params
    mcp.stdin.write(json.dumps(payload, ensure_ascii=False) + "\n")
    mcp.stdin.flush()
    line = mcp.stdout.readline()
    if not line:
        raise RuntimeError(mcp.stderr.read())
    return json.loads(line)
try:
    initialized = request(1, "initialize", {"protocolVersion": "2024-11-05"})
    assert initialized["result"]["serverInfo"]["name"] == "dns-auto-mcp"
    tools = {tool["name"] for tool in request(2, "tools/list")["result"]["tools"]}
    assert {"discover_merge_plan", "start_pdf_merge", "start_interactive_pdf_merge"} <= tools
finally:
    mcp.terminate()
    mcp.wait(timeout=10)
print(json.dumps({"gui_window": title, "mcp_initialize": "PASS", "mcp_pdf_tools": "PASS"}, ensure_ascii=False))