"""Codex 등 MCP 클라이언트가 DnS Auto 서비스를 호출하는 stdio 서버."""

from __future__ import annotations

import json
import os
import subprocess
import sys
import uuid
from pathlib import Path
from typing import Any

from core.job_manager import JobManager
from core.models import JobResult, JobState
from processing_cancellation import ProcessingCancellation
from services import conversion_service, discovery_service, pdf_service, sheet_service
from services.preflight_service import inspect_paths


def application_root() -> Path:
    """Return the user-visible directory, including from a PyInstaller bundle."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


APP_ROOT = application_root()
RESOURCE_ROOT = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
JOBS = JobManager()
TOOLS = [
    {"name": "inspect_files", "description": "허용 경로·확장자·파일 존재 여부를 읽기 전용으로 점검합니다.", "inputSchema": {"type": "object", "properties": {"paths": {"type": "array", "items": {"type": "string"}}}, "required": ["paths"]}},
    {"name": "discover_merge_plan", "description": "inputs와 profiles를 재귀 검색해 작업 종류·양식·원본·호환 프로필을 판별합니다. 임의 선택 없이 ready, needs_confirmation 또는 needs_clarification과 질문 목록을 반환합니다.", "inputSchema": {"type": "object", "properties": {"input_root": {"type": "string", "description": "생략하면 포터블 폴더의 inputs를 사용합니다."}, "operation": {"enum": ["auto", "excel", "pdf"], "default": "auto"}, "template_path": {"type": "string"}, "profile_name": {"type": "string"}, "interactive": {"type": "boolean", "default": False}}, "additionalProperties": False}},
    {"name": "start_document_conversion", "description": "DOCX·HWP 또는 XLS 변환 작업을 시작하고 작업 ID를 반환합니다.", "inputSchema": {"type": "object", "properties": {"kind": {"enum": ["docx_to_pdf", "hwp_to_pdf", "xls_to_xlsx"]}, "paths": {"type": "array", "items": {"type": "string"}}}, "required": ["kind", "paths"]}},
    {"name": "start_sheet_merge", "description": "저장된 Sheet 설정 프로필로 Excel 취합 작업을 시작합니다.", "inputSchema": {"type": "object", "properties": {"template_path": {"type": "string"}, "source_paths": {"type": "array", "items": {"type": "string"}}, "profile_path": {"type": "string"}}, "required": ["template_path", "source_paths", "profile_path"]}},
    {"name": "start_pdf_merge", "description": "저장된 PDF 매핑 프로필로 PDF/OCR 취합 작업을 시작합니다.", "inputSchema": {"type": "object", "properties": {"template_path": {"type": "string"}, "pdfs_by_sheet": {"type": "object"}, "profile_path": {"type": "string"}, "force_ocr": {"type": "boolean"}}, "required": ["template_path", "pdfs_by_sheet", "profile_path"]}},
    {"name": "start_interactive_sheet_merge", "description": "프로필 없이 Excel 설정 창을 열어 사용자가 시트·헤더·취합 영역을 지정한 뒤 취합합니다.", "inputSchema": {"type": "object", "properties": {"template_path": {"type": "string"}, "source_paths": {"type": "array", "items": {"type": "string"}}}, "required": ["template_path", "source_paths"]}},
    {"name": "start_interactive_pdf_merge", "description": "프로필 없이 PDF 설정·영역 드래그 창을 열어 사용자가 매핑한 뒤 PDF/OCR 취합을 계속합니다.", "inputSchema": {"type": "object", "properties": {"template_path": {"type": "string"}, "pdfs_by_sheet": {"type": "object"}, "force_ocr": {"type": "boolean"}}, "required": ["template_path"]}},
    {"name": "get_job_status", "description": "작업의 현재 상태와 마지막 진행 이벤트를 조회합니다.", "inputSchema": {"type": "object", "properties": {"job_id": {"type": "string"}}, "required": ["job_id"]}},
    {"name": "get_job_result", "description": "완료 작업의 산출물·실패 목록·요약을 조회합니다.", "inputSchema": {"type": "object", "properties": {"job_id": {"type": "string"}}, "required": ["job_id"]}},
    {"name": "cancel_job", "description": "Request cancellation of a running job after the current OCR or Office call finishes; no result file is saved.", "inputSchema": {"type": "object", "properties": {"job_id": {"type": "string"}}, "required": ["job_id"]}},
]


def policy() -> dict[str, Any]:
    policy_path = Path(os.environ.get("DNS_AUTO_MCP_POLICY", APP_ROOT / "mcp_policy.json"))
    if not policy_path.is_file():
        policy_path = RESOURCE_ROOT / "mcp_policy.json"
    with policy_path.open(encoding="utf-8") as file:
        settings = json.load(file)

    # Relative policy paths are anchored to the executable, not the caller's CWD.
    def resolved(value: str) -> str:
        path = Path(os.path.expandvars(value)).expanduser()
        return str((APP_ROOT / path).resolve() if not path.is_absolute() else path.resolve())

    settings["allowed_input_roots"] = [
        resolved(value) for value in settings.get("allowed_input_roots", ["inputs"])
    ]
    settings["output_root"] = resolved(settings.get("output_root", "outputs"))
    return settings


def checked(paths: list[str], settings: dict[str, Any]) -> list[str]:
    findings = inspect_paths(paths, settings["allowed_input_roots"])
    rejected = [item for item in findings if not item["ok"]]
    if rejected:
        raise ValueError("입력 파일 정책 검증 실패: " + json.dumps(rejected, ensure_ascii=False))
    return [str(Path(path).resolve()) for path in paths]


def content(value: dict[str, Any], error: bool = False) -> dict[str, Any]:
    return {"content": [{"type": "text", "text": json.dumps(value, ensure_ascii=False)}], "isError": error}


def _interactive_command(request_path: Path) -> list[str]:
    if getattr(sys, "frozen", False):
        return [sys.executable, "--interactive-request", str(request_path)]
    return [sys.executable, str(Path(__file__).resolve()), "--interactive-request", str(request_path)]


def run_interactive(payload: dict[str, Any]) -> JobResult:
    jobs_root = APP_ROOT / "jobs"
    jobs_root.mkdir(parents=True, exist_ok=True)
    token = uuid.uuid4().hex
    request_path = jobs_root / f"{token}.request.json"
    result_path = jobs_root / f"{token}.result.json"
    payload = {**payload, "result_path": str(result_path)}
    request_path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
    try:
        completed = subprocess.run(
            _interactive_command(request_path),
            cwd=str(APP_ROOT),
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            check=False,
        )
        if not result_path.is_file():
            message = completed.stderr.strip() or f"대화형 GUI가 결과 없이 종료되었습니다. 종료 코드: {completed.returncode}"
            return JobResult(state=JobState.FAILED, message=message, details={"interactive": True})
        data = json.loads(result_path.read_text(encoding="utf-8"))
        return JobResult(
            state=JobState(data.get("state", "failed")),
            output_files=list(data.get("output_files", [])),
            failed_files=list(data.get("failed_files", [])),
            message=str(data.get("message", "")),
            details=dict(data.get("details", {})),
        )
    finally:
        request_path.unlink(missing_ok=True)
        result_path.unlink(missing_ok=True)

def call(name: str, args: dict[str, Any]) -> dict[str, Any]:
    settings = policy()
    output = settings["output_root"]
    if name == "inspect_files":
        return {"files": inspect_paths(args["paths"], settings["allowed_input_roots"])}
    if name == "discover_merge_plan":
        return discovery_service.discover_merge_plan(APP_ROOT, output, args.get("input_root"), args.get("operation", "auto"), args.get("template_path"), args.get("profile_name"), bool(args.get("interactive")))
    if name == "start_document_conversion":
        paths = checked(args["paths"], settings)
        worker = {
            "docx_to_pdf": conversion_service.convert_docx_to_pdf,
            "hwp_to_pdf": conversion_service.convert_hwp_to_pdf,
            "xls_to_xlsx": conversion_service.convert_xls_to_xlsx,
        }[args["kind"]]
        cancellation = ProcessingCancellation()
        return {"job_id": JOBS.start(
            lambda report: worker(paths, output, report, cancellation),
            cancellation=cancellation,
        )}
    if name == "start_sheet_merge":
        paths = checked([args["template_path"], args["profile_path"], *args["source_paths"]], settings)
        cancellation = ProcessingCancellation()
        return {"job_id": JOBS.start(
            lambda report: sheet_service.merge_workbooks(
                paths[0], paths[2:], paths[1], output, report, cancellation=cancellation
            ), cancellation=cancellation
        )}
    if name == "start_pdf_merge":
        raw_pdfs = [path for paths in args["pdfs_by_sheet"].values() for path in paths]
        paths = checked([args["template_path"], args["profile_path"], *raw_pdfs], settings)
        checked_pdfs, offset = {}, 2
        for sheet_name, sheet_paths in args["pdfs_by_sheet"].items():
            checked_pdfs[sheet_name] = paths[offset:offset + len(sheet_paths)]
            offset += len(sheet_paths)
        cancellation = ProcessingCancellation()
        return {"job_id": JOBS.start(
            lambda report: pdf_service.merge_pdfs(
                paths[0], checked_pdfs, paths[1], output,
                bool(args.get("force_ocr")), report, cancellation
            ), cancellation=cancellation
        )}
    if name == "start_interactive_sheet_merge":
        paths = checked([args["template_path"], *args["source_paths"]], settings)
        payload = {"mode": "sheet", "template_path": paths[0], "source_paths": paths[1:], "output_root": output}
        return {"job_id": JOBS.start(lambda report: run_interactive(payload), initial_state=JobState.NEEDS_USER_ACTION), "state": JobState.NEEDS_USER_ACTION.value, "user_action": "Excel 설정 창에서 시트, 헤더 범위와 취합 방식을 지정해 주세요."}
    if name == "start_interactive_pdf_merge":
        pdfs_by_sheet = args.get("pdfs_by_sheet", {})
        raw_pdfs = [path for paths in pdfs_by_sheet.values() for path in paths]
        paths = checked([args["template_path"], *raw_pdfs], settings)
        checked_pdfs = {}
        offset = 1
        for sheet_name, sheet_paths in pdfs_by_sheet.items():
            checked_pdfs[sheet_name] = paths[offset:offset + len(sheet_paths)]
            offset += len(sheet_paths)
        payload = {"mode": "pdf", "template_path": paths[0], "pdfs_by_sheet": checked_pdfs, "force_ocr": bool(args.get("force_ocr")), "output_root": output}
        return {"job_id": JOBS.start(lambda report: run_interactive(payload), initial_state=JobState.NEEDS_USER_ACTION), "state": JobState.NEEDS_USER_ACTION.value, "user_action": "PDF 설정 창에서 시트·헤더를 지정하고 각 필드의 PDF 영역을 드래그해 주세요."}
    if name == "cancel_job":
        requested = JOBS.cancel(args["job_id"])
        job = JOBS.get(args["job_id"])
        return {
            "job_id": job.job_id, "cancel_requested": requested, "state": job.state.value,
            "message": ("Cancellation requested; it will take effect after the current external call." if requested else "The job can no longer be cancelled."),
        }
    if name in {"get_job_status", "get_job_result"}:
        job = JOBS.get(args["job_id"])
        result = {"job_id": job.job_id, "state": job.state.value}
        if job.progress:
            result["progress"] = job.progress.__dict__
        if name == "get_job_result":
            result["result"] = job.result.as_dict() if job.result else None
        return result
    raise ValueError(f"지원하지 않는 도구입니다: {name}")


def handle(request: dict[str, Any]) -> dict[str, Any] | None:
    if request.get("id") is None:
        return None
    request_id, method = request["id"], request.get("method")
    if method == "initialize":
        return {"jsonrpc": "2.0", "id": request_id, "result": {"protocolVersion": request.get("params", {}).get("protocolVersion", "2024-11-05"), "capabilities": {"tools": {}}, "serverInfo": {"name": "dns-auto-mcp", "version": "1.0.0"}}}
    if method == "tools/list":
        return {"jsonrpc": "2.0", "id": request_id, "result": {"tools": TOOLS}}
    if method == "tools/call":
        try:
            params = request.get("params", {})
            return {"jsonrpc": "2.0", "id": request_id, "result": content(call(params.get("name", ""), params.get("arguments", {})))}
        except (ValueError, KeyError, OSError, TypeError, json.JSONDecodeError) as error:
            return {"jsonrpc": "2.0", "id": request_id, "result": content({"error": str(error)}, True)}
    return {"jsonrpc": "2.0", "id": request_id, "error": {"code": -32601, "message": "Method not found"}}


def main() -> None:
    if len(sys.argv) == 3 and sys.argv[1] == "--interactive-request":
        from interactive_runner import run_request
        raise SystemExit(run_request(sys.argv[2]))
    sys.stdin.reconfigure(encoding="utf-8")
    sys.stdout.reconfigure(encoding="utf-8")
    for line in sys.stdin:
        try:
            response = handle(json.loads(line))
            if response is not None:
                print(json.dumps(response, ensure_ascii=False), flush=True)
        except json.JSONDecodeError:
            print(json.dumps({"jsonrpc": "2.0", "id": None, "error": {"code": -32700, "message": "Parse error"}}), flush=True)


if __name__ == "__main__":
    main()
