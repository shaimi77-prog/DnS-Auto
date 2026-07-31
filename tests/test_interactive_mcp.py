import json
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import mcp_server
from core.models import JobResult, JobState
from core.job_manager import JobManager


class InteractiveMcpTests(unittest.TestCase):
    def test_tool_schemas_expose_profileless_flows(self):
        tools = {tool["name"]: tool for tool in mcp_server.TOOLS}
        self.assertEqual(tools["start_interactive_sheet_merge"]["inputSchema"]["required"], ["template_path", "source_paths"])
        self.assertNotIn("profile_path", tools["start_interactive_sheet_merge"]["inputSchema"]["properties"])
        self.assertEqual(tools["start_interactive_pdf_merge"]["inputSchema"]["required"], ["template_path"])

    def test_interactive_job_is_immediately_waiting_for_user(self):
        manager = JobManager()
        with patch.object(mcp_server, "JOBS", manager), patch.object(mcp_server, "run_interactive", return_value=JobResult(state=JobState.SUCCEEDED)):
            inputs = mcp_server.APP_ROOT / "inputs"
            inputs.mkdir(exist_ok=True)
            template = inputs / "interactive_template.xlsx"
            source = inputs / "interactive_source.xlsx"
            template.write_bytes(b"test")
            source.write_bytes(b"test")
            try:
                response = mcp_server.call("start_interactive_sheet_merge", {"template_path": str(template), "source_paths": [str(source)]})
                self.assertEqual(response["state"], "needs_user_action")
                self.assertTrue(response["user_action"])
                job = manager.get(response["job_id"])
                self.assertIn(job.state, {JobState.NEEDS_USER_ACTION, JobState.SUCCEEDED})
            finally:
                template.unlink(missing_ok=True)
                source.unlink(missing_ok=True)

    def test_run_interactive_reads_child_result_contract(self):
        with tempfile.TemporaryDirectory(dir=mcp_server.APP_ROOT) as directory:
            root = Path(directory)
            payload = {"mode": "sheet", "template_path": str(root / "template.xlsx"), "source_paths": [], "output_root": str(root / "outputs")}
            def fake_run(command, **kwargs):
                request = json.loads(Path(command[-1]).read_text(encoding="utf-8"))
                Path(request["result_path"]).write_text(json.dumps({"state": "succeeded", "output_files": [str(root / "outputs" / "result.xlsx")], "message": "ok", "details": {"interactive": True}}), encoding="utf-8")
                class Completed:
                    returncode = 0
                    stderr = ""
                return Completed()
            with patch("mcp_server.subprocess.run", side_effect=fake_run):
                result = mcp_server.run_interactive(payload)
            self.assertEqual(result.state, JobState.SUCCEEDED)
            self.assertEqual(result.message, "ok")
            self.assertTrue(result.details["interactive"])


if __name__ == "__main__":
    unittest.main()