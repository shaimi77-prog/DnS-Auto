import json
import os
import subprocess
import sys
import time
from pathlib import Path
import unittest


ROOT = Path(__file__).resolve().parents[1]
SERVER = ROOT / "mcp_server.py"


class McpProtocolTests(unittest.TestCase):
    def setUp(self):
        self.process = subprocess.Popen(
            [sys.executable, "-u", str(SERVER)], stdin=subprocess.PIPE, stdout=subprocess.PIPE,
            text=True, encoding="utf-8"
        )

    def tearDown(self):
        if self.process.stdin:
            self.process.stdin.close()
        self.process.terminate()
        self.process.wait(timeout=5)
        if self.process.stdout:
            self.process.stdout.close()

    def request(self, payload):
        self.process.stdin.write(json.dumps(payload) + "\n")
        self.process.stdin.flush()
        return json.loads(self.process.stdout.readline())

    def test_initialize_and_tools_list(self):
        initialized = self.request({"jsonrpc": "2.0", "id": 1, "method": "initialize", "params": {"protocolVersion": "2024-11-05"}})
        self.assertEqual(initialized["result"]["serverInfo"]["name"], "dns-auto-mcp")
        self.assertEqual(initialized["result"]["serverInfo"]["version"], "1.0.0")
        listed = self.request({"jsonrpc": "2.0", "id": 2, "method": "tools/list"})
        self.assertEqual(
            {tool["name"] for tool in listed["result"]["tools"]},
            {"inspect_files", "discover_merge_plan", "start_document_conversion", "start_sheet_merge", "start_pdf_merge", "start_interactive_sheet_merge", "start_interactive_pdf_merge", "get_job_status", "get_job_result", "cancel_job"},
        )

    def test_rejects_path_outside_policy(self):
        response = self.request({"jsonrpc": "2.0", "id": 3, "method": "tools/call", "params": {"name": "inspect_files", "arguments": {"paths": [str(ROOT / "README.md")]}}})
        payload = json.loads(response["result"]["content"][0]["text"])
        self.assertTrue(response["result"]["isError"] is False)
        self.assertFalse(payload["files"][0]["ok"])

    def test_hwp_job_lifecycle_returns_explicit_user_action(self):
        """The MCP job API must remain observable even when HWP automation is disabled."""
        samples = sorted((Path(os.environ.get("DNS_AUTO_EMPIRICAL_ROOT", ROOT / "private_test_files"))).rglob("*.hwpx"))
        if not samples:
            self.skipTest("Legacy empirical HWPX fixture is not present in this project copy.")

        started = self.request({
            "jsonrpc": "2.0", "id": 4, "method": "tools/call",
            "params": {
                "name": "start_document_conversion",
                "arguments": {"kind": "hwp_to_pdf", "paths": [str(samples[0])]},
            },
        })
        payload = json.loads(started["result"]["content"][0]["text"])
        job_id = payload["job_id"]

        deadline = time.monotonic() + 5
        state = ""
        result = None
        while time.monotonic() < deadline:
            response = self.request({
                "jsonrpc": "2.0", "id": 5, "method": "tools/call",
                "params": {"name": "get_job_result", "arguments": {"job_id": job_id}},
            })
            payload = json.loads(response["result"]["content"][0]["text"])
            state, result = payload["state"], payload.get("result")
            if result is not None:
                break
            time.sleep(0.05)

        self.assertEqual(state, "needs_user_action")
        self.assertIsNotNone(result)
        self.assertTrue(result["message"])

    @unittest.skipUnless(
        os.environ.get("DNS_RUN_OFFICE_COM_TESTS") == "1",
        "Requires an interactive Windows desktop session with Microsoft Excel COM.",
    )
    def test_xls_conversion_job_creates_xlsx_output(self):
        samples = sorted((Path(os.environ.get("DNS_AUTO_EMPIRICAL_ROOT", ROOT / "private_test_files"))).rglob("*.xls"))
        self.assertTrue(samples, "An empirical XLS sample is required for this integration test.")

        started = self.request({
            "jsonrpc": "2.0", "id": 6, "method": "tools/call",
            "params": {
                "name": "start_document_conversion",
                "arguments": {"kind": "xls_to_xlsx", "paths": [str(samples[0])]},
            },
        })
        job_id = json.loads(started["result"]["content"][0]["text"])["job_id"]

        deadline = time.monotonic() + 45
        payload = None
        while time.monotonic() < deadline:
            response = self.request({
                "jsonrpc": "2.0", "id": 7, "method": "tools/call",
                "params": {"name": "get_job_result", "arguments": {"job_id": job_id}},
            })
            payload = json.loads(response["result"]["content"][0]["text"])
            if payload.get("result") is not None:
                break
            time.sleep(0.1)

        self.assertIsNotNone(payload)
        self.assertEqual(payload["state"], "succeeded", payload.get("result"))
        outputs = payload["result"]["output_files"]
        self.assertEqual(len(outputs), 1)
        self.assertTrue(Path(outputs[0]).is_file())
        self.assertGreater(Path(outputs[0]).stat().st_size, 0)

    def test_pdf_merge_job_creates_workbook_from_empirical_files(self):
        empirical = Path(os.environ.get("DNS_AUTO_EMPIRICAL_ROOT", ROOT / "private_test_files"))
        template = next(empirical.rglob("*.xlsx"), None)
        if template is None:
            self.skipTest("Legacy empirical PDF fixture is not present in this project copy.")
        input_dir = template.parent
        profile = next(empirical.rglob("시연용_다중시트.json"))
        pdfs_by_sheet = {
            month: [str(path) for path in sorted((input_dir / month).glob("*.pdf"))]
            for month in ("4월", "5월")
        }
        self.assertTrue(all(pdfs_by_sheet.values()))

        started = self.request({
            "jsonrpc": "2.0", "id": 8, "method": "tools/call",
            "params": {
                "name": "start_pdf_merge",
                "arguments": {
                    "template_path": str(template),
                    "pdfs_by_sheet": pdfs_by_sheet,
                    "profile_path": str(profile),
                },
            },
        })
        job_id = json.loads(started["result"]["content"][0]["text"])["job_id"]

        deadline = time.monotonic() + 30
        payload = None
        while time.monotonic() < deadline:
            response = self.request({
                "jsonrpc": "2.0", "id": 9, "method": "tools/call",
                "params": {"name": "get_job_result", "arguments": {"job_id": job_id}},
            })
            payload = json.loads(response["result"]["content"][0]["text"])
            if payload.get("result") is not None:
                break
            time.sleep(0.1)

        self.assertIsNotNone(payload)
        self.assertEqual(payload["state"], "succeeded", payload.get("result"))
        outputs = payload["result"]["output_files"]
        self.assertEqual(len(outputs), 1)
        self.assertTrue(Path(outputs[0]).is_file())
        self.assertGreater(Path(outputs[0]).stat().st_size, 0)


if __name__ == "__main__":
    unittest.main()
