import unittest
from unittest.mock import patch

from com_process_ownership import Ownership, ProcessIdentity, cleanup_com_session


class ComCleanupContractTests(unittest.TestCase):
    def test_close_failure_does_not_block_quit_or_uninitialize(self):
        calls = []

        def fail_close():
            calls.append("close")
            raise RuntimeError("close failed")

        result = cleanup_com_session(
            application="word",
            close_callbacks=[fail_close],
            quit_callback=lambda: calls.append("quit"),
            ownership=Ownership("unconfirmed"),
            co_uninitialize=lambda: calls.append("uninit"),
        )
        self.assertEqual(calls, ["close", "quit", "uninit"])
        self.assertEqual(result["close_status"], "failed")
        self.assertEqual(result["quit_status"], "completed")

    def test_unconfirmed_process_is_never_terminated(self):
        with patch("com_process_ownership.terminate_confirmed_process") as terminate:
            result = cleanup_com_session(
                application="excel",
                close_callbacks=[],
                quit_callback=lambda: None,
                ownership=Ownership("unconfirmed"),
                co_uninitialize=lambda: None,
                allow_forced_cleanup=True,
            )
        terminate.assert_not_called()
        self.assertEqual(result["process_exit_status"], "unknown")

    def test_confirmed_remaining_process_uses_limited_cleanup(self):
        owner = Ownership("confirmed", ProcessIdentity(123, "created", "hwp.exe"))
        with patch("com_process_ownership.wait_for_exit", return_value=False), patch(
            "com_process_ownership.terminate_confirmed_process", return_value=True
        ) as terminate:
            result = cleanup_com_session(
                application="hwp",
                close_callbacks=[],
                quit_callback=lambda: None,
                ownership=owner,
                co_uninitialize=lambda: None,
                allow_forced_cleanup=True,
                normal_exit_timeout=0,
                additional_exit_timeout=0,
            )
        terminate.assert_called_once_with(owner)
        self.assertEqual(result["forced_cleanup_status"], "completed")
        self.assertEqual(result["process_exit_status"], "exited")


if __name__ == "__main__":
    unittest.main()
