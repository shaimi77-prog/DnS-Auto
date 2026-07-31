import tempfile
import threading
import unittest
from pathlib import Path

from processing_cancellation import ProcessingCancellation


class ProcessingCancellationTests(unittest.TestCase):
    def test_cancel_blocks_save(self):
        cancellation = ProcessingCancellation()
        self.assertTrue(cancellation.request_cancel_all())
        self.assertFalse(cancellation.enter_save_phase())

    def test_save_blocks_late_cancel(self):
        cancellation = ProcessingCancellation()
        self.assertTrue(cancellation.enter_save_phase())
        self.assertFalse(cancellation.request_cancel_all())
        self.assertFalse(cancellation.should_cancel())

    def test_request_is_idempotent(self):
        cancellation = ProcessingCancellation()
        self.assertTrue(cancellation.request_cancel_all())
        self.assertFalse(cancellation.request_cancel_all())

    def test_rollback_removes_only_registered_outputs(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            existing = root / "existing.pdf"
            created = root / "created.pdf"
            existing.write_text("keep", encoding="utf-8")
            created.write_text("remove", encoding="utf-8")
            cancellation = ProcessingCancellation()
            cancellation.reserve_output(created)
            self.assertEqual([], cancellation.rollback_outputs(delay_seconds=0))
            self.assertTrue(existing.exists())
            self.assertFalse(created.exists())

    def test_registration_and_cancel_are_thread_safe(self):
        cancellation = ProcessingCancellation()
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "race.pdf"
            barrier = threading.Barrier(2)

            def register():
                barrier.wait()
                cancellation.reserve_output(output)

            thread = threading.Thread(target=register)
            thread.start()
            barrier.wait()
            cancellation.request_cancel_all()
            thread.join()
            self.assertIn(str(output), cancellation.created_outputs())

    def test_cleanup_failure_is_reported(self):
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "locked.pdf"
            output.write_text("locked", encoding="utf-8")
            cancellation = ProcessingCancellation()
            cancellation.reserve_output(output)

            def fail(_path):
                raise PermissionError("locked")

            failures = cancellation.rollback_outputs(
                attempts=2, delay_seconds=0, unlink=fail
            )
            self.assertEqual(str(output), failures[0][0])


if __name__ == "__main__":
    unittest.main()
