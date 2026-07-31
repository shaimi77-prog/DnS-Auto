"""Thread-safe cancellation and output rollback primitives."""

from __future__ import annotations

import os
import threading
import time
from pathlib import Path
from typing import Callable


class ProcessingCancellation:
    """Coordinate cancel-all and the irreversible final-save boundary."""

    def __init__(self):
        self._cancel_requested = threading.Event()
        self._lock = threading.RLock()
        self._save_started = False
        self._owned_outputs: list[Path] = []

    def request_cancel_all(self) -> bool:
        with self._lock:
            if self._save_started or self._cancel_requested.is_set():
                return False
            self._cancel_requested.set()
            return True

    def should_cancel(self) -> bool:
        return self._cancel_requested.is_set()

    @property
    def save_started(self) -> bool:
        with self._lock:
            return self._save_started

    def enter_save_phase(self) -> bool:
        with self._lock:
            if self._cancel_requested.is_set():
                return False
            self._save_started = True
            return True

    def reserve_output(self, path: str | os.PathLike[str]) -> Path:
        """Register a job-owned path before its writer can create a partial file."""
        output = Path(path)
        with self._lock:
            if output not in self._owned_outputs:
                self._owned_outputs.append(output)
        return output

    def created_outputs(self) -> list[str]:
        with self._lock:
            return [str(path) for path in self._owned_outputs]

    def rollback_outputs(
        self,
        attempts: int = 3,
        delay_seconds: float = 0.1,
        unlink: Callable[[Path], None] | None = None,
    ) -> list[tuple[str, str]]:
        remover = unlink or (lambda path: path.unlink(missing_ok=True))
        failures: list[tuple[str, str]] = []
        with self._lock:
            outputs = list(reversed(self._owned_outputs))
        for output in outputs:
            error = None
            for attempt in range(max(int(attempts), 1)):
                try:
                    remover(output)
                    error = None
                    break
                except OSError as exc:
                    error = exc
                    if attempt + 1 < attempts and delay_seconds:
                        time.sleep(delay_seconds)
            if error is not None and output.exists():
                failures.append((str(output), str(error)))
        return failures
