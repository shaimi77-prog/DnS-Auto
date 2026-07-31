"""비GUI 작업의 공통 데이터 계약."""

from __future__ import annotations

from dataclasses import asdict, dataclass, field
from enum import Enum
from pathlib import Path
from typing import Any


class JobState(str, Enum):
    QUEUED = "queued"
    RUNNING = "running"
    SUCCEEDED = "succeeded"
    FAILED = "failed"
    CANCELLED = "cancelled"
    NEEDS_USER_ACTION = "needs_user_action"


@dataclass(frozen=True)
class ProgressEvent:
    completed: int
    total: int
    message: str
    current_file: str | None = None
    current_sheet: str | None = None
    activity: str | None = None
    elapsed_seconds: int = 0
    estimated_remaining_seconds: int | None = None
    estimate_status: str = "calculating"


@dataclass
class JobResult:
    state: JobState
    output_files: list[str] = field(default_factory=list)
    failed_files: list[str] = field(default_factory=list)
    message: str = ""
    details: dict[str, Any] = field(default_factory=dict)

    def as_dict(self) -> dict[str, Any]:
        result = asdict(self)
        result["state"] = self.state.value
        return result


def normalized_path(path: str | Path) -> str:
    return str(Path(path).expanduser().resolve())
