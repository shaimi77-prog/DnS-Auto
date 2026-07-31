"""MCP의 장기 작업을 추적하는 메모리 기반 작업 관리자."""

from __future__ import annotations

import logging
import threading
import uuid
from dataclasses import dataclass, field
from typing import Callable

from core.models import JobResult, JobState, ProgressEvent
from processing_cancellation import ProcessingCancellation


@dataclass
class ManagedJob:
    job_id: str
    state: JobState = JobState.QUEUED
    progress: ProgressEvent | None = None
    result: JobResult | None = None
    events: list[ProgressEvent] = field(default_factory=list)
    cancellation: ProcessingCancellation | None = None


class JobManager:
    def __init__(self) -> None:
        self._jobs: dict[str, ManagedJob] = {}
        self._lock = threading.Lock()

    def start(
        self,
        worker: Callable[[Callable[[ProgressEvent], None]], JobResult],
        initial_state: JobState = JobState.RUNNING,
        cancellation: ProcessingCancellation | None = None,
    ) -> str:
        job = ManagedJob(job_id=uuid.uuid4().hex, state=initial_state, cancellation=cancellation)
        with self._lock:
            self._jobs[job.job_id] = job

        def report(event: ProgressEvent) -> None:
            with self._lock:
                job.progress = event
                job.events.append(event)

        def run() -> None:
            with self._lock:
                job.state = initial_state
            try:
                result = worker(report)
                with self._lock:
                    job.result = result
                    job.state = result.state
            except Exception as error:  # MCP에는 예외 대신 작업 결과로 보고한다.
                logging.exception("MCP 작업 실패: %s", job.job_id)
                with self._lock:
                    job.result = JobResult(state=JobState.FAILED, message=str(error))
                    job.state = JobState.FAILED

        threading.Thread(target=run, daemon=True, name=f"dns-job-{job.job_id[:8]}").start()
        return job.job_id

    def get(self, job_id: str) -> ManagedJob:
        with self._lock:
            try:
                return self._jobs[job_id]
            except KeyError as error:
                raise ValueError(f"알 수 없는 작업 ID입니다: {job_id}") from error


    def cancel(self, job_id: str) -> bool:
        with self._lock:
            try:
                job = self._jobs[job_id]
            except KeyError as error:
                raise ValueError(f"Unknown job ID: {job_id}") from error
            if job.state in {JobState.SUCCEEDED, JobState.FAILED, JobState.CANCELLED}:
                return False
            cancellation = job.cancellation
        return cancellation.request_cancel_all() if cancellation is not None else False
