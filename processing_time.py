"""GUI와 MCP가 함께 사용하는 작업 유형별 잔여 시간 추정기."""

from __future__ import annotations

import statistics
import time
from collections import defaultdict
from collections.abc import Iterable


class ProcessingTimeEstimator:
    """현재 작업에서 관측한 같은 유형의 처리시간으로 ETA를 계산한다."""

    def __init__(
        self,
        planned_work: Iterable[tuple[str, float]] = (),
        *,
        clock=time.monotonic,
        minimum_samples=2,
        sample_limit=8,
    ):
        self.planned_work = [
            (str(kind), max(float(weight), 1.0))
            for kind, weight in planned_work
        ]
        self.clock = clock
        self.minimum_samples = max(int(minimum_samples), 1)
        self.sample_limit = max(int(sample_limit), 1)
        self.samples = defaultdict(list)
        self.samples["native_text"]
        self.samples["ocr"]
        self.completed = 0
        self.started_at = None
        self.unit_started_at = None
        self.current = ("unknown", 1.0)

    def start(self):
        """COM/OCR 초기화를 포함한 전체 경과시간 측정을 시작한다."""
        if self.started_at is None:
            self.started_at = self.clock()

    def begin(self, work_type="unknown", weight=1):
        self.start()
        self.unit_started_at = self.clock()
        self.current = (str(work_type), max(float(weight), 1.0))

    def _record(self, kind, weight, duration):
        self.samples[kind].append(duration / weight)
        self.samples[kind] = self.samples[kind][-self.sample_limit :]

    def complete(
        self,
        *,
        work_type=None,
        weight=None,
        duration_seconds=None,
        initialization_seconds=0,
        ocr_initialization_seconds=0,
        successful=True,
    ):
        """현재 단위를 완료하고 성공한 처리시간만 표본에 추가한다."""
        kind = str(work_type or self.current[0])
        unit_weight = max(float(weight or self.current[1]), 1.0)
        if duration_seconds is None:
            duration = (
                0.0
                if self.unit_started_at is None
                else self.clock() - self.unit_started_at
            )
        else:
            duration = float(duration_seconds)
        excluded = max(float(initialization_seconds), 0.0) + max(
            float(ocr_initialization_seconds), 0.0
        )
        duration = max(duration - excluded, 0.0)
        if successful:
            self._record(kind, unit_weight, duration)

        self.completed += 1
        if self.completed <= len(self.planned_work):
            self.planned_work[self.completed - 1] = (kind, unit_weight)
        self.unit_started_at = None

    def observe(
        self,
        *,
        work_type="ocr",
        weight=1,
        duration_seconds=0,
        successful=True,
    ):
        """완료 개수를 늘리지 않고 추가 작업의 실측시간을 학습한다."""
        if not successful:
            return
        kind = str(work_type)
        unit_weight = max(float(weight), 1.0)
        duration = max(float(duration_seconds), 0.0)
        self._record(kind, unit_weight, duration)

    def discard_last_sample(self, work_type):
        """완료 직후 취소된 단위의 방금 기록한 정상 표본을 철회한다."""
        values = self.samples.get(str(work_type))
        if values:
            values.pop()

    def _average(self, kind):
        values = self.samples.get(kind, ())
        if len(values) < self.minimum_samples:
            return None
        median = statistics.median(values)
        deviations = [abs(value - median) for value in values]
        mad = statistics.median(deviations)
        filtered = [
            value
            for value in values
            if not mad or abs(value - median) <= 3 * mad
        ]
        weights = range(1, len(filtered) + 1)
        return sum(value * weight for value, weight in zip(filtered, weights)) / sum(
            weights
        )

    def metadata(self):
        elapsed = 0 if self.started_at is None else self.clock() - self.started_at
        remaining = self.planned_work[self.completed :]
        if not remaining:
            return {
                "elapsed_seconds": int(elapsed),
                "estimated_remaining_seconds": 0,
                "estimate_status": "available",
            }
        if any(kind == "unknown" for kind, _weight in remaining):
            return {
                "elapsed_seconds": int(elapsed),
                "estimated_remaining_seconds": None,
                "estimate_status": "calculating",
            }

        averages = {kind: self._average(kind) for kind, _weight in remaining}
        estimate = (
            None
            if any(average is None for average in averages.values())
            else sum(averages[kind] * weight for kind, weight in remaining)
        )
        return {
            "elapsed_seconds": int(elapsed),
            "estimated_remaining_seconds": (
                None if estimate is None else int(round(estimate))
            ),
            "estimate_status": "available" if estimate is not None else "calculating",
        }

    def summary(self):
        """개인정보 없이 최종 JobResult에 넣을 수 있는 시간 통계."""
        metadata = self.metadata()
        return {
            "elapsed_seconds": metadata["elapsed_seconds"],
            "sample_counts": {
                kind: len(values)
                for kind, values in sorted(self.samples.items())
                if values
            },
        }
