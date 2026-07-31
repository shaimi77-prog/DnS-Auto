"""GUI와 MCP가 함께 사용하는 작업 유형별 잔여 시간 추정기."""
import statistics
import time

class ProcessingTimeEstimator:
    def __init__(self, planned_work=(), *, clock=time.monotonic, minimum_samples=2):
        self.planned_work=list(planned_work); self.clock=clock; self.minimum_samples=minimum_samples
        self.samples={"native_text": [], "ocr": []}; self.completed=0; self.started_at=None; self.unit_started_at=None; self.current=("unknown",1.0)
    def begin(self, work_type="unknown", weight=1):
        now=self.clock(); self.started_at=now if self.started_at is None else self.started_at; self.unit_started_at=now; self.current=(work_type,max(float(weight),1.0))
    def complete(self, *, work_type=None, weight=None, duration_seconds=None, ocr_initialization_seconds=0):
        kind=work_type or self.current[0]; weight=max(float(weight or self.current[1]),1.0)
        duration=(self.clock()-self.unit_started_at if duration_seconds is None else duration_seconds)
        duration=max(float(duration)-max(float(ocr_initialization_seconds),0),0)
        if kind in self.samples:
            self.samples[kind].append(duration/weight); self.samples[kind]=self.samples[kind][-8:]
        self.completed += 1
        if self.completed <= len(self.planned_work): self.planned_work[self.completed-1]=(kind,weight)
    def _average(self, kind):
        values=self.samples[kind]
        if len(values)<self.minimum_samples: return None
        median=statistics.median(values); deviations=[abs(v-median) for v in values]; mad=statistics.median(deviations)
        filtered=[v for v in values if not mad or abs(v-median)<=3*mad]; weights=range(1,len(filtered)+1)
        return sum(v*w for v,w in zip(filtered,weights))/sum(weights)
    def metadata(self):
        remaining=self.planned_work[self.completed:]; averages={k:self._average(k) for k in self.samples}
        if any(k == "unknown" for k, w in remaining): return {"elapsed_seconds":int(0 if self.started_at is None else self.clock()-self.started_at),"estimated_remaining_seconds":None,"estimate_status":"calculating"}
        needed={k for k,w in remaining if k in averages}; available=all(averages[k] is not None for k in needed)
        estimate=sum((averages[k] or 0)*w for k,w in remaining if k in averages) if available else None
        elapsed=0 if self.started_at is None else self.clock()-self.started_at
        return {"elapsed_seconds":int(elapsed),"estimated_remaining_seconds":None if estimate is None else int(round(estimate)),"estimate_status":"available" if estimate is not None else "calculating"}
