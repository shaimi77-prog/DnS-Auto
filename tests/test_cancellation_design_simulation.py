"""취소 기능 설계 검증용 순수 상태 시뮬레이션."""

import unittest
from dataclasses import dataclass, field


@dataclass(frozen=True)
class WorkToken:
    sequence: int
    kind: str
    sheet: str
    path: str


@dataclass
class CancellationState:
    current: WorkToken | None = None
    skip_for: WorkToken | None = None
    cancel_all: bool = False
    saving: bool = False

    def begin(self, token):
        self.current = token

    def request_skip(self):
        if self.current is not None and not self.saving:
            self.skip_for = self.current

    def decision(self, token):
        if self.cancel_all:
            return "cancel_all"
        return "skip" if self.skip_for == token else "commit"

    def finish(self, token):
        if self.skip_for == token:
            self.skip_for = None
        if self.current == token:
            self.current = None

    def enter_save(self):
        if self.cancel_all:
            return False
        self.saving = True
        return True


@dataclass
class FileTransaction:
    changes: list[tuple[str, int, int, object]] = field(default_factory=list)
    append_deltas: dict[str, int] = field(default_factory=dict)
    summary_delta: dict[str, int] = field(default_factory=dict)
    eta_samples: list[float] = field(default_factory=list)

    def stage_append(self, sheet, initial_row, value):
        offset = self.append_deltas.get(sheet, 0)
        row = initial_row + offset
        self.append_deltas[sheet] = offset + 1
        self.changes.append((sheet, row, 1, value))

    def commit(self, workbook, append_state, summary, eta):
        for sheet, row, column, value in self.changes:
            workbook[(sheet, row, column)] = value
        for sheet, delta in self.append_deltas.items():
            append_state[sheet] += delta
        for name, delta in self.summary_delta.items():
            summary[name] = summary.get(name, 0) + delta
        eta.extend(self.eta_samples)


class CancellationDesignSimulationTests(unittest.TestCase):
    def test_direct_excel_write_leaves_partial_cells(self):
        workbook = {}
        workbook[("총무과", 2, 3)] = "부분 반영"
        # 현재 구현에서 이 시점에 건너뛰면 이미 쓴 셀이 남는다.
        self.assertIn(("총무과", 2, 3), workbook)

    def test_discarded_transaction_changes_nothing_including_append_cursor(self):
        workbook, append_state, summary, eta = {}, {"총무과": 2}, {}, []
        transaction = FileTransaction()
        transaction.stage_append("총무과", append_state["총무과"], "기관A")
        transaction.summary_delta["processed"] = 1
        transaction.eta_samples.append(3.0)
        # discard: commit을 호출하지 않는다.
        self.assertEqual(workbook, {})
        self.assertEqual(append_state, {"총무과": 2})
        self.assertEqual(summary, {})
        self.assertEqual(eta, [])

    def test_committed_transaction_updates_cells_and_cursor_together(self):
        workbook, append_state, summary, eta = {}, {"총무과": 2}, {}, []
        transaction = FileTransaction()
        transaction.stage_append("총무과", append_state["총무과"], "기관A")
        transaction.commit(workbook, append_state, summary, eta)
        self.assertEqual(workbook[("총무과", 2, 1)], "기관A")
        self.assertEqual(append_state["총무과"], 3)

    def test_skip_is_bound_to_one_assignment_not_same_physical_pdf(self):
        state = CancellationState()
        april = WorkToken(1, "pdf", "4월", "same.pdf")
        may = WorkToken(2, "pdf", "5월", "same.pdf")
        state.begin(april)
        state.request_skip()
        self.assertEqual(state.decision(april), "skip")
        state.finish(april)
        state.begin(may)
        self.assertEqual(state.decision(may), "commit")

    def test_repeated_skip_click_does_not_leak_to_next_file(self):
        state = CancellationState()
        first = WorkToken(1, "excel", "*", "first.xlsx")
        second = WorkToken(2, "excel", "*", "second.xlsx")
        state.begin(first)
        state.request_skip()
        state.request_skip()
        state.finish(first)
        state.begin(second)
        self.assertEqual(state.decision(second), "commit")

    def test_one_excel_file_is_atomic_across_selected_sheets(self):
        workbook, append_state, summary, eta = {}, {}, {}, []
        transaction = FileTransaction(
            changes=[
                ("총무과", 2, 3, "A"),
                ("보안과", 2, 3, "B"),
                ("의료과", 2, 3, "C"),
            ]
        )
        # skip: 어느 시트에도 커밋하지 않는다.
        self.assertEqual(workbook, {})
        transaction.commit(workbook, append_state, summary, eta)
        self.assertEqual(len(workbook), 3)

    def test_cancel_all_blocks_save(self):
        state = CancellationState(cancel_all=True)
        self.assertFalse(state.enter_save())

    def test_skip_request_is_ignored_after_save_phase_starts(self):
        state = CancellationState()
        token = WorkToken(1, "excel", "*", "last.xlsx")
        state.begin(token)
        self.assertTrue(state.enter_save())
        state.request_skip()
        self.assertEqual(state.decision(token), "commit")


if __name__ == "__main__":
    unittest.main()
