"""기준 PDF 파일·페이지 탐색 상태 모델."""
# Copyright (C) 2026 두부코드(DOOBOO_CODE)
# SPDX-License-Identifier: AGPL-3.0-only

from __future__ import annotations

import os
from dataclasses import dataclass


@dataclass
class ReferenceNavigationState:
    pdf_paths: tuple[str, ...]
    page_counts: tuple[int, ...]
    file_index: int = 0
    page_index: int = 0

    def __post_init__(self):
        self.pdf_paths = tuple(os.path.abspath(path) for path in self.pdf_paths)
        self.page_counts = tuple(int(count) for count in self.page_counts)
        if not self.pdf_paths or len(self.pdf_paths) != len(self.page_counts):
            raise ValueError("기준 페이지 탐색 파일과 페이지 수가 올바르지 않습니다.")
        if any(count < 1 for count in self.page_counts):
            raise ValueError("페이지가 없는 PDF는 기준 페이지로 탐색할 수 없습니다.")
        self.file_index = min(max(int(self.file_index), 0), len(self.pdf_paths) - 1)
        self.page_index = min(
            max(int(self.page_index), 0), self.page_counts[self.file_index] - 1
        )

    @classmethod
    def from_suggestion(cls, pdf_paths, page_counts, suggestion=None):
        normalized = tuple(os.path.abspath(path) for path in pdf_paths)
        file_index = 0
        page_index = 0
        if suggestion:
            suggested_path = os.path.abspath(suggestion.get("pdf_path", ""))
            if suggested_path in normalized:
                file_index = normalized.index(suggested_path)
                page_index = int(suggestion.get("page_index", 0))
        return cls(normalized, tuple(page_counts), file_index, page_index)

    @property
    def current_path(self):
        return self.pdf_paths[self.file_index]

    @property
    def current_page_count(self):
        return self.page_counts[self.file_index]

    @property
    def can_previous_file(self):
        return self.file_index > 0

    @property
    def can_next_file(self):
        return self.file_index < len(self.pdf_paths) - 1

    @property
    def can_previous_page(self):
        return self.page_index > 0

    @property
    def can_next_page(self):
        return self.page_index < self.current_page_count - 1

    def previous_file(self):
        if self.can_previous_file:
            self.file_index -= 1
            self.page_index = 0
        return self

    def next_file(self):
        if self.can_next_file:
            self.file_index += 1
            self.page_index = 0
        return self

    def previous_page(self):
        if self.can_previous_page:
            self.page_index -= 1
        return self

    def next_page(self):
        if self.can_next_page:
            self.page_index += 1
        return self
