"""HWPX 어댑터: python-hwpx 기반.

버그 회피:
- set_cell_text()는 빈 셀에서 lxml/ElementTree 혼용 에러가 발생 (v2.9.0) →
  cell.paragraphs[0].text 직접 할당으로 우회
- replace_text_in_runs()는 한글 공백이 run으로 쪼개질 때 매칭 실패 →
  위치 기반 편집을 권장

부가:
- manifest fallback 로그가 기본적으로 매우 시끄러움 → logging 레벨 조정
"""
from __future__ import annotations

import logging
import re
from pathlib import Path
from typing import Any, Iterator

# 경고성 로그 억제 (manifest fallback 등)
logging.getLogger("hwpx").setLevel(logging.ERROR)

from hwpx.document import HwpxDocument

from .base import DocumentAdapter, TableSchema

TAG_PATTERN = re.compile(r"\{\{\s*(\w+)\s*\}\}")


class HwpxAdapter(DocumentAdapter):
    format = "hwpx"

    def _open(self) -> None:
        self._doc = HwpxDocument.open(self.path)

    def save(self, path: Path | str | None = None) -> Path:
        target = Path(path) if path else self.path
        self._doc.save_to_path(target)
        self.path = target
        return target

    def close(self) -> None:
        self._doc.close()

    # ---- helpers ----
    def _iter_tables(self) -> Iterator[tuple[int, Any]]:
        idx = 0
        for section in self._doc.sections:
            for para in section.paragraphs:
                for tbl in para.tables:
                    yield idx, tbl
                    idx += 1

    def _get_table(self, table_index: int):
        for idx, tbl in self._iter_tables():
            if idx == table_index:
                return tbl
        raise IndexError(f"HWPX table index {table_index} not found")

    @staticmethod
    def _cell_text(cell) -> str:
        return " ".join(p.text for p in cell.paragraphs).strip()

    # ---- inspection ----
    def get_placeholders(self) -> list[str]:
        text = self._doc.export_text()
        return sorted(set(TAG_PATTERN.findall(text)))

    def get_tables(self, min_rows: int = 1, min_cols: int = 1,
                   preview_rows: int = 4, max_cell_len: int = 40) -> list[TableSchema]:
        schemas: list[TableSchema] = []
        for idx, tbl in self._iter_tables():
            rows, cols = tbl.row_count, tbl.column_count
            if rows < min_rows or cols < min_cols:
                continue
            preview: list[list[str]] = []
            for r in range(min(rows, preview_rows)):
                row_cells = []
                for c in range(cols):
                    text = self._cell_text(tbl.cell(r, c))
                    row_cells.append(text[:max_cell_len])
                preview.append(row_cells)
            schemas.append(TableSchema(index=idx, rows=rows, cols=cols, preview=preview))
        return schemas

    # ---- editing ----
    def render_template(self, context: dict[str, Any]) -> None:
        """본문 + 표 셀의 {{key}}를 paragraph 단위로 치환."""
        # 본문
        for section in self._doc.sections:
            for para in section.paragraphs:
                text = para.text
                if TAG_PATTERN.search(text):
                    para.text = TAG_PATTERN.sub(
                        lambda m: str(context.get(m.group(1), m.group(0))), text
                    )
        # 표 셀
        for _, tbl in self._iter_tables():
            for r in range(tbl.row_count):
                for c in range(tbl.column_count):
                    cell = tbl.cell(r, c)
                    for para in cell.paragraphs:
                        text = para.text
                        if TAG_PATTERN.search(text):
                            para.text = TAG_PATTERN.sub(
                                lambda m: str(context.get(m.group(1), m.group(0))), text
                            )

    def set_cell(self, table_index: int, row: int, col: int, value: str) -> str:
        """set_cell_text 버그 우회: paragraph.text 직접 할당."""
        tbl = self._get_table(table_index)
        cell = tbl.cell(row, col)
        paragraphs = list(cell.paragraphs)
        old = self._cell_text(cell)
        if paragraphs:
            paragraphs[0].text = value
            for p in paragraphs[1:]:
                p.text = ""
        return old

    def append_row(self, table_index: int, values: list[str]) -> None:
        """python-hwpx에는 표준 add_row API가 없음.
        대안: 템플릿에 충분한 빈 행을 미리 만들고 set_cell로 채우는 전략."""
        raise NotImplementedError(
            "HWPX는 python-hwpx에 동적 행 추가 공식 API가 없음. "
            "템플릿에 여분 행을 두고 set_cell로 채우는 방식을 권장."
        )
