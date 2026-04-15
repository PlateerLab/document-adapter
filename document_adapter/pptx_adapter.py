"""PPTX 어댑터: python-pptx + 자체 {{key}} 치환 엔진."""
from __future__ import annotations

import re
from pathlib import Path
from typing import Any, Iterator

from pptx import Presentation
from pptx.slide import Slide

from .base import DocumentAdapter, TableSchema

TAG_PATTERN = re.compile(r"\{\{\s*(\w+)\s*\}\}")


class PptxAdapter(DocumentAdapter):
    format = "pptx"

    def _open(self) -> None:
        self._prs = Presentation(self.path)

    def save(self, path: Path | str | None = None) -> Path:
        target = Path(path) if path else self.path
        self._prs.save(target)
        self.path = target
        return target

    # ---- helpers ----
    def _iter_tables(self) -> Iterator[tuple[int, int, Any]]:
        """(global_index, slide_number_1based, table) 순회."""
        g_idx = 0
        for s_idx, slide in enumerate(self._prs.slides, 1):
            for shape in slide.shapes:
                if shape.has_table:
                    yield g_idx, s_idx, shape.table
                    g_idx += 1

    def _iter_text_frames(self) -> Iterator[Any]:
        for slide in self._prs.slides:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    yield shape.text_frame
                if shape.has_table:
                    for row in shape.table.rows:
                        for cell in row.cells:
                            yield cell.text_frame

    # ---- inspection ----
    def get_placeholders(self) -> list[str]:
        keys: set[str] = set()
        for tf in self._iter_text_frames():
            keys.update(TAG_PATTERN.findall(tf.text))
        return sorted(keys)

    def get_tables(self, min_rows: int = 1, min_cols: int = 1,
                   preview_rows: int = 4, max_cell_len: int = 40) -> list[TableSchema]:
        schemas: list[TableSchema] = []
        for g_idx, s_idx, table in self._iter_tables():
            rows = list(table.rows)
            cols = list(table.columns)
            if len(rows) < min_rows or len(cols) < min_cols:
                continue
            preview: list[list[str]] = []
            for row in rows[:preview_rows]:
                preview.append([c.text.strip()[:max_cell_len] for c in row.cells])
            schemas.append(TableSchema(
                index=g_idx, rows=len(rows), cols=len(cols),
                preview=preview, location=f"slide {s_idx}",
            ))
        return schemas

    # ---- editing ----
    def render_template(self, context: dict[str, Any]) -> None:
        """paragraph 단위로 {{key}}를 치환. run이 쪼개진 경우를 처리하기 위해
        paragraph 전체 텍스트를 재조립 후 첫 run에 담는다 (서식 일부 손실 가능)."""
        for tf in self._iter_text_frames():
            for para in tf.paragraphs:
                full_text = "".join(run.text for run in para.runs)
                if not TAG_PATTERN.search(full_text):
                    continue
                rendered = TAG_PATTERN.sub(
                    lambda m: str(context.get(m.group(1), m.group(0))),
                    full_text,
                )
                if para.runs:
                    para.runs[0].text = rendered
                    for run in para.runs[1:]:
                        run.text = ""

    def _get_table(self, table_index: int):
        for g_idx, _, table in self._iter_tables():
            if g_idx == table_index:
                return table
        raise IndexError(f"PPTX table index {table_index} not found")

    def set_cell(self, table_index: int, row: int, col: int, value: str) -> str:
        table = self._get_table(table_index)
        cell = table.cell(row, col)
        old = cell.text
        _set_text_frame_preserving_format(cell.text_frame, value)
        return old

    def append_row(self, table_index: int, values: list[str]) -> None:
        """python-pptx는 표 행 추가 API를 제공하지 않는다.
        LLM에게는 '지원 안 함'으로 알리는 게 정직한 방식."""
        raise NotImplementedError(
            "PPTX는 python-pptx에 동적 행 추가 API가 없음. "
            "템플릿 단계에서 충분한 빈 행을 만들어 두고 set_cell로 채우는 방식을 권장."
        )


def _set_text_frame_preserving_format(text_frame, value: str) -> None:
    """Write ``value`` into the first run of ``text_frame`` without losing font/size.

    ``python-pptx`` exposes ``cell.text = value`` (which proxies to the text
    frame) but the setter deletes every run and replaces them with a single
    default-styled run. To keep template formatting intact we reuse the first
    existing run instead and blank all other runs/paragraphs.

    An empty text frame (no runs anywhere) falls back to the default setter —
    that is the one situation where there is no formatting to preserve.

    We compare paragraphs by index rather than identity because python-pptx
    returns a fresh Python wrapper on every ``paragraphs`` access, so
    ``para is first_para`` is always False and would cause the second loop
    to blank the run we just populated.
    """
    paragraphs = list(text_frame.paragraphs)
    first_idx = next((i for i, para in enumerate(paragraphs) if para.runs), None)

    if first_idx is None:
        text_frame.text = value
        return

    first_para = paragraphs[first_idx]
    first_para.runs[0].text = value
    for run in first_para.runs[1:]:
        run.text = ""

    for i, para in enumerate(paragraphs):
        if i == first_idx:
            continue
        for run in para.runs:
            run.text = ""
