"""PPTX 어댑터: python-pptx + 자체 {{key}} 치환 엔진.

표 구조:
- python-pptx는 ``cell.is_merge_origin`` / ``cell.is_spanned`` /
  ``cell.span_height`` / ``cell.span_width`` 로 병합 정보를 직접 노출.
- PPTX는 중첩 테이블이 없음 (셀은 text_frame만 보유).
"""
from __future__ import annotations

import re
import warnings
from copy import deepcopy
from pathlib import Path
from typing import Any, Iterator

from lxml import etree
from pptx import Presentation

from .base import (
    CellContent,
    CellOutOfBoundsError,
    DocumentAdapter,
    MergeInfo,
    MergedCellWriteError,
    NotImplementedForFormat,
    TableIndexError,
    TableSchema,
)

TAG_PATTERN = re.compile(r"\{\{\s*(\w+)\s*\}\}")
_A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"


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

    @staticmethod
    def _dimensions(table) -> tuple[int, int]:
        n_rows = len(list(table.rows))
        n_cols = len(list(table.columns))
        return n_rows, n_cols

    def _resolve_anchor_cell(
        self, table, row: int, col: int, *, allow_merge_redirect: bool
    ) -> tuple[Any, tuple[int, int], tuple[int, int], bool]:
        """(cell, anchor, span, is_anchor) 반환.

        non-anchor 좌표이고 allow_merge_redirect=False면 MergedCellWriteError.
        True면 anchor cell로 리디렉트하고 경고.
        """
        n_rows, n_cols = self._dimensions(table)
        if row < 0 or col < 0 or row >= n_rows or col >= n_cols:
            raise CellOutOfBoundsError(
                f"cell ({row},{col}) out of bounds ({n_rows}x{n_cols})"
            )
        cell = table.cell(row, col)

        # merge anchor 좌표 계산
        is_anchor = bool(getattr(cell, "is_merge_origin", False)) or not bool(
            getattr(cell, "is_spanned", False)
        )
        if cell.is_merge_origin:
            span = (cell.span_height, cell.span_width)
            anchor = (row, col)
        elif cell.is_spanned:
            # anchor는 앞쪽 어딘가. 탐색으로 origin 찾기
            anchor = self._find_merge_origin(table, row, col)
            origin_cell = table.cell(*anchor)
            span = (origin_cell.span_height, origin_cell.span_width)
            is_anchor = False
            if not allow_merge_redirect:
                raise MergedCellWriteError(
                    f"cell ({row},{col}) is part of a merged region anchored at "
                    f"({anchor[0]},{anchor[1]}) span={span}. "
                    f"Write to the anchor coordinate, or pass "
                    f"allow_merge_redirect=True."
                )
            warnings.warn(
                f"write to ({row},{col}) redirected to merge anchor "
                f"({anchor[0]},{anchor[1]})",
                stacklevel=3,
            )
            cell = origin_cell
        else:
            span = (1, 1)
            anchor = (row, col)

        return cell, anchor, span, is_anchor

    @staticmethod
    def _find_merge_origin(table, row: int, col: int) -> tuple[int, int]:
        """is_spanned 셀로부터 병합 origin 좌표를 역추적.

        python-pptx가 origin 좌표 자체를 직접 노출하지 않아 앞쪽 row/col을 훑어
        이 (row,col)을 포함하는 origin을 찾는다. 테이블 크기가 작을 때는 충분.
        """
        for r in range(row, -1, -1):
            for c in range(col, -1, -1):
                candidate = table.cell(r, c)
                if not candidate.is_merge_origin:
                    continue
                if (r + candidate.span_height > row) and (
                    c + candidate.span_width > col
                ):
                    return (r, c)
        # fallback
        return (row, col)

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
            n_rows, n_cols = self._dimensions(table)
            if n_rows < min_rows or n_cols < min_cols:
                continue

            visible_rows = min(n_rows, preview_rows)
            preview: list[list[str | None]] = [
                [None for _ in range(n_cols)] for _ in range(visible_rows)
            ]
            merges: list[MergeInfo] = []

            for r in range(n_rows):
                for c in range(n_cols):
                    cell = table.cell(r, c)
                    if cell.is_spanned:
                        continue  # non-anchor, preview stays None
                    # anchor (merge origin or standalone cell)
                    if r < visible_rows:
                        text = (cell.text or "").strip()
                        preview[r][c] = text[:max_cell_len]
                    if cell.is_merge_origin:
                        span = (cell.span_height, cell.span_width)
                        if span != (1, 1):
                            merges.append(MergeInfo(anchor=(r, c), span=span))

            schemas.append(
                TableSchema(
                    index=g_idx,
                    rows=n_rows,
                    cols=n_cols,
                    preview=preview,
                    location=f"slide {s_idx}",
                    merges=merges,
                )
            )
        return schemas

    def get_cell(self, table_index: int, row: int, col: int) -> CellContent:
        table = self._get_table(table_index)
        n_rows, n_cols = self._dimensions(table)
        if row < 0 or col < 0 or row >= n_rows or col >= n_cols:
            raise CellOutOfBoundsError(
                f"cell ({row},{col}) out of bounds ({n_rows}x{n_cols})"
            )
        cell = table.cell(row, col)

        if cell.is_merge_origin:
            anchor = (row, col)
            span = (cell.span_height, cell.span_width)
            is_anchor = True
            source_cell = cell
        elif cell.is_spanned:
            anchor = self._find_merge_origin(table, row, col)
            source_cell = table.cell(*anchor)
            span = (source_cell.span_height, source_cell.span_width)
            is_anchor = False
        else:
            anchor = (row, col)
            span = (1, 1)
            is_anchor = True
            source_cell = cell

        paragraphs_text = [p.text for p in source_cell.text_frame.paragraphs]
        text = source_cell.text or ""

        return CellContent(
            row=row,
            col=col,
            text=text,
            paragraphs=paragraphs_text,
            is_anchor=is_anchor,
            anchor=anchor,
            span=span,
            nested_table_indices=[],  # PPTX는 중첩 테이블 미지원
        )

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
        raise TableIndexError(f"PPTX table index {table_index} not found")

    def set_cell(
        self,
        table_index: int,
        row: int,
        col: int,
        value: str,
        *,
        allow_merge_redirect: bool = False,
    ) -> str:
        table = self._get_table(table_index)
        cell, _, _, _ = self._resolve_anchor_cell(
            table, row, col, allow_merge_redirect=allow_merge_redirect
        )
        old = cell.text
        _set_text_frame_preserving_format(cell.text_frame, value)
        return old

    def append_to_cell(
        self,
        table_index: int,
        row: int,
        col: int,
        value: str,
        separator: str = "  ",
        *,
        allow_merge_redirect: bool = False,
    ) -> str:
        table = self._get_table(table_index)
        cell, _, _, _ = self._resolve_anchor_cell(
            table, row, col, allow_merge_redirect=allow_merge_redirect
        )
        old = cell.text
        new_value = f"{old}{separator}{value}" if old else value
        _set_text_frame_preserving_format(cell.text_frame, new_value)
        return old

    def append_row(self, table_index: int, values: list[str]) -> None:
        """python-pptx는 표 행 추가 API를 제공하지 않는다.
        LLM에게는 '지원 안 함'으로 알리는 게 정직한 방식."""
        raise NotImplementedForFormat(
            "PPTX는 python-pptx에 동적 행 추가 API가 없음. "
            "템플릿 단계에서 충분한 빈 행을 만들어 두고 set_cell로 채우는 방식을 권장."
        )


def _set_text_frame_preserving_format(text_frame, value: str) -> None:
    """Write ``value`` into ``text_frame`` without losing run-level formatting.

    ``python-pptx`` exposes ``cell.text = value`` (which proxies to the text
    frame) but the setter deletes every run and replaces them with a single
    default-styled run. This destroys two kinds of formatting:

    1. **Runs that already exist** — font family, size, bold, color, etc.
    2. **Empty paragraphs that hold an ``<a:endParaRPr>``**, which is where
       PowerPoint stores the "what would the next character look like" run
       properties for an otherwise empty cell. Real-world templates put font
       information here so that the cell looks right even before any text
       is typed.

    Strategy:

    - If the paragraph already has runs, reuse the first one and blank the
      rest (simple case that covers pre-filled cells).
    - Otherwise, build a new ``<a:r>`` manually and clone ``<a:endParaRPr>``
      into its ``<a:rPr>`` so the empty-cell font survives.

    Paragraph comparison uses index, not identity, because python-pptx
    returns a fresh Python wrapper on every ``paragraphs`` access, which
    would cause a naive ``para is first_para`` check to always be False
    and blank the run we just populated.
    """
    paragraphs = list(text_frame.paragraphs)
    first_idx = next((i for i, para in enumerate(paragraphs) if para.runs), None)

    if first_idx is not None:
        first_para = paragraphs[first_idx]
        first_para.runs[0].text = value
        for run in first_para.runs[1:]:
            run.text = ""
        for i, para in enumerate(paragraphs):
            if i == first_idx:
                continue
            for run in para.runs:
                run.text = ""
        return

    target_para = paragraphs[0] if paragraphs else None
    if target_para is None:
        text_frame.text = value
        return

    p_el = target_para._p
    end_rpr = p_el.find(f"{{{_A_NS}}}endParaRPr")

    r_el = etree.SubElement(p_el, f"{{{_A_NS}}}r")
    if end_rpr is not None:
        rpr = deepcopy(end_rpr)
        rpr.tag = f"{{{_A_NS}}}rPr"
        r_el.insert(0, rpr)
    t_el = etree.SubElement(r_el, f"{{{_A_NS}}}t")
    t_el.text = value

    for i, para in enumerate(paragraphs):
        if i == 0:
            continue
        for run in para.runs:
            run.text = ""
