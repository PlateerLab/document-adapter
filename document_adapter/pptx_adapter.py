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
    ShapeInfo,
    TableIndexError,
    TableSchema,
)

TAG_PATTERN = re.compile(r"\{\{\s*(\w+)\s*\}\}")
_A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"

# OOXML EMU (English Metric Unit) → cm: 1 cm = 360000 EMU
_EMU_PER_CM = 360000


def _emu_to_cm(emu: Any) -> float | None:
    """EMU 값을 cm 1자리 반올림. None 또는 0 은 None."""
    if emu is None:
        return None
    try:
        v = int(emu)
    except (TypeError, ValueError):
        return None
    if v <= 0:
        return None
    return round(v / _EMU_PER_CM, 1)


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

            # 셀 크기 힌트 (EMU → cm). LLM 이 오버플로 위험 셀을 판단하는 데 사용.
            col_widths = [
                _emu_to_cm(getattr(col, "width", None)) for col in table.columns
            ]
            row_heights = [
                _emu_to_cm(getattr(row, "height", None)) for row in table.rows
            ]
            col_widths_out = col_widths if any(v is not None for v in col_widths) else None
            row_heights_out = row_heights if any(v is not None for v in row_heights) else None

            schemas.append(
                TableSchema(
                    index=g_idx,
                    rows=n_rows,
                    cols=n_cols,
                    preview=preview,
                    location=f"slide {s_idx}",
                    merges=merges,
                    column_widths_cm=col_widths_out,
                    row_heights_cm=row_heights_out,
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

        # 셀 크기 힌트: anchor 위치부터 span 만큼의 column/row 합.
        a_r, a_c = anchor
        r_span, c_span = span
        try:
            cols_list = list(table.columns)
            rows_list = list(table.rows)
            width_emu = sum(
                getattr(cols_list[i], "width", 0) or 0
                for i in range(a_c, min(a_c + c_span, len(cols_list)))
            )
            height_emu = sum(
                getattr(rows_list[i], "height", 0) or 0
                for i in range(a_r, min(a_r + r_span, len(rows_list)))
            )
        except (IndexError, AttributeError):
            width_emu = 0
            height_emu = 0

        return CellContent(
            row=row,
            col=col,
            text=text,
            paragraphs=paragraphs_text,
            is_anchor=is_anchor,
            anchor=anchor,
            span=span,
            nested_table_indices=[],  # PPTX는 중첩 테이블 미지원
            width_cm=_emu_to_cm(width_emu),
            height_cm=_emu_to_cm(height_emu),
            char_count=len(text),
        )

    # ---- shape text (v0.8+) ----
    def get_shapes(
        self,
        slide_index: int | None = None,
        min_text_len: int = 1,
        max_preview: int = 40,
    ) -> list[ShapeInfo]:
        """표 외 shape (textbox / placeholder / 도형 텍스트) 수집.

        ``slide_index`` 는 1-based. None 이면 전체.
        ``min_text_len`` 미만의 텍스트는 제외 (0 을 주면 빈 shape 도 포함).
        """
        shapes_out: list[ShapeInfo] = []
        for s_idx, slide in enumerate(self._prs.slides, 1):
            if slide_index is not None and s_idx != slide_index:
                continue
            for shape in slide.shapes:
                if shape.has_table:
                    continue  # 표는 get_tables 로
                if not shape.has_text_frame:
                    continue
                text = (shape.text_frame.text or "").strip()
                if len(text) < min_text_len:
                    continue
                ph_type = None
                try:
                    ph = shape.placeholder_format
                    if ph is not None and ph.type is not None:
                        ph_type = str(ph.type).rsplit(".", 1)[-1]
                except ValueError:
                    pass
                kind = "placeholder" if ph_type else "text_box"
                shapes_out.append(
                    ShapeInfo(
                        slide_index=s_idx,
                        shape_id=shape.shape_id,
                        name=shape.name,
                        kind=kind,
                        has_text=bool(text),
                        text=text,
                        text_preview=text[:max_preview],
                        placeholder_type=ph_type,
                    )
                )
        return shapes_out

    def set_shape_text(
        self,
        slide_index: int,
        shape_id: int,
        text: str,
    ) -> str:
        """shape 의 텍스트를 text 로 교체. 기존 run-level 포맷 보존."""
        for s_idx, slide in enumerate(self._prs.slides, 1):
            if s_idx != slide_index:
                continue
            for shape in slide.shapes:
                if shape.shape_id != shape_id:
                    continue
                if not shape.has_text_frame:
                    raise ValueError(
                        f"shape {shape_id} on slide {slide_index} has no text frame"
                    )
                old = shape.text_frame.text or ""
                _set_text_frame_preserving_format(shape.text_frame, text)
                return old
        raise ValueError(
            f"shape not found: slide_index={slide_index}, shape_id={shape_id}"
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
        """표 끝에 새 행 추가 — 마지막 <a:tr> 을 deepcopy 후 텍스트만 비움.

        python-pptx 공식 add_row 는 없지만 OOXML 스키마상 <a:tr> 을 붙이는 것만으로
        행이 추가된다. 마지막 행의 셀 구조 (gridSpan/rowSpan/hMerge/vMerge, tcPr
        스타일) 를 그대로 상속해 이전 행과 동일한 서식의 빈 행이 생긴다.

        제약:
          - 마지막 행이 위 행의 rowSpan 영역에 속하면 (vMerge="1" 또는 rowSpan>1 셀
            존재) 복제 시 교차 병합이 오동작하므로 ``NotImplementedForFormat``.
        """
        table = self._get_table(table_index)
        tbl_elem = table._tbl  # lxml <a:tbl>

        a = f"{{{_A_NS}}}"
        trs = tbl_elem.findall(f"{a}tr")
        if not trs:
            raise NotImplementedForFormat("cannot append row to empty PPTX table")

        last_row = trs[-1]
        for tc in last_row.findall(f"{a}tc"):
            if tc.get("vMerge") == "1":
                raise NotImplementedForFormat(
                    "last row participates in a cross-row merge (vMerge); "
                    "append_row is not safe for this table."
                )
            try:
                rs = int(tc.get("rowSpan", "1"))
            except (TypeError, ValueError):
                rs = 1
            if rs > 1:
                raise NotImplementedForFormat(
                    "last row contains a rowSpan anchor that extends past the table; "
                    "append_row is not safe for this table."
                )

        new_row = deepcopy(last_row)
        # 기존 run/paragraph 구조는 유지하고 <a:t>.text 만 비움 (스타일 보존)
        for tc in new_row.findall(f"{a}tc"):
            txBody = tc.find(f"{a}txBody")
            if txBody is None:
                continue
            for p in txBody.findall(f"{a}p"):
                for r_el in p.findall(f"{a}r"):
                    for t_el in r_el.findall(f"{a}t"):
                        t_el.text = ""
        tbl_elem.append(new_row)

        new_row_idx = len(trs)  # 새 행 인덱스 (append 전 길이 = 새 행 position)
        n_cols = len(list(table.columns))
        for i, value in enumerate(values):
            if i >= n_cols:
                break
            try:
                self.set_cell(table_index, new_row_idx, i, value)
            except MergedCellWriteError:
                # 복제로 상속된 병합의 non-anchor 좌표는 자연히 스킵
                continue


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
