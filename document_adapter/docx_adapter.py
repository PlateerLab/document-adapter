"""DOCX 어댑터: python-docx (편집) + docxtpl (템플릿 렌더).

표 구조:
- python-docx의 ``row.cells[col]``은 병합된 셀에 대해 동일한 ``_tc``를 여러 번 반환한다.
  이 성질을 이용해 (row, col) → ``_tc`` 매핑을 만든 뒤, 동일 ``_tc``가 등장한
  position들의 bounding box로 병합 anchor/span을 계산한다.
- 중첩 테이블은 ``cell.tables``를 DFS로 훑어 flat index를 부여.
"""
from __future__ import annotations

import re
import warnings
from copy import deepcopy
from pathlib import Path
from typing import Any, Iterator

from docx import Document
from docxtpl import DocxTemplate

from .base import (
    CellContent,
    CellOutOfBoundsError,
    DocumentAdapter,
    MergeInfo,
    MergedCellWriteError,
    TableIndexError,
    TableSchema,
)

TAG_PATTERN = re.compile(r"\{\{\s*(\w+)\s*\}\}")
_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

# OOXML EMU → cm (1 cm = 360000 EMU)
_EMU_PER_CM = 360000


def _emu_to_cm(emu: Any) -> float | None:
    if emu is None:
        return None
    try:
        v = int(emu)
    except (TypeError, ValueError):
        return None
    if v <= 0:
        return None
    return round(v / _EMU_PER_CM, 1)


def _build_grid(table) -> tuple[dict[tuple[int, int], dict], int, int]:
    """(row,col) → {anchor, span, is_anchor, cell} 매핑.

    python-docx 의 ``row.cells`` 는 가로(gridSpan)+세로(vMerge) 병합이 섞인 표에서
    내부적으로 깨진다(``tc_at_grid_offset`` ValueError). 그래서 OOXML 레이어
    (tblGrid/tr/tc + gridSpan + vMerge)에서 그리드를 직접 계산한다 — hwpx_core
    의 grid 처리와 동일한 접근. 각 행의 tc gridSpan 합이 곧 전체 열 수다.
    """
    from docx.oxml.ns import qn
    from docx.table import _Cell

    tbl = table._tbl
    grid_el = tbl.tblGrid
    n_cols = len(grid_el.gridCol_lst) if grid_el is not None else 0
    trs = tbl.tr_lst
    n_rows = len(trs)
    if n_rows == 0 or n_cols == 0:
        return {}, n_rows, n_cols

    # (r,c) → (anchor_info, is_anchor). anchor_info 는 병합 셀들이 공유하는
    # dict 로, rowspan 이 자라면 span 을 in-place 로 갱신한다.
    cells_info: dict[tuple[int, int], tuple[dict, bool]] = {}
    vmerge_active: dict[int, dict] = {}   # 시작열 → 진행 중인 세로병합 anchor_info

    for r, tr in enumerate(trs):
        col = 0
        for tc in tr.tc_lst:
            gs = tc.grid_span or 1
            tc_pr = tc.tcPr
            vmerge = tc_pr.find(qn("w:vMerge")) if tc_pr is not None else None
            vval = vmerge.get(qn("w:val")) if vmerge is not None else None
            is_continue = vmerge is not None and vval != "restart"

            if is_continue and col in vmerge_active:
                ai = vmerge_active[col]
                ai["span"] = (r - ai["anchor"][0] + 1, ai["span"][1])
                for cc in range(col, col + gs):
                    cells_info[(r, cc)] = (ai, False)
            else:
                ai = {"anchor": (r, col), "span": (1, gs), "cell": _Cell(tc, table)}
                for cc in range(col, col + gs):
                    cells_info[(r, cc)] = (ai, cc == col)
                if vval == "restart":
                    vmerge_active[col] = ai
                else:
                    vmerge_active.pop(col, None)
            col += gs

    grid: dict[tuple[int, int], dict] = {}
    for (r, c), (ai, is_anchor) in cells_info.items():
        grid[(r, c)] = {
            "anchor": ai["anchor"],
            "span": ai["span"],
            "is_anchor": is_anchor,
            "cell": ai["cell"],
        }
    return grid, n_rows, n_cols


class DocxAdapter(DocumentAdapter):
    format = "docx"

    def _open(self) -> None:
        self._doc = Document(str(self.path))

    def save(self, path: Path | str | None = None) -> Path:
        target = Path(path) if path else self.path
        self._doc.save(str(target))
        self.path = target
        return target

    # ---- helpers ----
    def _iter_tables(self) -> Iterator[tuple[int, Any, str]]:
        """Flat DFS (outer + nested). 각 yield: (flat_index, table, parent_path)."""
        idx_counter = [0]

        def walk(tbl, parent_path: str) -> Iterator[tuple[int, Any, str]]:
            current_idx = idx_counter[0]
            idx_counter[0] += 1
            yield current_idx, tbl, parent_path

            grid, _, _ = _build_grid(tbl)
            seen_tc: set[int] = set()
            for (r, c), info in grid.items():
                if not info["is_anchor"]:
                    continue
                tc_key = id(info["cell"]._tc)
                if tc_key in seen_tc:
                    continue
                seen_tc.add(tc_key)
                for nested in info["cell"].tables:
                    child_parent = (
                        f"{parent_path}.tables[{current_idx}].cell({r},{c})"
                    )
                    yield from walk(nested, child_parent)

        for tbl in self._doc.tables:
            yield from walk(tbl, "")

    def _get_table(self, table_index: int):
        for idx, tbl, _ in self._iter_tables():
            if idx == table_index:
                return tbl
        raise TableIndexError(f"DOCX table index {table_index} not found")

    def _resolve_anchor_cell(
        self, tbl, row: int, col: int, *, allow_merge_redirect: bool
    ) -> tuple[Any, dict]:
        """(row,col) → (cell, grid_info). non-anchor 정책 처리."""
        grid, n_rows, n_cols = _build_grid(tbl)
        if row < 0 or col < 0 or row >= n_rows or col >= n_cols:
            raise CellOutOfBoundsError(
                f"cell ({row},{col}) out of bounds ({n_rows}x{n_cols})"
            )
        info = grid.get((row, col))
        if info is None:
            raise CellOutOfBoundsError(
                f"cell ({row},{col}) does not resolve to any physical cell"
            )
        if not info["is_anchor"]:
            anchor_r, anchor_c = info["anchor"]
            if not allow_merge_redirect:
                raise MergedCellWriteError(
                    f"cell ({row},{col}) is part of a merged region anchored at "
                    f"({anchor_r},{anchor_c}) span={info['span']}. "
                    f"Write to the anchor coordinate, or pass "
                    f"allow_merge_redirect=True."
                )
            warnings.warn(
                f"write to ({row},{col}) redirected to merge anchor "
                f"({anchor_r},{anchor_c})",
                stacklevel=3,
            )
        return info["cell"], info

    # ---- inspection ----
    def get_placeholders(self) -> list[str]:
        keys: set[str] = set()
        for p in self._doc.paragraphs:
            keys.update(TAG_PATTERN.findall(p.text))
        # 모든 (중첩 포함) 표 셀에서 수집
        for _, tbl, _ in self._iter_tables():
            for row in tbl.rows:
                for cell in row.cells:
                    keys.update(TAG_PATTERN.findall(cell.text))
        return sorted(keys)

    def get_tables(self, min_rows: int = 1, min_cols: int = 1,
                   preview_rows: int = 4, max_cell_len: int = 40) -> list[TableSchema]:
        schemas: list[TableSchema] = []
        for idx, tbl, parent_path in self._iter_tables():
            grid, n_rows, n_cols = _build_grid(tbl)
            if n_rows < min_rows or n_cols < min_cols:
                continue

            visible_rows = min(n_rows, preview_rows)
            preview: list[list[str | None]] = [
                [None for _ in range(n_cols)] for _ in range(visible_rows)
            ]
            merges: list[MergeInfo] = []
            seen_anchors: set[tuple[int, int]] = set()

            # grid 순회 — 앵커 위치에만 텍스트 주입
            for (r, c), info in sorted(grid.items()):
                if info["anchor"] in seen_anchors:
                    continue
                if info["is_anchor"]:
                    seen_anchors.add(info["anchor"])
                    if r < visible_rows:
                        text = (info["cell"].text or "").strip()
                        preview[r][c] = text[:max_cell_len]
                    if info["span"] != (1, 1):
                        merges.append(MergeInfo(anchor=info["anchor"], span=info["span"]))

            # 셀 크기 힌트
            col_widths = [
                _emu_to_cm(getattr(col, "width", None)) for col in tbl.columns
            ]
            row_heights = [
                _emu_to_cm(getattr(row, "height", None)) for row in tbl.rows
            ]
            col_widths_out = col_widths if any(v is not None for v in col_widths) else None
            row_heights_out = row_heights if any(v is not None for v in row_heights) else None

            schemas.append(
                TableSchema(
                    index=idx,
                    rows=n_rows,
                    cols=n_cols,
                    preview=preview,
                    merges=merges,
                    parent_path=parent_path or None,
                    column_widths_cm=col_widths_out,
                    row_heights_cm=row_heights_out,
                )
            )
        return schemas

    def get_cell(self, table_index: int, row: int, col: int) -> CellContent:
        tbl = self._get_table(table_index)
        grid, n_rows, n_cols = _build_grid(tbl)
        if row < 0 or col < 0 or row >= n_rows or col >= n_cols:
            raise CellOutOfBoundsError(
                f"cell ({row},{col}) out of bounds ({n_rows}x{n_cols})"
            )
        info = grid.get((row, col))
        if info is None:
            raise CellOutOfBoundsError(
                f"cell ({row},{col}) does not resolve to any physical cell"
            )

        cell = info["cell"]
        paragraphs_text = [p.text for p in cell.paragraphs]
        text = cell.text or ""

        nested_indices: list[int] = []
        if info["is_anchor"] and list(cell.tables):
            nested_tc_ids = {id(t._tbl) for t in cell.tables}
            for child_idx, child_tbl, _ in self._iter_tables():
                if id(child_tbl._tbl) in nested_tc_ids:
                    nested_indices.append(child_idx)

        # 셀 크기 힌트 (anchor 기준 span 영역 합)
        a_r, a_c = info["anchor"]
        r_span, c_span = info["span"]
        try:
            cols_list = list(tbl.columns)
            rows_list = list(tbl.rows)
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
            is_anchor=info["is_anchor"],
            anchor=info["anchor"],
            span=info["span"],
            nested_table_indices=nested_indices,
            width_cm=_emu_to_cm(width_emu),
            height_cm=_emu_to_cm(height_emu),
            char_count=len(text),
        )

    # ---- editing ----
    def render_template(self, context: dict[str, Any], *,
                        on_missing: str = "blank") -> dict[str, list[str]]:
        """docxtpl 기반 Jinja2 렌더. 누락 키는 on_missing 정책(base 참조).
        - `{%tr for row in rows %}` / `{%tr endfor %}`는 **각각 별도 행**에 두어야 함
        - 같은 행에 두면 `<w:tr>` 전체가 `{% for %}`로 교체되어 endfor 손실
        """
        report = self._render_report(self.get_placeholders(), context, on_missing)
        tpl = DocxTemplate(self.path)
        if on_missing == "leave":
            import jinja2
            env = jinja2.Environment(undefined=jinja2.DebugUndefined,
                                     autoescape=True)
            tpl.render(context, jinja_env=env)
        else:
            # blank: docxtpl 기본(Jinja Undefined→""). error: 위에서 이미 raise.
            tpl.render(context)
        tpl.save(self.path)
        self._doc = Document(str(self.path))
        return report

    def set_cell(
        self,
        table_index: int,
        row: int,
        col: int,
        value: str,
        *,
        allow_merge_redirect: bool = False,
    ) -> str:
        tbl = self._get_table(table_index)
        cell, info = self._resolve_anchor_cell(
            tbl, row, col, allow_merge_redirect=allow_merge_redirect
        )
        old = cell.text
        _set_cell_preserving_format(cell, value)
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
        tbl = self._get_table(table_index)
        cell, info = self._resolve_anchor_cell(
            tbl, row, col, allow_merge_redirect=allow_merge_redirect
        )
        old = cell.text
        new_value = f"{old}{separator}{value}" if old else value
        _set_cell_preserving_format(cell, new_value)
        return old

    def append_row(self, table_index: int, values: list[str]) -> None:
        tbl = self._get_table(table_index)
        new_row = tbl.add_row()
        for i, v in enumerate(values):
            if i < len(new_row.cells):
                _set_cell_preserving_format(new_row.cells[i], v)


def _set_cell_preserving_format(cell, value: str) -> None:
    """Write ``value`` into ``cell`` without dropping run formatting.

    ``python-docx``'s ``cell.text = value`` setter wipes every paragraph and
    run in the cell, replacing them with a brand-new default-styled run. That
    destroys two kinds of formatting:

    1. **Existing runs** — font, size, bold, color on already-populated cells.
    2. **Paragraph mark run properties** — an empty cell often holds a
       ``<w:p><w:pPr><w:rPr>…</w:rPr></w:pPr></w:p>`` describing how the
       next typed character should look. Real templates put the table font
       here so the cell renders correctly even before any text exists.

    Strategy:

    - If any paragraph already has runs, reuse the first one and blank the
      rest.
    - Otherwise, append a new ``<w:r>`` into the first paragraph and clone
      its ``<w:pPr><w:rPr>`` into the new run's ``<w:rPr>`` so the empty-cell
      font survives.

    Paragraph identity is compared by index because python-docx returns a
    fresh Python wrapper on repeated ``cell.paragraphs`` accesses.
    """
    paragraphs = list(cell.paragraphs)
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
        cell.text = value
        return

    run = target_para.add_run(value)
    p_el = target_para._p
    ppr = p_el.find(f"{{{_W_NS}}}pPr")
    if ppr is not None:
        rpr_in_ppr = ppr.find(f"{{{_W_NS}}}rPr")
        if rpr_in_ppr is not None:
            cloned = deepcopy(rpr_in_ppr)
            cloned.tag = f"{{{_W_NS}}}rPr"
            run._r.insert(0, cloned)
