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


def _build_grid(table) -> tuple[dict[tuple[int, int], dict], int, int]:
    """(row,col) → {anchor, span, is_anchor, tc, cell} 매핑.

    python-docx는 병합된 셀에 대해 동일 ``_tc``를 여러 logical 위치에서 반환.
    이 특성을 이용해 anchor/span을 역산한다.
    """
    n_rows = len(table.rows)
    if n_rows == 0:
        return {}, 0, 0
    # logical column 수: 각 행의 cells 길이 중 최대값
    n_cols = max((len(row.cells) for row in table.rows), default=0)

    tc_to_positions: dict[int, tuple[Any, list[tuple[int, int]]]] = {}
    for r, row in enumerate(table.rows):
        row_cells = row.cells
        for c in range(n_cols):
            if c >= len(row_cells):
                continue
            cell = row_cells[c]
            key = id(cell._tc)
            if key not in tc_to_positions:
                tc_to_positions[key] = (cell, [])
            tc_to_positions[key][1].append((r, c))

    grid: dict[tuple[int, int], dict] = {}
    for cell, positions in tc_to_positions.values():
        if len(positions) == 1:
            (r, c) = positions[0]
            grid[(r, c)] = {
                "anchor": (r, c),
                "span": (1, 1),
                "is_anchor": True,
                "cell": cell,
            }
        else:
            anchor = min(positions)
            max_r = max(p[0] for p in positions)
            max_c = max(p[1] for p in positions)
            span = (max_r - anchor[0] + 1, max_c - anchor[1] + 1)
            for pos in positions:
                grid[pos] = {
                    "anchor": anchor,
                    "span": span,
                    "is_anchor": pos == anchor,
                    "cell": cell,
                }
    return grid, n_rows, n_cols


class DocxAdapter(DocumentAdapter):
    format = "docx"

    def _open(self) -> None:
        self._doc = Document(self.path)

    def save(self, path: Path | str | None = None) -> Path:
        target = Path(path) if path else self.path
        self._doc.save(target)
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

            schemas.append(
                TableSchema(
                    index=idx,
                    rows=n_rows,
                    cols=n_cols,
                    preview=preview,
                    merges=merges,
                    parent_path=parent_path or None,
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

        return CellContent(
            row=row,
            col=col,
            text=text,
            paragraphs=paragraphs_text,
            is_anchor=info["is_anchor"],
            anchor=info["anchor"],
            span=info["span"],
            nested_table_indices=nested_indices,
        )

    # ---- editing ----
    def render_template(self, context: dict[str, Any]) -> None:
        """docxtpl 기반 Jinja2 렌더. 참고:
        - `{%tr for row in rows %}` / `{%tr endfor %}`는 **각각 별도 행**에 두어야 함
        - 같은 행에 두면 `<w:tr>` 전체가 `{% for %}`로 교체되어 endfor 손실
        """
        tpl = DocxTemplate(self.path)
        tpl.render(context)
        tpl.save(self.path)
        self._doc = Document(self.path)

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
