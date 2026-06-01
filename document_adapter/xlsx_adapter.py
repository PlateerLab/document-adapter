"""XLSX 어댑터: openpyxl 기반. 각 워크시트를 하나의 표로 매핑한다.

- table_index = 워크시트 인덱스(0-based), location = 시트 이름
- 좌표 row/col 은 다른 어댑터와 동일하게 0-based 논리 좌표 (openpyxl 은 1-based)
- 병합 셀: top-left 가 anchor, 나머지는 non-anchor. openpyxl 은 병합 non-anchor
  셀이 읽기전용(MergedCell)이라 set_cell 은 anchor 로 redirect 한다.
- fill_form 은 base 구현이 get_tables/get_cell/set_cell 로 자동 동작한다.
"""
from __future__ import annotations

import re
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

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

# Excel 열 너비(문자 단위) → cm 근사. Calibri 11 기준 MDW≈7px, padding 5px.
# px → cm: px/96*2.54.
_PX_PER_CM = 96 / 2.54


def _colwidth_to_cm(chars: float | None) -> float | None:
    if chars is None:
        return None
    px = chars * 7 + 5
    return round(px / _PX_PER_CM, 1)


def _rowheight_to_cm(points: float | None) -> float | None:
    if points is None:
        return None
    return round(points / 72 * 2.54, 1)


class XlsxAdapter(DocumentAdapter):
    format = "xlsx"

    def _open(self) -> None:
        self._wb = load_workbook(str(self.path))

    def save(self, path: Path | str | None = None) -> Path:
        target = Path(path) if path else self.path
        self._wb.save(str(target))
        self.path = target
        return target

    # ---- helpers ----
    def _ws(self, table_index: int):
        sheets = self._wb.worksheets
        if table_index < 0 or table_index >= len(sheets):
            raise TableIndexError(f"XLSX sheet index {table_index} not found")
        return sheets[table_index]

    @staticmethod
    def _merge_map(ws) -> tuple[dict, dict]:
        """(anchor → span) 와 (covered cell → anchor) 매핑을 0-based 로 반환."""
        anchors: dict[tuple[int, int], tuple[int, int]] = {}
        covered: dict[tuple[int, int], tuple[int, int]] = {}
        for rng in ws.merged_cells.ranges:
            r0, c0 = rng.min_row - 1, rng.min_col - 1
            span = (rng.max_row - rng.min_row + 1, rng.max_col - rng.min_col + 1)
            anchors[(r0, c0)] = span
            for r in range(rng.min_row - 1, rng.max_row):
                for c in range(rng.min_col - 1, rng.max_col):
                    if (r, c) != (r0, c0):
                        covered[(r, c)] = (r0, c0)
        return anchors, covered

    @staticmethod
    def _dims(ws) -> tuple[int, int]:
        return ws.max_row or 0, ws.max_column or 0

    # ---- inspection ----
    def get_placeholders(self) -> list[str]:
        keys: set[str] = set()
        for ws in self._wb.worksheets:
            for row in ws.iter_rows():
                for cell in row:
                    if isinstance(cell.value, str):
                        keys.update(TAG_PATTERN.findall(cell.value))
        return sorted(keys)

    def get_tables(self, min_rows: int = 1, min_cols: int = 1,
                   preview_rows: int = 4, max_cell_len: int = 40) -> list[TableSchema]:
        schemas: list[TableSchema] = []
        for idx, ws in enumerate(self._wb.worksheets):
            rows, cols = self._dims(ws)
            if rows < min_rows or cols < min_cols:
                continue
            anchors, covered = self._merge_map(ws)
            visible = min(rows, preview_rows)
            preview: list[list[str | None]] = [
                [None] * cols for _ in range(visible)
            ]
            for r in range(visible):
                for c in range(cols):
                    if (r, c) in covered:
                        continue
                    v = ws.cell(row=r + 1, column=c + 1).value
                    preview[r][c] = ("" if v is None else str(v))[:max_cell_len]
            merges = [MergeInfo(anchor=a, span=s) for a, s in anchors.items()]
            col_widths = [
                _colwidth_to_cm(ws.column_dimensions[get_column_letter(c + 1)].width)
                for c in range(cols)
            ]
            row_heights = [
                _rowheight_to_cm(ws.row_dimensions[r + 1].height)
                for r in range(rows)
            ]
            schemas.append(TableSchema(
                index=idx, rows=rows, cols=cols, preview=preview,
                location=ws.title, merges=merges,
                column_widths_cm=col_widths if any(col_widths) else None,
                row_heights_cm=row_heights if any(row_heights) else None,
            ))
        return schemas

    def get_cell(self, table_index: int, row: int, col: int) -> CellContent:
        ws = self._ws(table_index)
        rows, cols = self._dims(ws)
        if row < 0 or col < 0 or row >= rows or col >= cols:
            raise CellOutOfBoundsError(
                f"cell ({row},{col}) out of bounds ({rows}x{cols})")
        anchors, covered = self._merge_map(ws)
        if (row, col) in covered:
            ar, ac = covered[(row, col)]
            is_anchor, anchor, span = False, (ar, ac), anchors[(ar, ac)]
            v = ws.cell(row=ar + 1, column=ac + 1).value
        else:
            is_anchor, anchor = True, (row, col)
            span = anchors.get((row, col), (1, 1))
            v = ws.cell(row=row + 1, column=col + 1).value
        text = "" if v is None else str(v)
        width_cm = _colwidth_to_cm(
            ws.column_dimensions[get_column_letter(anchor[1] + 1)].width)
        height_cm = _rowheight_to_cm(ws.row_dimensions[anchor[0] + 1].height)
        return CellContent(
            row=row, col=col, text=text, paragraphs=[text],
            is_anchor=is_anchor, anchor=anchor, span=span,
            width_cm=width_cm, height_cm=height_cm, char_count=len(text))

    # ---- editing ----
    def render_template(self, context: dict[str, Any], *,
                        on_missing: str = "blank") -> dict[str, list[str]]:
        report = self._render_report(self.get_placeholders(), context, on_missing)

        def repl(m: "re.Match[str]") -> str:
            key = m.group(1)
            if key in context:
                return str(context[key])
            return "" if on_missing == "blank" else m.group(0)

        for ws in self._wb.worksheets:
            for row in ws.iter_rows():
                for cell in row:
                    if isinstance(cell.value, str) and TAG_PATTERN.search(cell.value):
                        cell.value = TAG_PATTERN.sub(repl, cell.value)
        return report

    def _resolve_writable(self, ws, row: int, col: int,
                          allow_merge_redirect: bool) -> tuple[int, int]:
        """병합 non-anchor 좌표면 anchor 로 redirect (openpyxl MergedCell 은 읽기전용)."""
        _, covered = self._merge_map(ws)
        if (row, col) in covered:
            if not allow_merge_redirect:
                ar, ac = covered[(row, col)]
                raise MergedCellWriteError(
                    f"cell ({row},{col}) is part of a merge anchored at "
                    f"({ar},{ac}). Write to the anchor, or pass "
                    f"allow_merge_redirect=True.")
            return covered[(row, col)]
        return row, col

    def set_cell(self, table_index: int, row: int, col: int, value: str,
                 *, allow_merge_redirect: bool = False) -> str:
        ws = self._ws(table_index)
        rows, cols = self._dims(ws)
        if row < 0 or col < 0 or row >= rows or col >= cols:
            raise CellOutOfBoundsError(
                f"cell ({row},{col}) out of bounds ({rows}x{cols})")
        wr, wc = self._resolve_writable(ws, row, col, allow_merge_redirect)
        cell = ws.cell(row=wr + 1, column=wc + 1)
        old = "" if cell.value is None else str(cell.value)
        cell.value = value
        return old

    def append_to_cell(self, table_index: int, row: int, col: int, value: str,
                       separator: str = "  ", *,
                       allow_merge_redirect: bool = False) -> str:
        ws = self._ws(table_index)
        wr, wc = self._resolve_writable(ws, row, col, allow_merge_redirect)
        cell = ws.cell(row=wr + 1, column=wc + 1)
        old = "" if cell.value is None else str(cell.value)
        cell.value = f"{old}{separator}{value}" if old else value
        return old

    def append_row(self, table_index: int, values: list[str]) -> None:
        ws = self._ws(table_index)
        ws.append(list(values))
