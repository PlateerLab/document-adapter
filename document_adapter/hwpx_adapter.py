"""HWPX 어댑터: document_adapter.hwpx_core 기반 (python-hwpx 의존 없음).

- 패키지 로드/저장은 HwpxPackage가 처리 (bytes-copy 보존, 수정 XML만 재직렬화)
- 표 순회는 iter_grid 직접 사용 (cellAddr + cellSpan → logical grid)
- run-level 포맷은 paragraph 헬퍼가 첫 <hp:t>만 갈아끼워 유지
"""
from __future__ import annotations

import re
import warnings
from copy import deepcopy
from pathlib import Path
from typing import Any, Iterator

from lxml import etree

from document_adapter.hwpx_core import (
    HP_CELL_ADDR,
    HP_P,
    HP_RUN,
    HP_SUBLIST,
    HP_T,
    HP_TBL,
    HP_TC,
    HP_TR,
    HwpxPackage,
    cell_paragraph_texts,
    cell_paragraphs,
    cell_text,
    iter_grid,
    nested_tables,
    paragraph_text,
    set_paragraph_text,
    table_shape,
    write_cell,
)

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


class HwpxAdapter(DocumentAdapter):
    format = "hwpx"

    def _open(self) -> None:
        self._pkg = HwpxPackage.open(self.path)

    def save(self, path: Path | str | None = None) -> Path:
        target = Path(path) if path else self.path
        self._pkg.save(target)
        self.path = target
        return target

    def close(self) -> None:
        self._pkg.close()

    # ---- 테이블 순회 ----

    def _iter_tables(
        self,
    ) -> Iterator[tuple[int, etree._Element, str, str]]:
        """(flat_index, tbl_element, parent_path, section_part_name) 순회.

        최상위 테이블과 그 안의 중첩 테이블을 DFS 순서로 부여.
        """
        idx_counter = [0]

        def walk(tbl: etree._Element, parent_path: str, section_name: str):
            current_idx = idx_counter[0]
            idx_counter[0] += 1
            yield current_idx, tbl, parent_path, section_name
            seen_anchors: set[tuple[int, int]] = set()
            for entry in iter_grid(tbl):
                if not entry.is_anchor or entry.anchor in seen_anchors:
                    continue
                seen_anchors.add(entry.anchor)
                for child_tbl in nested_tables(entry.cell_element):
                    child_parent = (
                        f"{parent_path}.tables[{current_idx}].cell"
                        f"({entry.anchor[0]},{entry.anchor[1]})"
                    )
                    yield from walk(child_tbl, child_parent, section_name)

        for section_name, root in self._pkg.iter_section_roots():
            # 최상위 <hp:tbl> 찾기: root > hp:p > hp:run > hp:tbl
            for p in root.findall(HP_P):
                for run in p.findall(HP_RUN):
                    for tbl in run.findall(HP_TBL):
                        yield from walk(tbl, "", section_name)

    def _get_table(self, table_index: int) -> tuple[etree._Element, str]:
        """flat_index로 (tbl_element, section_part_name) 반환."""
        for idx, tbl, _, section_name in self._iter_tables():
            if idx == table_index:
                return tbl, section_name
        raise TableIndexError(f"HWPX table index {table_index} not found")

    def _find_grid_entry(self, tbl: etree._Element, row: int, col: int):
        rows, cols = table_shape(tbl)
        if row < 0 or col < 0 or row >= rows or col >= cols:
            raise CellOutOfBoundsError(
                f"cell ({row},{col}) out of bounds ({rows}x{cols})"
            )
        for entry in iter_grid(tbl):
            if (entry.row, entry.column) == (row, col):
                return entry
        raise CellOutOfBoundsError(
            f"cell ({row},{col}) does not resolve to any physical cell"
        )

    def _resolve_anchor_cell(
        self,
        tbl: etree._Element,
        row: int,
        col: int,
        *,
        allow_merge_redirect: bool,
    ):
        entry = self._find_grid_entry(tbl, row, col)
        if not entry.is_anchor:
            anchor_r, anchor_c = entry.anchor
            if not allow_merge_redirect:
                raise MergedCellWriteError(
                    f"cell ({row},{col}) is part of a merged region anchored at "
                    f"({anchor_r},{anchor_c}) span={entry.span}. "
                    f"Write to the anchor coordinate, or pass "
                    f"allow_merge_redirect=True."
                )
            warnings.warn(
                f"write to ({row},{col}) redirected to merge anchor "
                f"({anchor_r},{anchor_c})",
                stacklevel=3,
            )
        return entry

    # ---- 검사 ----

    def get_placeholders(self) -> list[str]:
        text = self._pkg.export_text()
        return sorted(set(TAG_PATTERN.findall(text)))

    def get_tables(
        self,
        min_rows: int = 1,
        min_cols: int = 1,
        preview_rows: int = 4,
        max_cell_len: int = 40,
    ) -> list[TableSchema]:
        schemas: list[TableSchema] = []
        for idx, tbl, parent_path, _ in self._iter_tables():
            rows, cols = table_shape(tbl)
            if rows < min_rows or cols < min_cols:
                continue

            visible_rows = min(rows, preview_rows)
            preview: list[list[str | None]] = [
                [None for _ in range(cols)] for _ in range(visible_rows)
            ]
            merges: list[MergeInfo] = []
            seen_anchors: set[tuple[int, int]] = set()

            for entry in iter_grid(tbl):
                if entry.anchor in seen_anchors:
                    continue
                if entry.is_anchor:
                    seen_anchors.add(entry.anchor)
                    if entry.row < visible_rows:
                        text = cell_text(entry.cell_element).strip()
                        preview[entry.row][entry.column] = text[:max_cell_len]
                    if entry.span != (1, 1):
                        merges.append(MergeInfo(anchor=entry.anchor, span=entry.span))

            schemas.append(
                TableSchema(
                    index=idx,
                    rows=rows,
                    cols=cols,
                    preview=preview,
                    merges=merges,
                    parent_path=parent_path or None,
                )
            )
        return schemas

    def get_cell(self, table_index: int, row: int, col: int) -> CellContent:
        tbl, _ = self._get_table(table_index)
        entry = self._find_grid_entry(tbl, row, col)

        tc = entry.cell_element
        text = cell_text(tc)
        paragraphs = cell_paragraph_texts(tc)

        nested_indices: list[int] = []
        if entry.is_anchor:
            child_tbls = nested_tables(tc)
            if child_tbls:
                nested_ids = {id(t) for t in child_tbls}
                for child_idx, child_tbl, _, _ in self._iter_tables():
                    if id(child_tbl) in nested_ids:
                        nested_indices.append(child_idx)

        return CellContent(
            row=row,
            col=col,
            text=text,
            paragraphs=paragraphs,
            is_anchor=entry.is_anchor,
            anchor=entry.anchor,
            span=entry.span,
            nested_table_indices=nested_indices,
        )

    # ---- 편집 ----

    def render_template(self, context: dict[str, Any]) -> None:
        """섹션의 모든 <hp:p> 에서 {{key}} 치환. paragraph 단위로 처리해
        run 포맷은 보존한다 (첫 <hp:t>에 치환 결과를 쓰고 나머지는 비움).
        """
        def substitute(p: etree._Element) -> bool:
            text = paragraph_text(p)
            if not TAG_PATTERN.search(text):
                return False
            new_text = TAG_PATTERN.sub(
                lambda m: str(context.get(m.group(1), m.group(0))), text
            )
            set_paragraph_text(p, new_text)
            return True

        for section_name, root in self._pkg.iter_section_roots():
            changed = False
            for p in root.iter(HP_P):
                if substitute(p):
                    changed = True
            if changed:
                self._pkg.mark_dirty(section_name)

    def set_cell(
        self,
        table_index: int,
        row: int,
        col: int,
        value: str,
        *,
        allow_merge_redirect: bool = False,
    ) -> str:
        tbl, section_name = self._get_table(table_index)
        entry = self._resolve_anchor_cell(
            tbl, row, col, allow_merge_redirect=allow_merge_redirect
        )
        tc = entry.cell_element
        old = cell_text(tc).strip()
        write_cell(tc, value)
        self._pkg.mark_dirty(section_name)
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
        tbl, section_name = self._get_table(table_index)
        entry = self._resolve_anchor_cell(
            tbl, row, col, allow_merge_redirect=allow_merge_redirect
        )
        tc = entry.cell_element
        old = cell_text(tc).strip()
        new_value = f"{old}{separator}{value}" if old else value
        write_cell(tc, new_value)
        self._pkg.mark_dirty(section_name)
        return old

    def append_row(self, table_index: int, values: list[str]) -> None:
        """표 끝에 새 행 추가: 마지막 <hp:tr> deepcopy → 각 셀 비우고
        cellAddr.rowAddr를 새 인덱스로 갱신. 제약은 기존과 동일:
          - 마지막 행이 rowSpan에 걸리면 NotImplementedForFormat
        """
        tbl, section_name = self._get_table(table_index)
        rows_before, _ = table_shape(tbl)

        trs = tbl.findall(HP_TR)
        if not trs:
            raise NotImplementedForFormat("cannot append row to empty HWPX table")

        last_row = trs[-1]
        for tc in last_row.findall(HP_TC):
            from document_adapter.hwpx_core.constants import HP_CELL_SPAN

            span = tc.find(HP_CELL_SPAN)
            addr = tc.find(HP_CELL_ADDR)
            if span is not None and addr is not None:
                try:
                    row_span = int(span.get("rowSpan", "1"))
                    row_addr = int(addr.get("rowAddr", "0"))
                except (TypeError, ValueError):
                    continue
                if row_addr + row_span - 1 != rows_before - 1:
                    raise NotImplementedForFormat(
                        "last row participates in a cross-row merge; "
                        "append_row is not safe for this table."
                    )

        new_row_idx = rows_before
        new_row = deepcopy(last_row)
        for tc in new_row.findall(HP_TC):
            addr = tc.find(HP_CELL_ADDR)
            if addr is not None:
                addr.set("rowAddr", str(new_row_idx))
            # 기존 텍스트만 비우고 run/paragraph 구조는 유지
            sublist = tc.find(HP_SUBLIST)
            if sublist is not None:
                for p in sublist.findall(HP_P):
                    for run in p.findall(HP_RUN):
                        for t in run.findall(HP_T):
                            t.text = ""

        tbl.append(new_row)

        # rowCnt 속성 갱신 (있을 때만)
        row_cnt_attr = tbl.get("rowCnt")
        if row_cnt_attr and row_cnt_attr.isdigit():
            tbl.set("rowCnt", str(int(row_cnt_attr) + 1))

        self._pkg.mark_dirty(section_name)

        # 값 채우기 (병합된 non-anchor 위치는 스킵)
        for i, value in enumerate(values):
            _, cols = table_shape(tbl)
            if i >= cols:
                break
            try:
                self.set_cell(table_index, new_row_idx, i, value)
            except MergedCellWriteError:
                continue
