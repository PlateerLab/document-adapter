"""HWPX 어댑터: python-hwpx 기반.

버그 회피:
- set_cell_text()는 빈 셀에서 lxml/ElementTree 혼용 에러가 발생 (v2.9.0) →
  cell.paragraphs[0].text 직접 할당으로 우회
- replace_text_in_runs()는 한글 공백이 run으로 쪼개질 때 매칭 실패 →
  위치 기반 편집을 권장

표 구조:
- iter_grid()로 병합 셀(rowSpan/colSpan)을 인식해 logical grid를 구성
- 셀 내부 중첩 테이블은 flat DFS로 인덱싱, parent_path로 위치 표시

부가:
- manifest fallback 로그가 기본적으로 매우 시끄러움 → logging 레벨 조정
"""
from __future__ import annotations

import logging
import re
import warnings
from copy import deepcopy
from pathlib import Path
from typing import Any, Iterator

# 경고성 로그 억제 (manifest fallback 등)
logging.getLogger("hwpx").setLevel(logging.ERROR)

from hwpx.document import HwpxDocument

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

_HP_NS = "http://www.hancom.co.kr/hwpml/2011/paragraph"
_HP_T = f"{{{_HP_NS}}}t"
_HP_RUN = f"{{{_HP_NS}}}run"
_HP_TBL = f"{{{_HP_NS}}}tbl"
_HP_TR = f"{{{_HP_NS}}}tr"
_HP_TC = f"{{{_HP_NS}}}tc"
_HP_CELL_ADDR = f"{{{_HP_NS}}}cellAddr"
_HP_CELL_SPAN = f"{{{_HP_NS}}}cellSpan"
_HP_SUBLIST = f"{{{_HP_NS}}}subList"
_HP_P = f"{{{_HP_NS}}}p"


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
    def _iter_tables(self) -> Iterator[tuple[int, Any, str]]:
        """(flat_index, table, parent_path) 순회. 중첩 테이블까지 DFS."""
        idx_counter = [0]

        def walk(tbl, parent_path: str) -> Iterator[tuple[int, Any, str]]:
            current_idx = idx_counter[0]
            idx_counter[0] += 1
            yield current_idx, tbl, parent_path
            seen_cell_ids: set[int] = set()
            for entry in tbl.iter_grid():
                if not entry.is_anchor:
                    continue
                cell = entry.cell
                cell_key = id(cell.element)
                if cell_key in seen_cell_ids:
                    continue
                seen_cell_ids.add(cell_key)
                for child_tbl in cell.tables:
                    child_parent = (
                        f"{parent_path}.tables[{current_idx}].cell"
                        f"({entry.anchor[0]},{entry.anchor[1]})"
                    )
                    yield from walk(child_tbl, child_parent)

        for section in self._doc.sections:
            for para in section.paragraphs:
                for tbl in para.tables:
                    yield from walk(tbl, "")

    def _get_table(self, table_index: int):
        for idx, tbl, _ in self._iter_tables():
            if idx == table_index:
                return tbl
        raise TableIndexError(f"HWPX table index {table_index} not found")

    def _find_grid_entry(self, tbl, row: int, col: int):
        """(row, col)의 HwpxTableGridPosition 반환. 경계/존재 검증 포함."""
        if row < 0 or col < 0 or row >= tbl.row_count or col >= tbl.column_count:
            raise CellOutOfBoundsError(
                f"cell ({row},{col}) out of bounds "
                f"({tbl.row_count}x{tbl.column_count})"
            )
        for entry in tbl.iter_grid():
            if (entry.row, entry.column) == (row, col):
                return entry
        raise CellOutOfBoundsError(
            f"cell ({row},{col}) does not resolve to any physical cell"
        )

    def _resolve_anchor_cell(
        self, tbl, row: int, col: int, *, allow_merge_redirect: bool
    ):
        """(row,col)에 대응하는 앵커 HwpxOxmlTableCell 반환.

        non-anchor 좌표이고 allow_merge_redirect가 False면 MergedCellWriteError.
        True면 앵커로 리디렉트하고 경고를 남긴다.
        """
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

    @staticmethod
    def _cell_text_raw(cell) -> str:
        """셀의 직접 텍스트만 추출 (중첩 테이블 제외). strip하지 않은 원문."""
        parts: list[str] = []
        for para in cell.paragraphs:
            for run in para.element.findall(_HP_RUN):
                for t in run.findall(_HP_T):
                    if t.text:
                        parts.append(t.text)
        return "".join(parts)

    @classmethod
    def _cell_text(cls, cell) -> str:
        """프리뷰용: strip된 텍스트."""
        return cls._cell_text_raw(cell).strip()

    @staticmethod
    def _cell_paragraph_texts(cell) -> list[str]:
        """셀의 각 paragraph 텍스트 (중첩 테이블 텍스트 제외)."""
        out: list[str] = []
        for para in cell.paragraphs:
            parts = []
            for run in para.element.findall(_HP_RUN):
                for t in run.findall(_HP_T):
                    if t.text:
                        parts.append(t.text)
            out.append("".join(parts))
        return out

    # ---- inspection ----
    def get_placeholders(self) -> list[str]:
        text = self._doc.export_text()
        return sorted(set(TAG_PATTERN.findall(text)))

    def get_tables(self, min_rows: int = 1, min_cols: int = 1,
                   preview_rows: int = 4, max_cell_len: int = 40) -> list[TableSchema]:
        schemas: list[TableSchema] = []
        for idx, tbl, parent_path in self._iter_tables():
            rows, cols = tbl.row_count, tbl.column_count
            if rows < min_rows or cols < min_cols:
                continue

            visible_rows = min(rows, preview_rows)
            preview: list[list[str | None]] = [
                [None for _ in range(cols)] for _ in range(visible_rows)
            ]
            merges: list[MergeInfo] = []
            seen_anchors: set[tuple[int, int]] = set()

            for entry in tbl.iter_grid():
                if entry.anchor in seen_anchors:
                    continue
                if entry.is_anchor:
                    seen_anchors.add(entry.anchor)
                    if entry.row < visible_rows:
                        text = self._cell_text(entry.cell)
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
        """셀 단건 조회. 전체 텍스트 + 병합/중첩 메타 반환."""
        tbl = self._get_table(table_index)
        entry = self._find_grid_entry(tbl, row, col)

        anchor_cell = entry.cell
        text = self._cell_text_raw(anchor_cell)
        paragraphs = self._cell_paragraph_texts(anchor_cell)

        # 중첩 테이블 flat index 찾기
        nested_indices: list[int] = []
        if entry.is_anchor and list(anchor_cell.tables):
            # _iter_tables로 descendant를 훑어 이 앵커 셀에서 파생된 테이블의 index를 수집
            nested_cell_ids = {id(t.element) for t in anchor_cell.tables}
            for child_idx, child_tbl, _ in self._iter_tables():
                if id(child_tbl.element) in nested_cell_ids:
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

    # ---- editing ----
    def render_template(self, context: dict[str, Any]) -> None:
        """본문 + 표 셀의 {{key}}를 paragraph 단위로 치환."""

        def substitute(para) -> None:
            text = para.text
            if TAG_PATTERN.search(text):
                para.text = TAG_PATTERN.sub(
                    lambda m: str(context.get(m.group(1), m.group(0))), text
                )

        for section in self._doc.sections:
            for para in section.paragraphs:
                substitute(para)
        for _, tbl, _ in self._iter_tables():
            for entry in tbl.iter_grid():
                if not entry.is_anchor:
                    continue
                for para in entry.cell.paragraphs:
                    substitute(para)

    def _write_cell(self, cell, value: str) -> None:
        """셀 첫 paragraph에 value를 쓰고 나머지는 비운다 (기존 run 스타일 보존)."""
        paragraphs = list(cell.paragraphs)
        if paragraphs:
            paragraphs[0].text = value
            for p in paragraphs[1:]:
                p.text = ""

    def set_cell(
        self,
        table_index: int,
        row: int,
        col: int,
        value: str,
        *,
        allow_merge_redirect: bool = False,
    ) -> str:
        """셀 값 교체. 원래 값 반환.

        병합 셀(non-anchor) 좌표는 기본적으로 ``MergedCellWriteError``.
        ``allow_merge_redirect=True``면 앵커로 리디렉트 + 경고.
        """
        tbl = self._get_table(table_index)
        entry = self._resolve_anchor_cell(
            tbl, row, col, allow_merge_redirect=allow_merge_redirect
        )
        cell = entry.cell
        old = self._cell_text(cell)
        self._write_cell(cell, value)
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
        """기존 텍스트 뒤에 ``separator + value``를 덧붙임. 원래 값 반환.

        라벨(예: ``"성  명"``)을 유지한 채 값을 추가하는 용도.
        빈 셀이면 separator 없이 value만 기록.
        """
        tbl = self._get_table(table_index)
        entry = self._resolve_anchor_cell(
            tbl, row, col, allow_merge_redirect=allow_merge_redirect
        )
        cell = entry.cell
        old = self._cell_text(cell)
        new_value = f"{old}{separator}{value}" if old else value
        self._write_cell(cell, new_value)
        return old

    def append_row(self, table_index: int, values: list[str]) -> None:
        """표 끝에 새 행 추가 (last row 구조를 복제).

        python-hwpx에 공식 add_row API는 없지만, 마지막 ``<hp:tr>``을 deepcopy
        한 뒤 각 셀의 텍스트를 비우고 ``cellAddr.rowAddr``를 새 인덱스로 갱신
        하는 방식으로 구현한다. 제약:

        - 마지막 행이 위 행의 rowSpan에 포함된 경우(교차 병합)는 지원하지 않음
        - 각 셀의 서식/폭은 마지막 행의 것을 그대로 상속
        """
        tbl = self._get_table(table_index)
        tbl_elem = tbl.element

        rows = tbl_elem.findall(_HP_TR)
        if not rows:
            raise NotImplementedForFormat("cannot append row to empty HWPX table")

        last_row = rows[-1]
        # 마지막 행에 위에서 내려오는 rowSpan 흔적이 있으면 복제 위험 → 거부
        for tc in last_row.findall(_HP_TC):
            span = tc.find(_HP_CELL_SPAN)
            addr = tc.find(_HP_CELL_ADDR)
            if span is not None and addr is not None:
                try:
                    row_span = int(span.get("rowSpan", "1"))
                    row_addr = int(addr.get("rowAddr", "0"))
                except (TypeError, ValueError):
                    continue
                if row_addr + row_span - 1 != tbl.row_count - 1:
                    raise NotImplementedForFormat(
                        "last row participates in a cross-row merge; "
                        "append_row is not safe for this table."
                    )

        new_row_idx = tbl.row_count
        new_row = deepcopy(last_row)
        for tc in new_row.findall(_HP_TC):
            addr = tc.find(_HP_CELL_ADDR)
            if addr is not None:
                addr.set("rowAddr", str(new_row_idx))
            # 기존 셀 텍스트 비우기: <hp:subList>/<hp:p>/<hp:run>/<hp:t> 모두 유지하고 text만 clear
            sublist = tc.find(_HP_SUBLIST)
            if sublist is not None:
                for p in sublist.findall(_HP_P):
                    for run in p.findall(_HP_RUN):
                        for t in run.findall(_HP_T):
                            t.text = ""

        tbl_elem.append(new_row)

        # rowCnt 속성 갱신 (있을 때만)
        row_cnt_attr = tbl_elem.get("rowCnt")
        if row_cnt_attr and row_cnt_attr.isdigit():
            tbl_elem.set("rowCnt", str(int(row_cnt_attr) + 1))

        # section dirty 처리
        tbl.mark_dirty()

        # 값 채우기 (병합된 non-anchor 위치는 스킵)
        for i, value in enumerate(values):
            if i >= tbl.column_count:
                break
            try:
                self.set_cell(table_index, new_row_idx, i, value)
            except MergedCellWriteError:
                # 복제 시 상속된 병합이 있으면 non-anchor 좌표는 스킵
                continue
