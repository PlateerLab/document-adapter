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
    HP_CELL_SPAN,
    HP_CELL_SZ,
    HP_DRAW_TEXT,
    HP_P,
    HP_RUN,
    HP_SUBLIST,
    HP_T,
    HP_TBL,
    HP_TC,
    HP_TR,
    HP_FORM_TEXT,
    HP_LIST_ITEM,
    FORM_CONTROL_TAGS,
    HwpxPackage,
    cell_paragraph_texts,
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
    _has_template,
    _ParaHandle,
)

TAG_PATTERN = re.compile(r"\{\{\s*(\w+)\s*\}\}")

# HWPX HU (Hwp Unit) → cm. 1 cm = 7200/2.54 ≈ 2834.6457 HU
_HU_PER_CM = 2834.6456692913


def _hu_to_cm(hu: Any) -> float | None:
    if hu is None:
        return None
    try:
        v = int(hu)
    except (TypeError, ValueError):
        return None
    if v <= 0:
        return None
    return round(v / _HU_PER_CM, 1)


def _cell_size_cm(tc_elem) -> tuple[float | None, float | None]:
    """<hp:tc>/<hp:cellSz>에서 (width_cm, height_cm) 추출. cellSpan 이 있는
    병합 셀의 경우 HWPX 는 anchor cell 자체 크기를 cellSz 에 저장하므로
    별도 span 합산 불필요."""
    sz = tc_elem.find(HP_CELL_SZ)
    if sz is None:
        return None, None
    return _hu_to_cm(sz.get("width")), _hu_to_cm(sz.get("height"))


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
            # 섹션의 모든 <hp:tbl> descendant 중 "top-level" 만 선별.
            # top-level = <hp:tc> (cell) 의 descendant 가 아닌 것.
            # 이러면 <hp:p>/<hp:run>/<hp:ctrl>/<hp:header>/<hp:footer>/도형 등
            # 어디에 놓여있든 표를 누락 없이 발견 (xgen-doc2chunk 의 _process_ctrl
            # 가 처리하던 범위 포함). cell 내부 <hp:tbl> 은 walk 재귀에서 처리.
            for tbl in root.iter(HP_TBL):
                parent = tbl.getparent()
                is_in_cell = False
                while parent is not None:
                    if parent.tag == HP_TC:
                        is_in_cell = True
                        break
                    parent = parent.getparent()
                if not is_in_cell:
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
            # col/row → cm 매핑. colSpan/rowSpan 이 1 인 앵커 셀에서만 width/height
            # 를 직접 수집 (다른 셀에 걸쳐있지 않아 깔끔). span>1 셀의 cellSz 는 anchor
            # 위치 폭 전체를 표현하므로 일부 col/row 는 None 으로 남을 수 있다.
            col_width_map: dict[int, float] = {}
            row_height_map: dict[int, float] = {}

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
                    rs, cs = entry.span
                    w_cm, h_cm = _cell_size_cm(entry.cell_element)
                    if cs == 1 and w_cm is not None and entry.column not in col_width_map:
                        col_width_map[entry.column] = w_cm
                    if rs == 1 and h_cm is not None and entry.row not in row_height_map:
                        row_height_map[entry.row] = h_cm

            col_widths = [col_width_map.get(c) for c in range(cols)]
            row_heights = [row_height_map.get(r) for r in range(rows)]
            col_widths_out = col_widths if any(v is not None for v in col_widths) else None
            row_heights_out = row_heights if any(v is not None for v in row_heights) else None

            schemas.append(
                TableSchema(
                    index=idx,
                    rows=rows,
                    cols=cols,
                    preview=preview,
                    merges=merges,
                    parent_path=parent_path or None,
                    column_widths_cm=col_widths_out,
                    row_heights_cm=row_heights_out,
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

        width_cm, height_cm = _cell_size_cm(tc)

        return CellContent(
            row=row,
            col=col,
            text=text,
            paragraphs=paragraphs,
            is_anchor=entry.is_anchor,
            anchor=entry.anchor,
            span=entry.span,
            nested_table_indices=nested_indices,
            width_cm=width_cm,
            height_cm=height_cm,
            char_count=len(text),
        )

    # ---- 편집 ----

    def render_template(self, context: dict[str, Any], *,
                        on_missing: str = "blank") -> dict[str, list[str]]:
        """섹션의 모든 <hp:p> 에서 {{key}} 치환. paragraph 단위로 처리해
        run 포맷은 보존한다 (첫 <hp:t>에 치환 결과를 쓰고 나머지는 비움).
        누락 키 처리는 on_missing 정책을 따른다 (base 참조).
        """
        report = self._render_report(self.get_placeholders(), context, on_missing)

        def substitute(p: etree._Element) -> bool:
            text = paragraph_text(p)
            if not _has_template(text):
                return False
            set_paragraph_text(p, self._render_text_block(text, context, on_missing))
            return True

        for section_name, root in self._pkg.iter_section_roots():
            changed = False
            for p in root.iter(HP_P):
                if substitute(p):
                    changed = True
            if changed:
                self._pkg.mark_dirty(section_name)
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
        # v0.14: 위치 지정 insert_row 로 위임 (at_row=None → 맨 끝).
        self.insert_row(table_index, values, at_row=None)

    # ---- 위치 지정 행/열 삽입 (v0.14+) ----

    @staticmethod
    def _tc_addr_span(tc) -> tuple[int, int, int, int]:
        """<hp:tc> 의 (rowAddr, colAddr, rowSpan, colSpan). 파싱 실패 시 기본값."""
        addr = tc.find(HP_CELL_ADDR)
        span = tc.find(HP_CELL_SPAN)

        def _i(el, name: str, default: str) -> int:
            try:
                return int(el.get(name, default)) if el is not None else int(default)
            except (TypeError, ValueError):
                return int(default)

        return (
            _i(addr, "rowAddr", "0"),
            _i(addr, "colAddr", "0"),
            _i(span, "rowSpan", "1"),
            _i(span, "colSpan", "1"),
        )

    @staticmethod
    def _blank_copied_hwpx_tc(tc) -> None:
        """deepcopy 된 <hp:tc> 정리: 중첩 표 제거 + 텍스트 비움 (구조/서식 유지)."""
        for nested in nested_tables(tc):
            parent = nested.getparent()
            if parent is not None:
                parent.remove(nested)
        sublist = tc.find(HP_SUBLIST)
        if sublist is not None:
            for p in sublist.findall(HP_P):
                for run in p.findall(HP_RUN):
                    for t in run.findall(HP_T):
                        t.text = ""

    def insert_row(
        self,
        table_index: int,
        values: list[str],
        at_row: int | None = None,
    ) -> int:
        """지정 위치에 새 행 삽입 — 인접 <hp:tr> deepcopy 로 서식 상속.

        - 템플릿 행: at_row 위치의 기존 행 (맨 끝이면 마지막 행).
        - 복사 셀의 rowSpan 은 1 로 리셋 (새 행은 독립), 중첩 표 제거.
        - 삽입 경계를 세로 병합이 가로지르거나, 템플릿 행이 위쪽 병합에
          물려 전체 열을 갖지 못하면 NotImplementedForFormat.
        - values 는 논리 grid 열 기준 (병합 non-anchor 위치는 스킵).
        """
        tbl, section_name = self._get_table(table_index)
        n_rows, n_cols = table_shape(tbl)
        trs = tbl.findall(HP_TR)
        if not trs:
            raise NotImplementedForFormat(
                "cannot insert a row into an empty HWPX table"
            )
        if at_row is None:
            at_row = n_rows
        if at_row < 0 or at_row > n_rows:
            raise CellOutOfBoundsError(f"at_row {at_row} out of range 0..{n_rows}")

        # 세로 병합 경계 가드
        if 0 < at_row < n_rows:
            for tr in trs:
                for tc in tr.findall(HP_TC):
                    r, _c, rs, _cs = self._tc_addr_span(tc)
                    if r < at_row < r + rs:
                        raise NotImplementedForFormat(
                            f"insertion point row {at_row} crosses a vertical "
                            f"merge anchored at row {r}; inserting here would "
                            f"split the merged region."
                        )

        # 템플릿 행 완전성: 직접 소유한 colSpan 합이 전체 열수여야 복사 가능
        # (위쪽 병합에 물린 행은 일부 열의 tc 가 없어 복사본이 불완전해진다)
        template_idx = at_row if at_row < n_rows else n_rows - 1
        template_tr = trs[template_idx]
        owned_cols = sum(
            self._tc_addr_span(tc)[3] for tc in template_tr.findall(HP_TC)
        )
        if owned_cols != n_cols:
            raise NotImplementedForFormat(
                f"template row {template_idx} participates in a cross-row "
                f"merge (owns {owned_cols}/{n_cols} columns); choose another "
                f"insertion point."
            )

        new_tr = deepcopy(template_tr)
        for tc in new_tr.findall(HP_TC):
            addr = tc.find(HP_CELL_ADDR)
            if addr is not None:
                addr.set("rowAddr", str(at_row))
            span = tc.find(HP_CELL_SPAN)
            if span is not None:
                span.set("rowSpan", "1")
            self._blank_copied_hwpx_tc(tc)

        if at_row < n_rows:
            # 밀려나는 행들의 rowAddr 재번호 (+1)
            for tr in trs[at_row:]:
                for tc in tr.findall(HP_TC):
                    addr = tc.find(HP_CELL_ADDR)
                    if addr is None:
                        continue
                    r, _c, _rs, _cs = self._tc_addr_span(tc)
                    addr.set("rowAddr", str(r + 1))
            trs[at_row].addprevious(new_tr)
        else:
            trs[-1].addnext(new_tr)

        row_cnt_attr = tbl.get("rowCnt")
        if row_cnt_attr and row_cnt_attr.isdigit():
            tbl.set("rowCnt", str(int(row_cnt_attr) + 1))

        self._pkg.mark_dirty(section_name)

        # 값 채우기 (논리 grid 열 기준, 병합 non-anchor 는 스킵)
        for i, value in enumerate(values):
            if i >= n_cols:
                break
            if not value:
                continue
            try:
                self.set_cell(table_index, at_row, i, value)
            except MergedCellWriteError:
                continue
        return at_row

    def insert_column(
        self,
        table_index: int,
        values: list[str],
        at_col: int | None = None,
    ) -> int:
        """지정 위치에 새 열 삽입 — 행별 인접 <hp:tc> deepcopy 로 서식 상속.

        - 템플릿 셀: 각 행에서 왼쪽 이웃 열을 덮는 셀 (at_col=0 이면 오른쪽
          이웃). 해당 행에 없으면(위쪽 세로 병합에 물린 열) 그 행의 첫 셀.
        - 복사 셀의 rowSpan/colSpan 은 1 로 리셋, 중첩 표 제거.
        - 표 전체 폭 유지: 모든 cellSz width 를 비례 축소해 새 열 폭을 흡수.
        - 삽입 경계를 가로 병합(colSpan)이 가로지르면 NotImplementedForFormat.
        - values 는 위 행부터 (values[0] 이 보통 헤더).
        """
        tbl, section_name = self._get_table(table_index)
        n_rows, n_cols = table_shape(tbl)
        trs = tbl.findall(HP_TR)
        if not trs or n_cols == 0:
            raise NotImplementedForFormat(
                "cannot insert a column into an empty HWPX table"
            )
        if at_col is None:
            at_col = n_cols
        if at_col < 0 or at_col > n_cols:
            raise CellOutOfBoundsError(f"at_col {at_col} out of range 0..{n_cols}")

        # 가로 병합 경계 가드
        if 0 < at_col < n_cols:
            for tr in trs:
                for tc in tr.findall(HP_TC):
                    _r, c, _rs, cs = self._tc_addr_span(tc)
                    if c < at_col < c + cs:
                        raise NotImplementedForFormat(
                            f"insertion point column {at_col} crosses a "
                            f"horizontal merge anchored at column {c}; "
                            f"inserting here would split the merged region."
                        )

        template_col = at_col - 1 if at_col > 0 else 0

        # 새 열 폭 = 템플릿 열 폭 (colSpan 병합이면 등분). 표 폭 유지를 위한
        # 비례 축소 계수는 첫 행의 폭 합으로 계산. 폭 정보가 없으면 재배분 생략.
        def _tc_width(tc) -> int | None:
            sz = tc.find(HP_CELL_SZ)
            if sz is None:
                return None
            try:
                return int(sz.get("width", ""))
            except (TypeError, ValueError):
                return None

        template_width: int | None = None
        for tr in trs:
            for tc in tr.findall(HP_TC):
                _r, c, _rs, cs = self._tc_addr_span(tc)
                if c <= template_col < c + cs:
                    w = _tc_width(tc)
                    if w is not None and w > 0:
                        template_width = max(1, w // cs)
                    break
            if template_width is not None:
                break

        scale = 1.0
        if template_width is not None:
            first_widths = [_tc_width(tc) for tc in trs[0].findall(HP_TC)]
            if all(w is not None and w > 0 for w in first_widths) and first_widths:
                total = sum(first_widths)  # type: ignore[arg-type]
                scale = total / (total + template_width)

        # 행별 삽입
        for r_idx, tr in enumerate(trs):
            tcs = tr.findall(HP_TC)
            if not tcs:
                continue
            template_tc = None
            insert_before = None
            row_addr = self._tc_addr_span(tcs[0])[0]
            for tc in tcs:
                _r, c, _rs, cs = self._tc_addr_span(tc)
                if c <= template_col < c + cs:
                    template_tc = tc
                if c >= at_col and insert_before is None:
                    insert_before = tc
                # 밀려나는 열들의 colAddr 재번호 (+1)
                if c >= at_col:
                    addr = tc.find(HP_CELL_ADDR)
                    if addr is not None:
                        addr.set("colAddr", str(c + 1))
                # 폭 비례 축소
                if scale != 1.0:
                    sz = tc.find(HP_CELL_SZ)
                    w = _tc_width(tc)
                    if sz is not None and w is not None and w > 0:
                        sz.set("width", str(max(1, int(round(w * scale)))))

            new_tc = deepcopy(template_tc if template_tc is not None else tcs[0])
            addr = new_tc.find(HP_CELL_ADDR)
            if addr is not None:
                addr.set("rowAddr", str(row_addr))
                addr.set("colAddr", str(at_col))
            span = new_tc.find(HP_CELL_SPAN)
            if span is not None:
                span.set("rowSpan", "1")
                span.set("colSpan", "1")
            if template_width is not None:
                sz = new_tc.find(HP_CELL_SZ)
                if sz is not None:
                    sz.set(
                        "width", str(max(1, int(round(template_width * scale))))
                    )
            self._blank_copied_hwpx_tc(new_tc)
            if insert_before is not None:
                insert_before.addprevious(new_tc)
            else:
                tcs[-1].addnext(new_tc)

        col_cnt_attr = tbl.get("colCnt")
        if col_cnt_attr and col_cnt_attr.isdigit():
            tbl.set("colCnt", str(int(col_cnt_attr) + 1))

        self._pkg.mark_dirty(section_name)

        # 값 채우기 (행 인덱스 기준)
        for r, value in enumerate(values):
            if r >= n_rows:
                break
            if not value:
                continue
            try:
                self.set_cell(table_index, r, at_col, value)
            except (MergedCellWriteError, CellOutOfBoundsError):
                continue
        return at_col

    # ---- paragraph text ops (v0.13+) ----

    def _iter_text_paragraphs(self) -> Iterator[_ParaHandle]:
        """섹션의 모든 <hp:p> 를 문서 순서로 순회 (본문+표 셀+글상자).

        - ``root.iter(HP_P)`` 가 표 셀·글상자(drawText) 내부 문단까지 트리
          순서로 정확히 한 번씩 방문한다 — DOCX 처럼 소스를 나눠 합칠 필요 없음.
        - scope 는 가장 가까운 ancestor 로 판정: HP_TC → "table",
          drawText → "shape", 그 외 "body". 표 문단의 location 은 set_cell 과
          동일한 flat table_index 좌표계로 표기.
        - 텍스트 변경 시 해당 섹션을 mark_dirty — 누락하면 저장이 안 된다.
        - is_heading 은 항상 False (HWPX 스타일 해석은 범위 외 —
          find_text 의 context_before 가 섹션 판단을 대신한다).
        - 한계: <hp:t> 의 자식 요소(형광펜 마커 등) 뒤 tail 텍스트는 다루지
          않음 (기존 paragraph_text / set_paragraph_text 와 동일 범위).
        """
        # 표 요소 → flat index (set_cell 좌표계와 일치).
        # 주의: lxml 프록시는 참조가 사라지면 GC 후 재생성돼 id() 가 바뀐다.
        # flat_tables 리스트로 프록시를 generator 수명 동안 살려둬야
        # root.iter() 가 같은 노드에 대해 같은 프록시(=같은 id)를 돌려준다.
        flat_tables = [(idx, tbl) for idx, tbl, _, _ in self._iter_tables()]
        tbl_index_by_id: dict[int, int] = {
            id(tbl): idx for idx, tbl in flat_tables
        }

        for section_name, root in self._pkg.iter_section_roots():
            for n, p in enumerate(root.iter(HP_P)):
                scope = "body"
                location = f"{section_name}#p[{n}]"
                anc = p.getparent()
                while anc is not None:
                    if anc.tag == HP_TC:
                        scope = "table"
                        tbl_el = anc.getparent()
                        while tbl_el is not None and tbl_el.tag != HP_TBL:
                            tbl_el = tbl_el.getparent()
                        addr = anc.find(HP_CELL_ADDR)
                        row = addr.get("rowAddr", "?") if addr is not None else "?"
                        col = addr.get("colAddr", "?") if addr is not None else "?"
                        t_idx = (
                            tbl_index_by_id.get(id(tbl_el), -1)
                            if tbl_el is not None else -1
                        )
                        location = f"table[{t_idx}].cell({row},{col})"
                        break
                    if anc.tag == HP_DRAW_TEXT:
                        scope = "shape"
                        location = f"{section_name}#p[{n}](shape)"
                        break
                    anc = anc.getparent()

                ts = [t for run in p.findall(HP_RUN)
                      for t in run.findall(HP_T)]

                def get_texts(_ts=ts) -> list[str]:
                    return [(t.text or "") for t in _ts]

                def set_texts(new_texts: list[str], _ts=ts,
                              _sec=section_name) -> None:
                    changed = False
                    for t, new_t in zip(_ts, new_texts):
                        if (t.text or "") != new_t:
                            t.text = new_t
                            changed = True
                    if changed:
                        self._pkg.mark_dirty(_sec)

                yield _ParaHandle(
                    scope=scope,
                    location=location,
                    is_heading=False,
                    get_texts=get_texts,
                    set_texts=set_texts,
                )

    # ---- 폼 컨트롤 (체크박스/라디오/콤보/리스트/에디트) ----
    def get_form_controls(self) -> list[dict[str, Any]]:
        """폼 컨트롤 목록. 표가 아닌 인터랙티브 필드를 노출."""
        out: list[dict[str, Any]] = []
        for _, root in self._pkg.iter_section_roots():
            for el in root.iter():
                kind = FORM_CONTROL_TAGS.get(el.tag)
                if not kind:
                    continue
                info: dict[str, Any] = {
                    "name": el.get("name", ""),
                    "kind": kind,
                    "caption": el.get("caption", ""),
                }
                if kind in ("checkBtn", "radioBtn"):
                    info["value"] = el.get("value", "")
                    info["checked"] = el.get("value", "").upper() == "CHECKED"
                elif kind == "edit":
                    te = el.find(HP_FORM_TEXT)
                    info["value"] = (te.text or "") if te is not None else ""
                else:  # comboBox / listBox: 현재값은 <text> 자식, 옵션은 listItem
                    te = el.find(HP_FORM_TEXT)
                    info["value"] = (te.text or "") if te is not None else ""
                    info["items"] = [
                        li.get("displayText", "")
                        for li in el.findall(HP_LIST_ITEM)
                    ]
                out.append(info)
        return out

    def set_form_control(self, name: str, value: Any) -> str:
        """이름으로 폼 컨트롤 값을 설정하고 기존 값을 반환."""
        truthy = {"y", "yes", "true", "1", "checked", "check",
                  "체크", "선택", "o", "on", "예", "✓"}
        for section_name, root in self._pkg.iter_section_roots():
            for el in root.iter():
                kind = FORM_CONTROL_TAGS.get(el.tag)
                if not kind or el.get("name") != name:
                    continue
                if kind in ("checkBtn", "radioBtn"):
                    old = el.get("value", "")
                    checked = value is True or str(value).strip().lower() in truthy
                    el.set("value", "CHECKED" if checked else "UNCHECKED")
                elif kind in ("edit", "comboBox", "listBox"):
                    # 현재 값은 <text> 자식에 기록 (editable combo 포함)
                    te = el.find(HP_FORM_TEXT)
                    if te is None:
                        te = etree.SubElement(el, HP_FORM_TEXT)
                    old = te.text or ""
                    te.text = str(value)
                else:
                    old = el.get("value", "")
                    el.set("value", str(value))
                self._pkg.mark_dirty(section_name)
                return old
        raise ValueError(f"form control not found: {name!r}")
