"""<hp:tbl> 에서 병합 셀 인식된 logical grid를 생성.

python-hwpx의 ``Table.iter_grid()``를 대체. xgen-doc2chunk의 ``_build_cell_grid``
패턴을 차용하되, **읽기만이 아니라 쓰기도 지원**하기 위해 각 ``GridEntry``가
anchor cell의 lxml ``<hp:tc>`` Element 자체를 노출한다.

HWPX 표 구조:
  <hp:tbl rowCnt colCnt>
    <hp:tr>+
      <hp:tc>+
        <hp:cellAddr rowAddr colAddr>
        <hp:cellSpan rowSpan colSpan>  (병합 시만 존재 또는 >1)
        <hp:subList> → <hp:p> → <hp:run> → <hp:t>
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Iterator

from lxml import etree

from document_adapter.hwpx_core.constants import HP_CELL_ADDR, HP_CELL_SPAN, HP_TC, HP_TR


@dataclass(frozen=True)
class GridEntry:
    """Logical grid의 한 슬롯.

    - ``is_anchor=True``: 이 (row, col)이 셀의 anchor 좌표
    - ``is_anchor=False``: 병합된 셀에 덮여 있는 non-anchor 슬롯. ``anchor``가 앵커 좌표.
    ``cell_element``는 항상 anchor 셀의 ``<hp:tc>``.
    """
    row: int
    column: int
    is_anchor: bool
    anchor: tuple[int, int]
    span: tuple[int, int]  # (rowspan, colspan)
    cell_element: etree._Element


def _parse_cell_position(tc: etree._Element) -> tuple[int, int]:
    addr = tc.find(HP_CELL_ADDR)
    if addr is None:
        return 0, 0
    try:
        row = int(addr.get("rowAddr", "0"))
    except (TypeError, ValueError):
        row = 0
    try:
        col = int(addr.get("colAddr", "0"))
    except (TypeError, ValueError):
        col = 0
    return row, col


def _parse_cell_span(tc: etree._Element) -> tuple[int, int]:
    span = tc.find(HP_CELL_SPAN)
    if span is None:
        return 1, 1
    try:
        rs = int(span.get("rowSpan", "1"))
    except (TypeError, ValueError):
        rs = 1
    try:
        cs = int(span.get("colSpan", "1"))
    except (TypeError, ValueError):
        cs = 1
    return max(1, rs), max(1, cs)


def table_shape(tbl: etree._Element) -> tuple[int, int]:
    """표의 (rows, cols). rowCnt/colCnt 속성이 없으면 앵커 좌표로 추정."""
    rows = _safe_int(tbl.get("rowCnt"))
    cols = _safe_int(tbl.get("colCnt"))
    if rows > 0 and cols > 0:
        return rows, cols

    max_row = -1
    max_col = -1
    for tr in tbl.findall(HP_TR):
        for tc in tr.findall(HP_TC):
            r, c = _parse_cell_position(tc)
            rs, cs = _parse_cell_span(tc)
            max_row = max(max_row, r + rs - 1)
            max_col = max(max_col, c + cs - 1)
    return max(rows, max_row + 1), max(cols, max_col + 1)


def _safe_int(v: str | None) -> int:
    if v is None:
        return 0
    try:
        return int(v)
    except ValueError:
        return 0


def iter_grid(tbl: etree._Element) -> Iterator[GridEntry]:
    """<hp:tbl> 요소를 병합 셀 인식해 logical grid 순서로 순회.

    row-major: (0,0), (0,1), ..., (0,C-1), (1,0), ...
    """
    rows, cols = table_shape(tbl)
    if rows <= 0 or cols <= 0:
        return

    # 1) anchor 좌표 → {span, cell_element}
    anchors: dict[tuple[int, int], tuple[tuple[int, int], etree._Element]] = {}
    for tr in tbl.findall(HP_TR):
        for tc in tr.findall(HP_TC):
            r, c = _parse_cell_position(tc)
            span = _parse_cell_span(tc)
            anchors[(r, c)] = (span, tc)

    # 2) 각 (row, col)이 어느 anchor에 속하는지 역산 매핑
    owner: dict[tuple[int, int], tuple[int, int]] = {}
    for (ar, ac), (span, _tc) in anchors.items():
        rs, cs = span
        for dr in range(rs):
            for dc in range(cs):
                slot = (ar + dr, ac + dc)
                # 여러 anchor가 같은 slot을 주장하면 가장 가까운(= 자기 자신) 우선
                if slot in owner:
                    # 이미 자기 자신으로 설정됐다면 유지
                    continue
                owner[slot] = (ar, ac)

    # 3) row-major 순회
    for r in range(rows):
        for c in range(cols):
            slot = (r, c)
            anchor_coord = owner.get(slot)
            if anchor_coord is None:
                # grid에 선언되지 않은 슬롯 (손상 문서). 빈 앵커 취급.
                continue
            span, cell_el = anchors[anchor_coord]
            is_anchor = (anchor_coord == slot)
            yield GridEntry(
                row=r,
                column=c,
                is_anchor=is_anchor,
                anchor=anchor_coord,
                span=span,
                cell_element=cell_el,
            )
