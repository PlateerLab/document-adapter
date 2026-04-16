"""<hp:p> paragraph 텍스트 읽기/쓰기 헬퍼.

python-hwpx의 ``Paragraph.text`` property를 lxml 수준에서 재현:
  - 읽기: 모든 <hp:run>/<hp:t>의 text 연결
  - 쓰기: 첫 run의 첫 <hp:t>에 text 기록, 같은 run의 다른 <hp:t>는 빈 문자열로,
          이어지는 run들의 <hp:t>도 모두 비워 원본 run 포맷(charPr)을 보존
"""
from __future__ import annotations

from lxml import etree

from document_adapter.hwpx_core.constants import HP_P, HP_RUN, HP_SUBLIST, HP_T, HP_TC


def paragraph_text(p_elem: etree._Element) -> str:
    """<hp:p> 하위의 모든 <hp:t> text를 순서대로 이어 붙인다."""
    parts: list[str] = []
    for run in p_elem.findall(HP_RUN):
        for t in run.findall(HP_T):
            if t.text:
                parts.append(t.text)
    return "".join(parts)


def set_paragraph_text(p_elem: etree._Element, value: str) -> None:
    """첫 <hp:t>에 value, 나머지는 빈 문자열. run 포맷은 그대로.

    <hp:t>가 하나도 없으면 <hp:run> 첫 개체에 <hp:t>를 생성한다.
    <hp:run>도 하나도 없으면 p_elem.text fallback (스타일 없음).
    """
    first_t: etree._Element | None = None
    all_ts: list[etree._Element] = []
    for run in p_elem.findall(HP_RUN):
        for t in run.findall(HP_T):
            if first_t is None:
                first_t = t
            all_ts.append(t)

    if first_t is not None:
        first_t.text = value
        for t in all_ts[1:]:
            t.text = ""
        return

    # <hp:t>가 없는 경우: 첫 <hp:run>에 <hp:t> 생성
    first_run = p_elem.find(HP_RUN)
    if first_run is not None:
        t = etree.SubElement(first_run, HP_T)
        t.text = value
        return

    # <hp:run>도 없음 — 드문 경우. 직접 text 할당.
    p_elem.text = value


def cell_paragraphs(tc_elem: etree._Element) -> list[etree._Element]:
    """셀의 직접 자식 <hp:p> (중첩 테이블의 p는 제외).

    HWPX 구조: <hp:tc> → <hp:subList> → <hp:p>+
    """
    sublist = tc_elem.find(HP_SUBLIST)
    if sublist is None:
        return []
    return sublist.findall(HP_P)


def cell_text(tc_elem: etree._Element) -> str:
    """셀의 직접 텍스트. 중첩 <hp:tbl> 내부 텍스트는 제외.

    각 paragraph의 텍스트를 개행으로 join하지 않고 그대로 concat — 현재 adapter의
    ``_cell_text_raw`` 동작과 일치.
    """
    parts: list[str] = []
    for p in cell_paragraphs(tc_elem):
        parts.append(paragraph_text(p))
    return "".join(parts)


def cell_paragraph_texts(tc_elem: etree._Element) -> list[str]:
    """셀의 각 paragraph 텍스트 리스트."""
    return [paragraph_text(p) for p in cell_paragraphs(tc_elem)]


def write_cell(tc_elem: etree._Element, value: str) -> None:
    """셀 첫 paragraph에 value, 나머지는 비움 (run 스타일 보존)."""
    paragraphs = cell_paragraphs(tc_elem)
    if not paragraphs:
        return
    set_paragraph_text(paragraphs[0], value)
    for p in paragraphs[1:]:
        set_paragraph_text(p, "")


def nested_tables(tc_elem: etree._Element) -> list[etree._Element]:
    """셀 안에 직접 포함된 <hp:tbl> 요소들 (첫 subList/p 레벨에서).

    HWPX 구조에서 중첩 테이블은 <hp:tc> → <hp:subList> → <hp:p> → <hp:run> → <hp:tbl>.
    """
    from document_adapter.hwpx_core.constants import HP_TBL

    results: list[etree._Element] = []
    for p in cell_paragraphs(tc_elem):
        for run in p.findall(HP_RUN):
            for child in run.findall(HP_TBL):
                results.append(child)
    return results
