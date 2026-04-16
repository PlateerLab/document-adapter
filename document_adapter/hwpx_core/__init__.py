"""HWPX 저수준 패키지 — python-hwpx 의존 없이 zipfile+lxml로 HWPX 문서 편집.

이 패키지는 표 편집에 필요한 최소 기능만 제공한다:
- ZIP 컨테이너 입출력 (수정 안 한 파일은 bytes 그대로 보존)
- XML 트리의 lazy 파싱과 dirty 추적
- 병합 셀 인식된 logical grid 순회

구조:
- constants: HWPX XML 네임스페이스
- package.HwpxPackage: python-hwpx의 HwpxDocument 대체
- grid.iter_grid: python-hwpx의 Table.iter_grid() 대체
"""
from document_adapter.hwpx_core.constants import (
    HC_NS,
    HH_NS,
    HP_NS,
    HS_NS,
    OPF_NS,
    HP_CELL_ADDR,
    HP_CELL_SPAN,
    HP_CELL_SZ,
    HP_P,
    HP_RUN,
    HP_SUBLIST,
    HP_T,
    HP_TBL,
    HP_TC,
    HP_TR,
)
from document_adapter.hwpx_core.grid import GridEntry, iter_grid, table_shape
from document_adapter.hwpx_core.package import HwpxPackage
from document_adapter.hwpx_core.paragraph import (
    cell_paragraph_texts,
    cell_paragraphs,
    cell_text,
    nested_tables,
    paragraph_text,
    set_paragraph_text,
    write_cell,
)

__all__ = [
    "HC_NS",
    "HH_NS",
    "HP_NS",
    "HS_NS",
    "OPF_NS",
    "HP_CELL_ADDR",
    "HP_CELL_SPAN",
    "HP_CELL_SZ",
    "HP_P",
    "HP_RUN",
    "HP_SUBLIST",
    "HP_T",
    "HP_TBL",
    "HP_TC",
    "HP_TR",
    "GridEntry",
    "HwpxPackage",
    "iter_grid",
    "table_shape",
    "cell_paragraph_texts",
    "cell_paragraphs",
    "cell_text",
    "nested_tables",
    "paragraph_text",
    "set_paragraph_text",
    "write_cell",
]
