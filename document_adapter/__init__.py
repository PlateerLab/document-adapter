"""Document template editing — 통합 어댑터.

사용법:
    from document_adapter import load
    doc = load("report.docx")
    schema = doc.get_schema()
    doc.set_cell(0, 1, 1, "홍길동")
    doc.save("report_filled.docx")
"""
from __future__ import annotations

from pathlib import Path

from .base import DocumentAdapter, DocumentSchema, TableSchema
from .docx_adapter import DocxAdapter
from .hwpx_adapter import HwpxAdapter
from .pptx_adapter import PptxAdapter

__all__ = [
    "load",
    "DocumentAdapter",
    "DocumentSchema",
    "TableSchema",
    "DocxAdapter",
    "PptxAdapter",
    "HwpxAdapter",
]

_ADAPTERS: dict[str, type[DocumentAdapter]] = {
    ".docx": DocxAdapter,
    ".pptx": PptxAdapter,
    ".hwpx": HwpxAdapter,
}


def load(path: str | Path) -> DocumentAdapter:
    """확장자로 적절한 어댑터를 선택해 문서를 연다."""
    p = Path(path)
    suffix = p.suffix.lower()
    cls = _ADAPTERS.get(suffix)
    if cls is None:
        raise ValueError(
            f"지원하지 않는 포맷: {suffix}. 지원: {sorted(_ADAPTERS.keys())}"
        )
    return cls(p)
