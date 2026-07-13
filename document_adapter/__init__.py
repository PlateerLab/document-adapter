"""Document template editing — 통합 어댑터.

사용법:
    from document_adapter import load, create_document

    # 기존 문서 편집
    doc = load("report.docx")
    schema = doc.get_schema()
    doc.set_cell(0, 1, 1, "홍길동")
    doc.save("report_filled.docx")

    # 새 문서 생성 (v0.15+) — markdown / sheet spec → 스타일 잡힌 문서
    create_document("회의록.docx", markdown="# 주간 회의록\\n...", lang="ko")
"""
from __future__ import annotations

from pathlib import Path

from .base import DocumentAdapter, DocumentSchema, TableSchema, TextMatch
from .docx_adapter import DocxAdapter
from .generate import create_document
from .hwpx_adapter import HwpxAdapter
from .pptx_adapter import PptxAdapter
from .xlsx_adapter import XlsxAdapter

__all__ = [
    "load",
    "create_document",
    "DocumentAdapter",
    "DocumentSchema",
    "TableSchema",
    "TextMatch",
    "DocxAdapter",
    "PptxAdapter",
    "HwpxAdapter",
    "XlsxAdapter",
]

_ADAPTERS: dict[str, type[DocumentAdapter]] = {
    ".docx": DocxAdapter,
    ".pptx": PptxAdapter,
    ".hwpx": HwpxAdapter,
    ".xlsx": XlsxAdapter,
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
