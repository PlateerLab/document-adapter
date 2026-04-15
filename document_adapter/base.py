"""DocumentAdapter 공통 인터페이스.

세 포맷(DOCX/PPTX/HWPX)의 공통 작업을 추상화:
- 템플릿 렌더링 ({{key}} 치환)
- 표 스키마 추출 (LLM 입력용)
- 셀 수정 / 행 추가
"""
from __future__ import annotations

from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any


@dataclass
class TableSchema:
    """표 한 개의 구조 (LLM에게 넘길 형태)."""
    index: int
    rows: int
    cols: int
    preview: list[list[str]]
    location: str | None = None

    def to_dict(self) -> dict[str, Any]:
        return {
            "index": self.index,
            "rows": self.rows,
            "cols": self.cols,
            "location": self.location,
            "preview": self.preview,
        }


@dataclass
class DocumentSchema:
    """문서 전체 스키마."""
    format: str
    source: str
    placeholders: list[str] = field(default_factory=list)
    tables: list[TableSchema] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return {
            "format": self.format,
            "source": self.source,
            "placeholders": self.placeholders,
            "tables": [t.to_dict() for t in self.tables],
        }


class DocumentAdapter(ABC):
    """모든 포맷 어댑터의 공통 부모."""

    format: str = ""

    def __init__(self, path: Path) -> None:
        self.path = Path(path)
        self._open()

    # ---- lifecycle ----
    @abstractmethod
    def _open(self) -> None: ...

    @abstractmethod
    def save(self, path: Path | str | None = None) -> Path: ...

    def close(self) -> None:
        """일부 포맷(HWPX)은 명시적 close 필요."""
        pass

    # ---- inspection ----
    @abstractmethod
    def get_placeholders(self) -> list[str]:
        """본문에서 사용된 {{key}} 목록 반환."""

    @abstractmethod
    def get_tables(self, min_rows: int = 1, min_cols: int = 1,
                   preview_rows: int = 4, max_cell_len: int = 40) -> list[TableSchema]:
        """필터 조건을 만족하는 표 스키마 목록."""

    def get_schema(self) -> DocumentSchema:
        return DocumentSchema(
            format=self.format,
            source=str(self.path),
            placeholders=self.get_placeholders(),
            tables=self.get_tables(),
        )

    # ---- editing ----
    @abstractmethod
    def render_template(self, context: dict[str, Any]) -> None:
        """템플릿의 {{key}}를 context 값으로 치환."""

    @abstractmethod
    def set_cell(self, table_index: int, row: int, col: int, value: str) -> str:
        """셀 값 교체. 원래 값 반환."""

    @abstractmethod
    def append_row(self, table_index: int, values: list[str]) -> None:
        """표 끝에 새 행 추가."""

    # ---- context manager ----
    def __enter__(self) -> "DocumentAdapter":
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        self.close()
