"""DocumentAdapter 공통 인터페이스.

세 포맷(DOCX/PPTX/HWPX)의 공통 작업을 추상화:
- 템플릿 렌더링 ({{key}} 치환)
- 표 스키마 추출 (LLM 입력용)
- 셀 수정 / 값 추가 / 행 추가
"""
from __future__ import annotations

from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any


# ---- custom exceptions -----------------------------------------------------
# 표준 예외를 상속해 기존 ``except ValueError/IndexError`` 흐름과 호환.

class MergedCellWriteError(ValueError):
    """병합 영역의 non-anchor 좌표에 쓰기를 시도했을 때."""


class CellOutOfBoundsError(IndexError):
    """(row, col)이 표 경계를 벗어남."""


class TableIndexError(IndexError):
    """table_index로 표를 찾지 못함."""


class NotImplementedForFormat(NotImplementedError):
    """특정 포맷이 지원하지 않는 연산."""


# ---- dataclasses -----------------------------------------------------------


@dataclass
class MergeInfo:
    """병합 셀 정보. anchor=(row,col)에서 span=(rows,cols)만큼 병합."""
    anchor: tuple[int, int]
    span: tuple[int, int]

    def to_dict(self) -> dict[str, Any]:
        return {"anchor": list(self.anchor), "span": list(self.span)}


@dataclass
class TableSchema:
    """표 한 개의 구조 (LLM에게 넘길 형태).

    preview는 logical grid(rows × cols) 형태. 병합된 non-anchor 슬롯은 ``None``.
    merges는 span>1x1인 앵커 목록 (LLM이 병합 구조를 재구성할 수 있게).
    parent_path는 중첩 테이블 위치 표시 (예: ``"tables[0].cell(1,2)"``).
    """
    index: int
    rows: int
    cols: int
    preview: list[list[str | None]]
    location: str | None = None
    merges: list[MergeInfo] = field(default_factory=list)
    parent_path: str | None = None

    def to_dict(self) -> dict[str, Any]:
        return {
            "index": self.index,
            "rows": self.rows,
            "cols": self.cols,
            "location": self.location,
            "parent_path": self.parent_path,
            "preview": self.preview,
            "merges": [m.to_dict() for m in self.merges],
        }


@dataclass
class CellContent:
    """단일 셀의 전체 내용 + 병합/중첩 메타.

    프리뷰가 max_cell_len으로 잘리는 것과 달리 ``text``는 전체 본문을 보유.
    """
    row: int
    col: int
    text: str                         # 셀 전체 평문 (자르지 않음)
    paragraphs: list[str]             # paragraph 단위 분리
    is_anchor: bool                   # (row,col)이 병합 앵커인지
    anchor: tuple[int, int]           # 이 logical 위치의 앵커 좌표
    span: tuple[int, int]             # (row_span, col_span)
    nested_table_indices: list[int] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return {
            "row": self.row,
            "col": self.col,
            "text": self.text,
            "paragraphs": self.paragraphs,
            "is_anchor": self.is_anchor,
            "anchor": list(self.anchor),
            "span": list(self.span),
            "nested_table_indices": self.nested_table_indices,
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

    @abstractmethod
    def get_cell(self, table_index: int, row: int, col: int) -> CellContent:
        """셀 단건 조회. preview로 잘린 전체 텍스트와 병합/중첩 메타 반환."""

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
    def append_to_cell(self, table_index: int, row: int, col: int, value: str,
                       separator: str = "  ") -> str:
        """기존 셀 텍스트 뒤에 ``separator + value``를 덧붙임.

        라벨(예: "성  명")을 유지한 채 사용자 입력을 추가하는 용도.
        빈 셀이면 separator 없이 value만 기록. 원래 값 반환.
        """

    @abstractmethod
    def append_row(self, table_index: int, values: list[str]) -> None:
        """표 끝에 새 행 추가."""

    # ---- context manager ----
    def __enter__(self) -> "DocumentAdapter":
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        self.close()
