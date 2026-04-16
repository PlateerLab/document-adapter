"""DocumentAdapter 공통 인터페이스.

세 포맷(DOCX/PPTX/HWPX)의 공통 작업을 추상화:
- 템플릿 렌더링 ({{key}} 치환)
- 표 스키마 추출 (LLM 입력용)
- 셀 수정 / 값 추가 / 행 추가
- 라벨 기반 일괄 채우기 (fill_form)
"""
from __future__ import annotations

import re
from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any


_LABEL_NORMALIZE_RE = re.compile(r"[\s·・／/\-:：()\[\]<>*#]+")


def _normalize_label(s: str) -> str:
    """라벨 정규화: 공백/특수문자 제거 + 소문자.

    "성 명" == "성명", "접수번호" == "접수 번호", "Name:" == "name" 매칭.
    """
    if not s:
        return ""
    return _LABEL_NORMALIZE_RE.sub("", s).lower().strip()


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
    column_widths_cm / row_heights_cm 는 셀의 물리적 크기 (cm, 1자리 반올림).
    LLM 이 긴 값을 좁은 셀에 넣어 오버플로 시키는 것을 방지하기 위한 힌트.
    포맷/문서가 크기 정보를 제공하지 않으면 ``None``.
    """
    index: int
    rows: int
    cols: int
    preview: list[list[str | None]]
    location: str | None = None
    merges: list[MergeInfo] = field(default_factory=list)
    parent_path: str | None = None
    column_widths_cm: list[float] | None = None
    row_heights_cm: list[float] | None = None

    def to_dict(self) -> dict[str, Any]:
        d: dict[str, Any] = {
            "index": self.index,
            "rows": self.rows,
            "cols": self.cols,
            "location": self.location,
            "parent_path": self.parent_path,
            "preview": self.preview,
            "merges": [m.to_dict() for m in self.merges],
        }
        if self.column_widths_cm is not None:
            d["column_widths_cm"] = self.column_widths_cm
        if self.row_heights_cm is not None:
            d["row_heights_cm"] = self.row_heights_cm
        return d


@dataclass
class CellContent:
    """단일 셀의 전체 내용 + 병합/중첩 메타 + 크기 힌트.

    프리뷰가 max_cell_len으로 잘리는 것과 달리 ``text``는 전체 본문을 보유.
    width_cm / height_cm 는 LLM 이 긴 값 오버플로를 피할 수 있도록 제공하는
    셀 크기 힌트 (anchor 기준 span 적용, cm 1자리). char_count 는 ``text`` 길이.
    """
    row: int
    col: int
    text: str                         # 셀 전체 평문 (자르지 않음)
    paragraphs: list[str]             # paragraph 단위 분리
    is_anchor: bool                   # (row,col)이 병합 앵커인지
    anchor: tuple[int, int]           # 이 logical 위치의 앵커 좌표
    span: tuple[int, int]             # (row_span, col_span)
    nested_table_indices: list[int] = field(default_factory=list)
    width_cm: float | None = None
    height_cm: float | None = None
    char_count: int | None = None

    def to_dict(self) -> dict[str, Any]:
        d: dict[str, Any] = {
            "row": self.row,
            "col": self.col,
            "text": self.text,
            "paragraphs": self.paragraphs,
            "is_anchor": self.is_anchor,
            "anchor": list(self.anchor),
            "span": list(self.span),
            "nested_table_indices": self.nested_table_indices,
        }
        if self.width_cm is not None:
            d["width_cm"] = self.width_cm
        if self.height_cm is not None:
            d["height_cm"] = self.height_cm
        if self.char_count is not None:
            d["char_count"] = self.char_count
        return d


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

    # ---- label-based form filling ----
    def fill_form(
        self,
        data: dict[str, str],
        *,
        direction: str = "auto",
        strict: bool = False,
    ) -> dict[str, Any]:
        """라벨 이름으로 값 셀을 찾아 일괄 채우기.

        LLM 이 좌표 (table_index, row, col) 를 직접 계산하지 않고 "접수번호",
        "성명" 같은 **사람이 읽는 라벨** 로 값을 넣게 하는 API. 양식 문서의
        전형적인 라벨-값 패턴 (라벨 오른쪽 또는 아래 셀이 값) 을 자동 탐지한다.

        Args:
            data: {라벨: 값} dict. 라벨은 셀 텍스트와 whitespace/특수문자 제거
                후 정규화 비교 (예: "성 명" == "성명").
            direction: 값 셀 탐색 방향.
                - "auto" (기본): right → below → same 순서로 빈 셀 우선
                - "right": 라벨 셀 오른쪽
                - "below": 라벨 셀 아래
                - "same": 라벨 셀 자체에 append_to_cell
            strict: True 면 매칭 실패 시 ValueError. False 면 skip + not_found 에 기록.

        Returns:
            {
              "filled":   [{"label", "table_index", "row", "col", "action", "old_value", "new_value"}],
              "not_found": [라벨 목록],
              "ambiguous": [{"label", "candidates": [(t,r,c), ...]}],
            }
        """
        if direction not in ("auto", "right", "below", "same"):
            raise ValueError(
                f"direction must be one of auto/right/below/same, got {direction!r}"
            )

        # 전 표의 anchor cell 맵 수집 — (정규화라벨 → [(t, r, c, rowspan, colspan, current_text)])
        # 동일 라벨이 여러 곳에 있으면 ambiguous 로 분류.
        label_index: dict[str, list[tuple[int, int, int, int, int, str]]] = {}
        tables = self.get_tables(preview_rows=10_000, max_cell_len=10_000)
        for t in tables:
            merge_map = {m.anchor: m.span for m in t.merges}
            for r, row in enumerate(t.preview):
                for c, val in enumerate(row):
                    if val is None:
                        continue  # non-anchor
                    text_norm = _normalize_label(val)
                    if not text_norm:
                        continue
                    rs, cs = merge_map.get((r, c), (1, 1))
                    label_index.setdefault(text_norm, []).append(
                        (t.index, r, c, rs, cs, val)
                    )

        filled: list[dict[str, Any]] = []
        not_found: list[str] = []
        ambiguous: list[dict[str, Any]] = []

        for label, value in data.items():
            key = _normalize_label(label)
            candidates = label_index.get(key, [])
            if not candidates:
                if strict:
                    raise ValueError(f"label not found: {label!r}")
                not_found.append(label)
                continue
            if len(candidates) > 1:
                # 여러 곳에 있으면 채우지 않음 — LLM 이 명확히 지정하게 유도
                ambiguous.append({
                    "label": label,
                    "candidates": [(t, r, c) for (t, r, c, *_) in candidates],
                })
                continue

            t_idx, r, c, rs, cs, current_text = candidates[0]
            # "옆 셀이 다른 라벨이면 skip" 판단에 사용자가 요청한 라벨들만 cross-check.
            # (label_index 전체를 쓰면 값 셀 텍스트까지 라벨로 오판 → 덮어쓰기 실패)
            user_keys = {_normalize_label(k) for k in data.keys()}
            other_label_keys = user_keys - {key}
            action, coord, old = self._fill_one_cell(
                t_idx, r, c, rs, cs, str(value), direction, other_label_keys
            )
            if coord is None:
                # 탐색 실패 — fallback 으로 same 에 append_to_cell 권할지 고민했지만
                # 의도치 않은 라벨 오염 방지 차원에서 그냥 not_found 에 기록.
                not_found.append(label)
                continue
            filled.append({
                "label": label,
                "table_index": coord[0],
                "row": coord[1],
                "col": coord[2],
                "action": action,
                "old_value": old,
                "new_value": str(value),
            })

        return {"filled": filled, "not_found": not_found, "ambiguous": ambiguous}

    def _fill_one_cell(
        self,
        t_idx: int,
        r: int,
        c: int,
        rowspan: int,
        colspan: int,
        value: str,
        direction: str,
        other_label_keys: set[str],
    ) -> tuple[str, tuple[int, int, int] | None, str]:
        """한 라벨에 대해 direction 에 따라 값 셀을 찾아 값을 쓴다.

        auto 모드 규칙:
          - right → below → same 순서
          - right/below: target cell 이 **다른 라벨** 이면 skip (덮어쓰면 라벨 오염).
            그 외 (빈 셀 또는 예시 값 있는 값 셀) → set_cell 로 덮어쓰기.
          - 마지막 same: 현재 라벨 셀에 append_to_cell.

        Returns: (action, (table_idx, row, col) or None, old_value)
        """
        if direction == "same":
            order = ["same"]
        elif direction == "right":
            order = ["right"]
        elif direction == "below":
            order = ["below"]
        else:  # auto
            order = ["right", "below", "same"]

        for mode in order:
            if mode == "right":
                target_r, target_c = r, c + colspan
            elif mode == "below":
                target_r, target_c = r + rowspan, c
            else:  # same
                target_r, target_c = r, c

            try:
                cell = self.get_cell(t_idx, target_r, target_c)
            except (IndexError, ValueError):
                continue

            target_text = (cell.text or "").strip()
            target_key = _normalize_label(target_text)

            if mode == "same":
                old = self.append_to_cell(
                    t_idx, target_r, target_c, value,
                    allow_merge_redirect=not cell.is_anchor,
                )
                return "append_to_cell", (t_idx, target_r, target_c), old

            # right / below
            if direction == "auto" and target_key and target_key in other_label_keys:
                # 옆 셀이 다른 라벨 → 덮어쓰면 라벨 손상. 다음 mode 시도.
                continue
            try:
                old = self.set_cell(
                    t_idx, target_r, target_c, value,
                    allow_merge_redirect=not cell.is_anchor,
                )
            except MergedCellWriteError:
                continue
            return "set_cell", (t_idx, target_r, target_c), old

        return "none", None, ""

    # ---- context manager ----
    def __enter__(self) -> "DocumentAdapter":
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        self.close()
