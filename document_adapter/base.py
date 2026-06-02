"""DocumentAdapter 공통 인터페이스.

세 포맷(DOCX/PPTX/HWPX)의 공통 작업을 추상화:
- 템플릿 렌더링 ({{key}} 치환)
- 표 스키마 추출 (LLM 입력용)
- 셀 수정 / 값 추가 / 행 추가
- 라벨 기반 일괄 채우기 (fill_form)
"""
from __future__ import annotations

import inspect
import re
from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable


_LABEL_NORMALIZE_RE = re.compile(r"[\s·・／/\-:：()\[\]<>*#]+")


def _estimate_text_width_cm(text: str) -> float:
    """10pt 기준 대략적인 글자폭 합(cm). 한글/CJK ~0.35cm, 그 외 ~0.20cm.

    정밀 측정이 아니라 "이 값이 칸 너비를 넘겨 줄바꿈/세로깨짐 위험이 있나"를
    가늠하기 위한 보수적 추정.
    """
    w = 0.0
    for ch in text:
        w += 0.35 if ord(ch) > 0x2E7F else 0.20
    return w


# 셀 패딩/마진·폰트 추정 오차를 감안한 여유 배율. 경미한 초과는 무시하고
# 심각한 초과(예: 2.0cm 텍스트를 0.4cm 칸에)만 경고하기 위함.
_OVERFLOW_TOLERANCE = 1.15


def _overflow_risk(text: str, width_cm: float | None) -> bool:
    """값이 셀 너비를 넘겨 오버플로(세로 깨짐)될 위험이 있는지.

    width_cm 가 없으면(None) 판단 보류(False). 폭 정보가 있을 때만,
    추정 글자폭이 칸폭의 _OVERFLOW_TOLERANCE 배를 넘을 때 경고한다.
    """
    if not width_cm or width_cm <= 0:
        return False
    return _estimate_text_width_cm(text) > width_cm * _OVERFLOW_TOLERANCE


def _accepts_kwarg(fn: Callable[..., Any], name: str) -> bool:
    """fn 이 ``name`` 키워드 인자(또는 **kwargs)를 받을 수 있는지."""
    try:
        params = inspect.signature(fn).parameters
    except (TypeError, ValueError):
        return True  # introspection 불가 시 그냥 시도
    if name in params:
        return True
    return any(p.kind == p.VAR_KEYWORD for p in params.values())


def _call_cell_writer(fn: Callable[..., str], *args: Any,
                      allow_merge_redirect: bool) -> str:
    """set_cell/append_to_cell 호출 래퍼.

    어댑터가 ``allow_merge_redirect`` 를 구현했으면 전달하고, (구버전 ABC 계약을
    따라) 구현하지 않았으면 생략해 호출한다. fill_form 이 모든 어댑터에서
    동작하도록 보장 — TypeError 로 깨지지 않는다.
    """
    if _accepts_kwarg(fn, "allow_merge_redirect"):
        return fn(*args, allow_merge_redirect=allow_merge_redirect)
    return fn(*args)


def _normalize_label(s: str) -> str:
    """라벨 정규화: 공백/특수문자 제거 + 소문자.

    "성 명" == "성명", "접수번호" == "접수 번호", "Name:" == "name" 매칭.
    """
    if not s:
        return ""
    return _LABEL_NORMALIZE_RE.sub("", s).lower().strip()


def _split_dot_path(label: str) -> tuple[str | None, str]:
    """dot-path 분리: '피해자.금액' → ('피해자', '금액'). dot 없으면 (None, label)."""
    if "." in label:
        section, actual = label.rsplit(".", 1)
        return section.strip() or None, actual.strip()
    return None, label


def _candidate_context_labels(
    candidate: tuple[int, int, int, int, int, str],
    tables_by_idx: dict[int, Any],
) -> list[str]:
    """candidate 셀 주변에서 section/row 라벨 후보 텍스트 수집.

    전형적으로 양식은:
      - 같은 table 의 (r, 0) 또는 (r-1..0, 0) 위치 anchor cell 이 섹션 헤더
      - 또는 (r-1, c), (0, c) 가 header row
    단순 휴리스틱: 같은 row 의 col=0 anchor text + 그 위로 올라가면서 나오는
    col=0 anchor text 몇 개를 수집.
    """
    t_idx, r = candidate[0], candidate[1]
    t = tables_by_idx.get(t_idx)
    if t is None:
        return []

    preview = t.preview
    labels: list[str] = []
    # candidate row 자신 또는 위쪽으로 올라가며 col=0 의 가장 가까운 anchor 1개.
    # 병합 anchor 의 rowSpan 으로 candidate row 가 덮이므로 그게 진짜 섹션 헤더.
    # candidate 가 col=0 자체이면 자기 자신이 아니라 **위쪽** row 의 col=0 라벨.
    candidate_col = candidate[2]
    start_r = min(r, len(preview) - 1)
    # col=0 후보인 경우 자기 자신을 skip
    if candidate_col == 0:
        start_r = r - 1
    for cur_r in range(start_r, -1, -1):
        if cur_r < 0:
            break
        row = preview[cur_r]
        if not row:
            continue
        val = row[0]
        if val:
            labels.append(val.strip())
            break
    return labels


def _candidate_matches_section(
    candidate: tuple[int, int, int, int, int, str],
    section_hint_norm: str,
    tables_by_idx: dict[int, Any],
) -> bool:
    """candidate 의 주변 섹션 라벨에 section_hint_norm 이 포함되면 매칭."""
    if not section_hint_norm:
        return True
    for ctx_label in _candidate_context_labels(candidate, tables_by_idx):
        if section_hint_norm in _normalize_label(ctx_label):
            return True
    return False


def _describe_cell_context(
    candidate: tuple[int, int, int, int, int, str],
    tables_by_idx: dict[int, Any],
) -> str:
    """LLM 에게 보여줄 candidate 셀의 사람 읽는 컨텍스트 문자열."""
    labels = _candidate_context_labels(candidate, tables_by_idx)
    if labels:
        return " / ".join(labels)
    t_idx = candidate[0]
    t = tables_by_idx.get(t_idx)
    loc = getattr(t, "location", None) if t else None
    return loc or f"table[{t_idx}]"


def _suggest_dot_path(
    candidate: tuple[int, int, int, int, int, str],
    tables_by_idx: dict[int, Any],
) -> str:
    """ambiguous hint 용 dot-path prefix 후보 (첫 후보의 섹션 라벨)."""
    labels = _candidate_context_labels(candidate, tables_by_idx)
    return labels[-1] if labels else "섹션"


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
    # 병합/누락 컬럼은 크기를 알 수 없어 해당 슬롯이 None 일 수 있다
    # (예: [1.5, None, 2.0]). 소비자는 per-element None 을 처리해야 한다.
    column_widths_cm: list[float | None] | None = None
    row_heights_cm: list[float | None] | None = None

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
class ShapeInfo:
    """PPTX 의 표 외 shape (textbox / placeholder / 도형 내 텍스트) 메타.

    ``shape_id`` 는 슬라이드 내 고유 숫자. ``name`` 은 사람 이름 (예: "Title 1").
    ``text`` 는 full text. ``text_preview`` 는 40자 컷.
    """
    slide_index: int
    shape_id: int
    name: str
    kind: str                  # shape 종류 (placeholder / text_box / group / picture 등)
    has_text: bool
    text: str = ""
    text_preview: str = ""
    placeholder_type: str | None = None

    def to_dict(self) -> dict[str, Any]:
        d = {
            "slide_index": self.slide_index,
            "shape_id": self.shape_id,
            "name": self.name,
            "kind": self.kind,
            "has_text": self.has_text,
            "text_preview": self.text_preview,
        }
        if self.text and self.text != self.text_preview:
            d["text"] = self.text
        if self.placeholder_type:
            d["placeholder_type"] = self.placeholder_type
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
    def save(self, path: Path | str | None = None) -> Path:
        """문서를 저장하고 경로를 반환.

        byte 안정성은 포맷별로 다르다: HWPX 는 수정 안 한 파트를 byte-identical
        하게 보존하지만, DOCX/PPTX 는 python-docx/pptx 가 패키지 전체를
        재작성한다(내용 보존, byte-identical 보장 없음). README 참조.
        """
        ...

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

    def get_form_fields(self) -> list[dict[str, Any]]:
        """표 셀 텍스트에서 라벨 후보와 **중복 여부(dot-path 필요)**를 미리 보여준다.

        fill_form 호출 *전에* LLM 이 "어떤 라벨이 여러 곳에 있어 dot-path 로 구분해야
        하는지" 를 알 수 있게 한다 — ambiguous 응답을 받고 재시도하는 라운드를 줄인다.

        각 항목: {label, normalized, ambiguous, occurrences:[{table_index,row,col}]}.
        (라벨로 보이는 비어있지 않은 anchor 셀 기준 — 빈 양식에서 가장 유용.)
        """
        index: dict[str, list[dict[str, Any]]] = {}
        labels: dict[str, str] = {}
        for t in self.get_tables(preview_rows=10_000, max_cell_len=200):
            for r, row in enumerate(t.preview):
                for c, val in enumerate(row):
                    if not val:
                        continue
                    norm = _normalize_label(val)
                    if not norm:
                        continue
                    index.setdefault(norm, []).append(
                        {"table_index": t.index, "row": r, "col": c})
                    labels.setdefault(norm, val)
        fields = [
            {
                "label": labels[norm],
                "normalized": norm,
                "ambiguous": len(occ) > 1,
                "occurrences": occ,
            }
            for norm, occ in index.items()
        ]
        return sorted(fields, key=lambda f: f["label"])

    # ---- editing ----
    @abstractmethod
    def render_template(self, context: dict[str, Any], *,
                        on_missing: str = "blank") -> dict[str, list[str]]:
        """템플릿의 {{key}}를 context 값으로 치환한다.

        on_missing: 템플릿에 있으나 context 에 없는 키의 처리 (3포맷 동일 정책):
          - "blank"(기본): 빈 문자열로 치환 — 출력에 미완성 플레이스홀더를 남기지 않음
          - "leave": ``{{key}}`` 를 그대로 둠
          - "error": 누락 키가 있으면 ValueError
        반환: ``{"used": [...], "missing": [...]}`` (템플릿에 등장한 키 기준).
        """

    def _render_report(self, placeholders: list[str], context: dict[str, Any],
                       on_missing: str) -> dict[str, list[str]]:
        """render_template 공통: on_missing 검증 + used/missing 분류.

        각 어댑터가 치환 *전에* 호출한다. on_missing='error' 이고 누락 키가
        있으면 여기서 ValueError 를 던져 문서를 건드리지 않는다.
        """
        if on_missing not in ("blank", "leave", "error"):
            raise ValueError(
                f"on_missing must be blank/leave/error, got {on_missing!r}")
        provided = set(context)
        used = [p for p in placeholders if p in provided]
        missing = [p for p in placeholders if p not in provided]
        if on_missing == "error" and missing:
            raise ValueError(f"missing placeholders: {missing}")
        return {"used": used, "missing": missing}

    @abstractmethod
    def set_cell(self, table_index: int, row: int, col: int, value: str,
                 *, allow_merge_redirect: bool = False) -> str:
        """셀 값 교체. 원래 값 반환.

        ``allow_merge_redirect`` 가 True 면 병합 영역의 non-anchor 좌표를
        anchor 셀로 redirect 해 쓴다. False(기본)면 MergedCellWriteError.
        (``fill_form`` 이 이 인자로 호출하므로 구현체는 반드시 받아들여야 한다.)
        """

    @abstractmethod
    def append_to_cell(self, table_index: int, row: int, col: int, value: str,
                       separator: str = "  ",
                       *, allow_merge_redirect: bool = False) -> str:
        """기존 셀 텍스트 뒤에 ``separator + value``를 덧붙임.

        라벨(예: "성  명")을 유지한 채 사용자 입력을 추가하는 용도.
        빈 셀이면 separator 없이 value만 기록. 원래 값 반환.
        ``allow_merge_redirect`` 의 의미는 :meth:`set_cell` 과 동일.
        """

    @abstractmethod
    def append_row(self, table_index: int, values: list[str]) -> None:
        """표 끝에 새 행 추가."""

    # ---- shape text (v0.8+, PPTX 전용) ----
    def get_shapes(
        self,
        slide_index: int | None = None,
        min_text_len: int = 1,
        max_preview: int = 40,
    ) -> list[ShapeInfo]:
        """표 외 shape (textbox / placeholder / 도형 텍스트) 목록.

        PPTX 에서만 실질적 의미가 있다. DOCX/HWPX 는 기본 빈 리스트 반환
        (해당 포맷은 표와 paragraph 위주로 편집).
        """
        return []

    def set_shape_text(
        self,
        slide_index: int,
        shape_id: int,
        text: str,
    ) -> str:
        """shape 의 텍스트 프레임을 text 로 교체. 기존 run-level 포맷 보존.

        Returns: 기존 텍스트.
        PPTX 만 지원 — 그 외 포맷은 NotImplementedForFormat.
        """
        raise NotImplementedForFormat(
            f"{self.format} does not support shape-level text editing"
        )

    # ---- form controls (체크박스/라디오/콤보/에디트 — HWPX 전용) ----
    def get_form_controls(self) -> list[dict[str, Any]]:
        """폼 컨트롤 목록. 표가 아닌 인터랙티브 필드(체크박스/에디트 등).

        HWPX 에서만 의미가 있다. 그 외 포맷은 빈 리스트.
        각 항목: {name, kind, caption, value, checked?}.
        """
        return []

    def set_form_control(self, name: str, value: Any) -> str:
        """이름으로 폼 컨트롤 값을 설정. 기존 값을 반환.

        - check/radio: value 가 truthy("Y"/"체크"/True/"checked") → CHECKED
        - edit/combo: 문자열을 값으로 기록
        HWPX 만 지원 — 그 외 포맷은 NotImplementedForFormat.
        """
        raise NotImplementedForFormat(
            f"{self.format} does not support form controls"
        )

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

        **Dot-path 로 ambiguous 해소**: 한 양식에 같은 라벨이 여러 번 등장하면
        (예: "금액" 이 피해자 섹션, 지급정지계좌 섹션, 피해금이전계좌 섹션 각각)
        ``"피해자.금액"`` 처럼 dot-path 로 section hint 를 지정하면 section 컨텍스트가
        일치하는 후보만 선택한다.

        Args:
            data: {라벨: 값} dict. 라벨은 셀 텍스트와 whitespace/특수문자 제거
                후 정규화 비교 (예: "성 명" == "성명"). "섹션힌트.라벨" 형태도 허용.
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
              "ambiguous": [{"label", "candidates": [{"table_index", "row", "col", "context"}, ...],
                             "hint": "dot-path 예시 (예: '피해자.금액')"}],
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
        tables_by_idx = {t.index: t for t in tables}
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
        overflow: list[dict[str, Any]] = []

        for label, value in data.items():
            section_hint, actual_label = _split_dot_path(label)
            key = _normalize_label(actual_label)
            all_candidates = label_index.get(key, [])

            # dot-path 가 있으면 section_hint 로 candidate 필터
            if section_hint and all_candidates:
                hint_norm = _normalize_label(section_hint)
                filtered = [
                    cand for cand in all_candidates
                    if _candidate_matches_section(cand, hint_norm, tables_by_idx)
                ]
                if filtered:
                    candidates = filtered
                else:
                    # hint 와 매칭 안 되면 원본 후보 유지 (ambiguous 또는 single)
                    candidates = all_candidates
            else:
                candidates = all_candidates

            if not candidates:
                if strict:
                    raise ValueError(f"label not found: {label!r}")
                not_found.append(label)
                continue
            if len(candidates) > 1:
                # 여러 곳 — 각 후보의 section context 수집해서 LLM 이 구분 가능하게
                ambiguous.append({
                    "label": label,
                    "candidates": [
                        {
                            "table_index": cand[0],
                            "row": cand[1],
                            "col": cand[2],
                            "context": _describe_cell_context(cand, tables_by_idx),
                        }
                        for cand in candidates
                    ],
                    "hint": (
                        f"dot-path 로 재호출 예시: "
                        f"'{_suggest_dot_path(candidates[0], tables_by_idx)}.{actual_label}'"
                    ),
                })
                continue

            t_idx, r, c, rs, cs, current_text = candidates[0]
            # auto 모드에서 인접 라벨 오염 방지 — 문서 내 **모든 anchor cell 텍스트**
            # 를 보호 대상으로. (값 셀이 포함되어 덮어쓰기 차단될 수 있지만 라벨 파괴가
            # 더 치명적이라 보수적 default. 예시값이 있는 양식에서 값 셀을 덮어쓰려면
            # direction="right"/"below" 로 명시.)
            other_label_keys = set(label_index.keys()) - {key}
            tinfo = tables_by_idx.get(t_idx)
            ncols = tinfo.cols if tinfo is not None else c + cs + 1
            nrows = tinfo.rows if tinfo is not None else r + rs + 1
            action, coord, old = self._fill_one_cell(
                t_idx, r, c, rs, cs, str(value), direction,
                other_label_keys, ncols, nrows
            )
            if coord is None:
                # 탐색 실패 — fallback 으로 same 에 append_to_cell 권할지 고민했지만
                # 의도치 않은 라벨 오염 방지 차원에서 그냥 not_found 에 기록.
                not_found.append(label)
                continue
            # 오버플로 위험 판정: 값이 들어간 칸 너비(width_cm) 대비 추정 글자폭.
            try:
                wcm = self.get_cell(coord[0], coord[1], coord[2]).width_cm
            except (IndexError, ValueError):
                wcm = None
            risk = _overflow_risk(str(value), wcm)
            entry = {
                "label": label,
                "table_index": coord[0],
                "row": coord[1],
                "col": coord[2],
                "action": action,
                "old_value": old,
                "new_value": str(value),
                "overflow_risk": risk,
            }
            if wcm is not None:
                entry["cell_width_cm"] = wcm
            filled.append(entry)
            if risk:
                overflow.append({"label": label, "table_index": coord[0],
                                 "row": coord[1], "col": coord[2],
                                 "cell_width_cm": wcm, "value": str(value)})

        return {"filled": filled, "not_found": not_found,
                "ambiguous": ambiguous, "overflow_warnings": overflow}

    def _scan_value_cell_right(
        self,
        t_idx: int,
        r: int,
        start_c: int,
        ncols: int,
    ) -> int | None:
        """라벨 오른쪽으로 다음 라벨 전까지 스캔해 **가장 넓은 빈 anchor 값칸**의
        열을 반환. 라벨과 값 영역 사이의 얇은 스페이서 칸(예: 0.4cm)에 값이 들어가
        세로로 깨지는 것을 방지한다 — 이미 추출하는 cell width(cm)를 활용. 없으면 None.

        non-anchor 슬롯은 건너뛰고, 비어있지 않은 anchor(=다음 라벨/값 영역)에서 멈춘다.
        스캔은 ``ncols`` 로 명시적으로 bound 한다 (get_cell 의 경계 예외에 의존하지 않음).
        """
        best: tuple[float, int] | None = None
        tc = start_c
        while tc < ncols:
            try:
                cell = self.get_cell(t_idx, r, tc)
            except (IndexError, ValueError):
                break
            if not cell.is_anchor:
                tc += 1
                continue
            if (cell.text or "").strip():
                break  # 다음 라벨/값 영역 시작 → 스캔 종료
            width = cell.width_cm if cell.width_cm is not None else 0.0
            if best is None or width > best[0]:
                best = (width, tc)
            tc += max(1, cell.span[1])
        return best[1] if best else None

    def _scan_value_cell_below(
        self,
        t_idx: int,
        c: int,
        start_r: int,
        nrows: int,
    ) -> int | None:
        """라벨 아래로 다음 라벨 전까지 스캔해 **가장 높은(height) 빈 anchor 값칸**의
        행을 반환. 라벨과 값 영역 사이의 얇은 스페이서 행에 값이 들어가는 것을
        방지한다 (right 스캔의 세로 대칭). non-anchor 슬롯은 건너뛰고, 비어있지
        않은 anchor 에서 멈춘다. nrows 로 bound. 없으면 None.
        """
        best: tuple[float, int] | None = None
        tr = start_r
        while tr < nrows:
            try:
                cell = self.get_cell(t_idx, tr, c)
            except (IndexError, ValueError):
                break
            if not cell.is_anchor:
                tr += 1
                continue
            if (cell.text or "").strip():
                break
            height = cell.height_cm if cell.height_cm is not None else 0.0
            if best is None or height > best[0]:
                best = (height, tr)
            tr += max(1, cell.span[0])
        return best[1] if best else None

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
        ncols: int,
        nrows: int,
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
                if direction == "auto":
                    # 얇은 스페이서 칸을 건너뛰고 가장 넓은 빈 값칸을 고른다.
                    tc = self._scan_value_cell_right(
                        t_idx, r, c + colspan, ncols)
                    if tc is None:
                        continue  # 오른쪽에 적합한 빈 값칸 없음 → below/same
                    target_r, target_c = r, tc
                else:
                    target_r, target_c = r, c + colspan
            elif mode == "below":
                if direction == "auto":
                    # 얇은 스페이서 행을 건너뛰고 적정한 빈 값칸을 고른다(right 대칭).
                    tr = self._scan_value_cell_below(
                        t_idx, c, r + rowspan, nrows)
                    if tr is None:
                        continue
                    target_r, target_c = tr, c
                else:
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
                old = _call_cell_writer(
                    self.append_to_cell,
                    t_idx, target_r, target_c, value,
                    allow_merge_redirect=not cell.is_anchor,
                )
                return "append_to_cell", (t_idx, target_r, target_c), old

            # right / below
            if direction == "auto" and not cell.is_anchor:
                # target 이 병합의 non-anchor 면 anchor 로 redirect 되어 엉뚱한 셀에
                # 쓰일 위험 (예: 스페이서 행의 병합 anchor). auto 에서는 skip 하고
                # 다음 mode 로.
                continue
            if direction == "auto" and target_key and target_key in other_label_keys:
                # 옆 셀이 다른 라벨 → 덮어쓰면 라벨 손상. 다음 mode 시도.
                continue
            try:
                old = _call_cell_writer(
                    self.set_cell,
                    t_idx, target_r, target_c, value,
                    allow_merge_redirect=(direction != "auto"),
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
