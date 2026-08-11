"""문서 생성 (v0.15+) — 확장자 디스패치 진입점.

    from document_adapter import create_document

    create_document("회의록.docx", markdown="# 주간 회의록\\n...", lang="ko")
    create_document("보고.pptx", markdown="# 표지\\n---\\n# 슬라이드2\\n- 불릿")
    create_document("공문.hwpx", markdown="# 제목\\n본문...")
    create_document("매출.xlsx", sheets=[{"name": "...", "headers": [...], "rows": [...]}])

.xlsx sheet spec 은 시트당 세 가지 역할을 조합할 수 있다 (v0.17+):
    {"name": "데이터",   "headers": [...], "rows": [...]}      # 표
    {"name": "차트",     "charts": [{...}]}                     # 다른 시트 참조 차트
    {"name": "보고서",   "markdown": "# 제목\\n본문"}            # markdown → 셀+서식
자세한 스키마는 ``xlsx_writer`` 모듈 docstring 참조.

설계 원칙:
- LLM 산출물은 markdown / sheet spec(dict) 뿐 — XML 을 직접 쓰는 경로 없음.
- 렌더러가 곧 검증기: 잘못된 산출물은 이중어 ValueError. 재시도는 호출
  레이어(도구/에이전트)의 몫이다 (1회 재시도 계약).
- 생성-편집 왕복 보장: 반환 즉시 ``load(path)`` 가 성립해 기존 편집 도구
  (inspect_document → set_cell / insert_row / ...)로 이어진다.
- markdown 렌더러 3종(.docx/.pptx/.hwpx)은 같은 파서(markdown_parser)를
  공유한다 — .pptx 는 ``---``/레벨 1~2 헤딩으로 슬라이드를 나눈다.
"""
from __future__ import annotations

from pathlib import Path
from typing import Any

from .docx_writer import base_font_for_lang, docx_from_markdown
from .hwpx_writer import hwpx_from_markdown
from .markdown_parser import Block, Span, parse_inline, parse_markdown
from .pptx_writer import pptx_from_markdown
from .xlsx_writer import xlsx_from_sheets

__all__ = [
    "create_document",
    "docx_from_markdown",
    "pptx_from_markdown",
    "hwpx_from_markdown",
    "xlsx_from_sheets",
    "parse_markdown",
    "parse_inline",
    "base_font_for_lang",
    "Block",
    "Span",
]

_MARKDOWN_EXTS = {".docx", ".pptx", ".hwpx"}
_SHEETS_EXTS = {".xlsx"}
_SUPPORTED = sorted(_MARKDOWN_EXTS | _SHEETS_EXTS)

_MARKDOWN_RENDERERS = {
    ".docx": docx_from_markdown,
    ".pptx": pptx_from_markdown,
    ".hwpx": hwpx_from_markdown,
}


def create_document(
    path: str | Path,
    *,
    markdown: str | None = None,
    sheets: list[dict[str, Any]] | None = None,
    lang: str = "ko",
    overwrite: bool = False,
) -> Path:
    """확장자에 맞는 렌더러로 새 문서를 생성해 path 에 저장한다.

    Args:
        path: 출력 경로. 확장자가 렌더러를 고른다 (.docx/.pptx/.hwpx=markdown,
            .xlsx=sheets). 부모 디렉토리는 자동 생성.
        markdown: .docx/.pptx/.hwpx 용 본문 (제약된 markdown 서브셋 —
            .pptx 는 ``---`` 또는 레벨 1~2 헤딩이 슬라이드 구분).
        sheets: .xlsx 용 sheet spec 리스트 (xlsx_writer 스키마).
        lang: 기본 폰트 스택 선택 (ko → 맑은 고딕).
        overwrite: False(기본)면 기존 파일이 있을 때 ValueError.

    Returns:
        생성된 파일 경로. 반환 즉시 ``document_adapter.load()`` 가능.

    Raises:
        ValueError: 확장자-인자 불일치 / 스펙 검증 실패 / 파일 존재.
            전부 이중어 메시지 — 호출 레이어의 재시도 계약용.
    """
    out = Path(path)
    suffix = out.suffix.lower()

    if suffix not in _MARKDOWN_EXTS | _SHEETS_EXTS:
        raise ValueError(
            f"unsupported extension {suffix!r} for create_document "
            f"(supported: {', '.join(_SUPPORTED)}). "
            f"create_document 가 지원하지 않는 확장자입니다: {suffix!r} "
            f"(지원: {', '.join(_SUPPORTED)})."
        )

    if markdown is not None and sheets is not None:
        raise ValueError(
            "pass either `markdown` or `sheets`, not both. "
            "`markdown` 과 `sheets` 는 동시에 줄 수 없습니다."
        )

    if suffix in _MARKDOWN_EXTS:
        if markdown is None:
            raise ValueError(
                f"{suffix} generation requires `markdown`. "
                f"{suffix} 생성에는 `markdown` 인자가 필요합니다."
            )
        content = _MARKDOWN_RENDERERS[suffix](markdown, lang=lang)
    else:  # _SHEETS_EXTS
        if sheets is None:
            raise ValueError(
                f"{suffix} generation requires `sheets` (sheet spec list). "
                f"{suffix} 생성에는 `sheets`(sheet spec 리스트) 인자가 필요합니다."
            )
        content = xlsx_from_sheets(sheets)

    if out.exists() and not overwrite:
        raise ValueError(
            f"file already exists: {out} — pass overwrite=True to replace. "
            f"파일이 이미 존재합니다: {out} — 덮어쓰려면 overwrite=True 를 전달하세요."
        )

    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_bytes(content)
    return out
