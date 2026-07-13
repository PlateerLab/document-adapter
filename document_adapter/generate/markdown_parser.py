"""제약된 markdown 서브셋 → 블록 IR(중간표현) 파서.

문서 생성(v0.15+)의 공용 프론트엔드. 렌더러(docx_writer, 이후 hwpx_writer)가
이 IR 을 소비하므로 포맷별로 markdown 파싱을 중복 구현하지 않는다 —
edit2docs(참고 구현)는 파서가 docx 엔진에 인라인이라 재사용이 불가능했다.

지원 서브셋 (LLM 도구 설명과 1:1 로 일치해야 한다):
    # ~ ###### 헤딩 · 일반 문단 · - / * 불릿 · 1. 번호 목록 ·
    **굵게** / *기울임* / `코드` 인라인 · 파이프 표 (|---| 구분행 필수) ·
    > 인용 · --- 수평선 · ``` 코드펜스

지원 외 문법(HTML/이미지/각주/중첩 리스트/셀 병합)은 **에러 없이 일반
텍스트로 관용 처리**한다 — LLM 이 서브셋을 살짝 벗어나도 문서가 밋밋해질
뿐 생성이 실패하지 않는다. 실패는 렌더러(빈 본문 등)의 몫.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field

__all__ = ["Span", "Block", "parse_markdown", "parse_inline"]


@dataclass(frozen=True)
class Span:
    """인라인 run 단위 — 렌더러가 run/charPr 로 1:1 매핑한다."""

    text: str
    bold: bool = False
    italic: bool = False
    code: bool = False


@dataclass(frozen=True)
class Block:
    """블록 단위 IR. kind 에 따라 사용하는 필드가 다르다.

    kind:
        heading   — level(1..6) + spans
        paragraph — spans
        bullet    — spans (줄당 Block 1개)
        numbered  — spans (줄당 Block 1개)
        quote     — spans
        table     — rows (rows[0] = 헤더행), 각 셀은 Span 튜플
        hr        — (필드 없음)
        code      — lines (펜스 안 원문 그대로)
    """

    kind: str
    level: int = 0
    spans: tuple[Span, ...] = ()
    rows: tuple[tuple[tuple[Span, ...], ...], ...] = field(default=())
    lines: tuple[str, ...] = ()


# **bold** / *italic* / `code` — edit2docs 검증 정규식과 동일 계열.
# bold 를 italic 보다 먼저 매칭해야 "**x**" 가 "*(*x)*" 로 쪼개지지 않는다.
_INLINE = re.compile(
    r"(\*\*(?P<bold>.+?)\*\*)|(\*(?P<italic>[^*]+?)\*)|(`(?P<code>[^`]+?)`)"
)

_HEADING = re.compile(r"^(?P<hashes>#{1,6})\s+(?P<text>.+)$")
_BULLET = re.compile(r"^[-*]\s+(?P<text>.+)$")
_NUMBERED = re.compile(r"^\d+[.)]\s+(?P<text>.+)$")
_HR = re.compile(r"^(-{3,}|\*{3,}|_{3,})$")
# 구분행: 모든 셀이 :?-{2,}:? 일 때만. 느슨한 판정은 "| - | - |" 같은
# 실데이터 행을 삼킨다 (edit2docs 에서 실제로 발생했던 함정).
_SEPARATOR_CELL = re.compile(r":?-{2,}:?")


def parse_inline(text: str) -> tuple[Span, ...]:
    """인라인 마크업을 Span 튜플로. 마크업이 없으면 단일 plain Span."""
    spans: list[Span] = []
    pos = 0
    for m in _INLINE.finditer(text):
        if m.start() > pos:
            spans.append(Span(text[pos : m.start()]))
        if m.group("bold") is not None:
            spans.append(Span(m.group("bold"), bold=True))
        elif m.group("italic") is not None:
            spans.append(Span(m.group("italic"), italic=True))
        else:
            spans.append(Span(m.group("code"), code=True))
        pos = m.end()
    if pos < len(text):
        spans.append(Span(text[pos:]))
    return tuple(spans)


def _is_table_row(line: str) -> bool:
    s = line.strip()
    return s.startswith("|") and s.endswith("|") and s.count("|") >= 2


def _split_table_row(line: str) -> list[str]:
    return [cell.strip() for cell in line.strip().strip("|").split("|")]


def _is_separator_row(line: str) -> bool:
    cells = _split_table_row(line)
    return bool(cells) and all(_SEPARATOR_CELL.fullmatch(c) for c in cells)


def _parse_table(lines: list[str], start: int) -> tuple[Block | None, int]:
    """start 위치의 표 후보를 파싱. (Block|None, 소비한 다음 인덱스).

    2번째 줄이 구분행이 아니면 표가 아니다 → None 반환, 호출측이 일반
    문단으로 처리한다.
    """
    if start + 1 >= len(lines) or not _is_table_row(lines[start + 1]):
        return None, start
    if not _is_separator_row(lines[start + 1]):
        return None, start

    headers = _split_table_row(lines[start])
    n_cols = len(headers)
    rows: list[tuple[tuple[Span, ...], ...]] = [
        tuple(parse_inline(h) for h in headers)
    ]
    i = start + 2
    while i < len(lines) and _is_table_row(lines[i]):
        if _is_separator_row(lines[i]):
            i += 1
            continue
        cells = _split_table_row(lines[i])
        # 헤더 기준 정규화: 부족분 빈 셀, 초과분 버림 (관용)
        cells = (cells + [""] * n_cols)[:n_cols]
        rows.append(tuple(parse_inline(c) for c in cells))
        i += 1
    return Block(kind="table", rows=tuple(rows)), i


def parse_markdown(md: str) -> list[Block]:
    """markdown 서브셋 → Block 리스트. 지원 외 문법은 문단으로 관용 처리."""
    lines = md.replace("\r\n", "\n").split("\n")
    blocks: list[Block] = []
    i = 0
    while i < len(lines):
        stripped = lines[i].strip()

        if not stripped:
            i += 1
            continue

        # 코드펜스 (닫는 펜스 없으면 EOF 까지)
        if stripped.startswith("```"):
            i += 1
            code_lines: list[str] = []
            closed = False
            while i < len(lines):
                if lines[i].strip().startswith("```"):
                    closed = True
                    break
                code_lines.append(lines[i])
                i += 1
            i += 1  # 닫는 펜스 (또는 EOF 초과 — 무해)
            if not closed:
                # 미폐쇄 펜스: 원문 끝의 개행이 만든 trailing 빈 줄 제거
                while code_lines and not code_lines[-1].strip():
                    code_lines.pop()
            blocks.append(Block(kind="code", lines=tuple(code_lines or [""])))
            continue

        # 표 (구분행 확인까지 통과해야 표)
        if _is_table_row(stripped):
            table, next_i = _parse_table(lines, i)
            if table is not None:
                blocks.append(table)
                i = next_i
                continue
            # 구분행 없음 → 표 아님, 아래 일반 규칙으로 흘려보냄

        m = _HEADING.match(stripped)
        if m:
            blocks.append(
                Block(
                    kind="heading",
                    level=len(m.group("hashes")),
                    spans=parse_inline(m.group("text").strip()),
                )
            )
            i += 1
            continue

        if _HR.match(stripped):
            blocks.append(Block(kind="hr"))
            i += 1
            continue

        m = _BULLET.match(stripped)
        if m:
            blocks.append(Block(kind="bullet", spans=parse_inline(m.group("text"))))
            i += 1
            continue

        m = _NUMBERED.match(stripped)
        if m:
            blocks.append(Block(kind="numbered", spans=parse_inline(m.group("text"))))
            i += 1
            continue

        if stripped.startswith(">"):
            blocks.append(
                Block(kind="quote", spans=parse_inline(stripped.lstrip("> ").strip()))
            )
            i += 1
            continue

        blocks.append(Block(kind="paragraph", spans=parse_inline(stripped)))
        i += 1

    return blocks
