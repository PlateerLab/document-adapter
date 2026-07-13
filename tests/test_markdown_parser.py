"""generate.markdown_parser — 블록 IR 파서 단위 테스트."""
from __future__ import annotations

from document_adapter.generate.markdown_parser import (
    Block,
    Span,
    parse_inline,
    parse_markdown,
)


def kinds(blocks: list[Block]) -> list[str]:
    return [b.kind for b in blocks]


def text_of(spans) -> str:
    return "".join(s.text for s in spans)


# -------- 블록 종류별 --------

def test_basic_blocks() -> None:
    md = (
        "# 제목\n"
        "\n"
        "본문 문단입니다.\n"
        "\n"
        "## 섹션\n"
        "- 불릿 하나\n"
        "* 불릿 둘\n"
        "1. 첫째\n"
        "2) 둘째\n"
        "> 인용문\n"
        "---\n"
    )
    blocks = parse_markdown(md)
    assert kinds(blocks) == [
        "heading", "paragraph", "heading",
        "bullet", "bullet", "numbered", "numbered", "quote", "hr",
    ]
    assert blocks[0].level == 1
    assert blocks[2].level == 2
    assert text_of(blocks[3].spans) == "불릿 하나"
    assert text_of(blocks[7].spans) == "인용문"


def test_heading_level_capped_at_six() -> None:
    blocks = parse_markdown("###### 여섯\n")
    assert blocks[0].kind == "heading" and blocks[0].level == 6
    # 7개 이상 # 은 헤딩 정규식에 안 걸려 문단으로 관용 처리
    blocks = parse_markdown("####### 일곱\n")
    assert blocks[0].kind == "paragraph"


def test_crlf_normalization() -> None:
    blocks = parse_markdown("# A\r\n\r\n본문\r\n")
    assert kinds(blocks) == ["heading", "paragraph"]


# -------- 인라인 --------

def test_inline_spans() -> None:
    spans = parse_inline("앞 **굵게** 중간 *기울임* 그리고 `code` 끝")
    assert [(s.text, s.bold, s.italic, s.code) for s in spans] == [
        ("앞 ", False, False, False),
        ("굵게", True, False, False),
        (" 중간 ", False, False, False),
        ("기울임", False, True, False),
        (" 그리고 ", False, False, False),
        ("code", False, False, True),
        (" 끝", False, False, False),
    ]


def test_inline_plain() -> None:
    spans = parse_inline("마크업 없음")
    assert spans == (Span("마크업 없음"),)


def test_inline_in_heading_and_bullet() -> None:
    blocks = parse_markdown("# **강조** 제목\n- *이탤릭* 항목\n")
    assert blocks[0].spans[0].bold is True
    assert blocks[1].spans[0].italic is True


# -------- 표 --------

def test_table_with_separator() -> None:
    md = (
        "| 이름 | 부서 |\n"
        "|---|---|\n"
        "| 김경윤 | AI플랫폼 |\n"
        "| 홍길동 | 개발1팀 |\n"
    )
    blocks = parse_markdown(md)
    assert kinds(blocks) == ["table"]
    t = blocks[0]
    assert len(t.rows) == 3  # 헤더 + 2
    assert text_of(t.rows[0][0]) == "이름"
    assert text_of(t.rows[2][1]) == "개발1팀"


def test_table_dash_data_row_is_not_separator() -> None:
    """`| - | - |` 는 구분행이 아니라 실데이터 — 셀 전부 `--` 이상일 때만 구분행."""
    md = (
        "| a | b |\n"
        "|---|---|\n"
        "| - | - |\n"
    )
    blocks = parse_markdown(md)
    assert blocks[0].kind == "table"
    assert len(blocks[0].rows) == 2  # 헤더 + 데이터('-','-')
    assert text_of(blocks[0].rows[1][0]) == "-"


def test_table_without_separator_is_not_table() -> None:
    md = "| a | b |\n| c | d |\n"
    blocks = parse_markdown(md)
    assert "table" not in kinds(blocks)
    assert kinds(blocks) == ["paragraph", "paragraph"]


def test_table_ragged_rows_normalized() -> None:
    md = (
        "| a | b | c |\n"
        "|---|---|---|\n"
        "| 1 |\n"
        "| 1 | 2 | 3 | 4 |\n"
    )
    t = parse_markdown(md)[0]
    assert all(len(row) == 3 for row in t.rows)
    assert text_of(t.rows[1][1]) == ""      # 부족분 빈칸
    assert text_of(t.rows[2][2]) == "3"     # 초과분 버림


def test_table_alignment_separator() -> None:
    md = "| a | b |\n|:---|---:|\n| 1 | 2 |\n"
    blocks = parse_markdown(md)
    assert blocks[0].kind == "table"
    assert len(blocks[0].rows) == 2


# -------- 코드펜스 --------

def test_code_fence() -> None:
    md = "```\nline1\n  line2\n```\n뒤 문단\n"
    blocks = parse_markdown(md)
    assert kinds(blocks) == ["code", "paragraph"]
    assert blocks[0].lines == ("line1", "  line2")


def test_code_fence_unclosed_runs_to_eof() -> None:
    blocks = parse_markdown("```\nonly\n")
    assert kinds(blocks) == ["code"]
    assert blocks[0].lines == ("only",)


# -------- 관용 처리 (지원 외 문법) --------

def test_unsupported_syntax_degrades_to_paragraph() -> None:
    md = "<div>html</div>\n\n![img](x.png)\n\n[^1]: 각주\n"
    blocks = parse_markdown(md)
    assert kinds(blocks) == ["paragraph"] * 3  # 에러 없이 텍스트로


def test_nested_bullet_flattens() -> None:
    """중첩 리스트는 지원 외 — 들여쓰기가 무시되고 평탄한 불릿으로."""
    blocks = parse_markdown("- 상위\n  - 하위\n")
    assert kinds(blocks) == ["bullet", "bullet"]
