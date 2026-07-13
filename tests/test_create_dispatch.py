"""create_document 디스패처 — 확장자-인자 매트릭스 / 가드 / call_tool 경유."""
from __future__ import annotations

from pathlib import Path

import pytest

from document_adapter import create_document
from document_adapter.tools import call_tool

MD = "# 제목\n본문"
SHEETS = [{"name": "S", "headers": ["a"], "rows": [["1"]]}]


# -------- 확장자-인자 매트릭스 --------

def test_docx_requires_markdown(tmp_path: Path) -> None:
    with pytest.raises(ValueError) as exc:
        create_document(tmp_path / "a.docx", sheets=SHEETS)
    assert "markdown" in str(exc.value)


def test_xlsx_requires_sheets(tmp_path: Path) -> None:
    with pytest.raises(ValueError) as exc:
        create_document(tmp_path / "a.xlsx", markdown=MD)
    assert "sheets" in str(exc.value)


def test_both_args_rejected(tmp_path: Path) -> None:
    with pytest.raises(ValueError) as exc:
        create_document(tmp_path / "a.docx", markdown=MD, sheets=SHEETS)
    assert "동시에" in str(exc.value)


def test_hwpx_from_markdown(tmp_path: Path) -> None:
    # v0.16 부터 hwpx 도 markdown 으로 생성된다 — 생성 후 load() 왕복 성립.
    from document_adapter import load

    out = create_document(tmp_path / "a.hwpx", markdown=MD)
    assert out.exists() and out.stat().st_size > 0
    assert type(load(out)).__name__ == "HwpxAdapter"


def test_pptx_from_markdown(tmp_path: Path) -> None:
    # .pptx 는 `---` / 레벨 1~2 헤딩으로 슬라이드를 나눈다.
    from document_adapter import load

    out = create_document(tmp_path / "a.pptx", markdown="# 표지\n---\n# 슬라이드2\n- 불릿")
    assert out.exists() and out.stat().st_size > 0
    assert type(load(out)).__name__ == "PptxAdapter"


def test_unsupported_extension(tmp_path: Path) -> None:
    with pytest.raises(ValueError) as exc:
        create_document(tmp_path / "a.pdf", markdown=MD)
    assert ".pdf" in str(exc.value)


# -------- 파일 가드 --------

def test_overwrite_guard(tmp_path: Path) -> None:
    out = create_document(tmp_path / "a.docx", markdown=MD)
    with pytest.raises(ValueError) as exc:
        create_document(out, markdown=MD)
    assert "존재" in str(exc.value)
    # overwrite=True 는 통과
    create_document(out, markdown="# 새 문서", overwrite=True)


def test_parent_dir_auto_created(tmp_path: Path) -> None:
    out = create_document(tmp_path / "하위" / "폴더" / "a.docx", markdown=MD)
    assert out.exists()


# -------- call_tool 경유 (MCP/Claude 도구 표면) --------

def test_call_tool_create_document(tmp_path: Path) -> None:
    result = call_tool("create_document", {
        "path": str(tmp_path / "회의록.docx"),
        "markdown": "# 회의록\n\n| a | b |\n|---|---|\n| 1 | 2 |\n",
    })
    assert "error" not in result
    assert result["format"] == "docx"
    assert result["tables"] == 1
    assert result["table_shapes"] == ["2x2"]
    assert Path(result["output_path"]).exists()


def test_call_tool_error_serialized(tmp_path: Path) -> None:
    result = call_tool("create_document", {
        "path": str(tmp_path / "a.xlsx"),
        "sheets": [{"headers": ["a"], "rows": []}],  # name 누락
    })
    assert result["error"] == "ValueError"
    assert "name" in result["message"]
