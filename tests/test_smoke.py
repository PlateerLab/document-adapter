"""Smoke test — 세 포맷에 대해 템플릿 생성 → 렌더 → 검증.

외부 리소스 없이 pytest로 돌아간다:
    pytest tests/test_smoke.py -v
"""
from __future__ import annotations

from pathlib import Path

import pytest
from docx import Document
from hwpx.document import HwpxDocument
from pptx import Presentation
from pptx.util import Inches

from document_adapter import load
from document_adapter.tools import call_tool


# -------- 헬퍼: 공정한 템플릿을 그때그때 생성 --------

def _make_docx(path: Path) -> None:
    doc = Document()
    doc.add_paragraph("{{ title }}")
    doc.add_paragraph("작성자: {{ author }}")
    doc.save(path)


def _make_pptx(path: Path) -> None:
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "{{ title }}"
    slide.placeholders[1].text = "작성자: {{ author }}"
    prs.save(path)


def _make_hwpx(path: Path) -> None:
    doc = HwpxDocument.new()
    doc.add_paragraph("{{ title }}")
    doc.add_paragraph("작성자: {{ author }}")
    doc.save_to_path(path)


FACTORIES = {
    "docx": _make_docx,
    "pptx": _make_pptx,
    "hwpx": _make_hwpx,
}

CONTEXT = {"title": "통합 테스트", "author": "tester"}


@pytest.mark.parametrize("fmt", ["docx", "pptx", "hwpx"])
def test_render_template_dispatcher(tmp_path: Path, fmt: str) -> None:
    src = tmp_path / f"template.{fmt}"
    FACTORIES[fmt](src)

    doc = load(src)
    assert set(doc.get_placeholders()) == {"title", "author"}
    doc.render_template(CONTEXT)
    doc.save()
    doc.close()

    doc2 = load(src)
    assert doc2.get_placeholders() == []
    doc2.close()


@pytest.mark.parametrize("fmt", ["docx", "pptx", "hwpx"])
def test_tool_inspect(tmp_path: Path, fmt: str) -> None:
    src = tmp_path / f"template.{fmt}"
    FACTORIES[fmt](src)

    result = call_tool("inspect_document", {"path": str(src)})
    assert result["format"] == fmt
    assert "title" in result["placeholders"]
    assert "author" in result["placeholders"]


def test_tool_render(tmp_path: Path) -> None:
    src = tmp_path / "template.docx"
    _make_docx(src)

    result = call_tool("render_template", {
        "path": str(src),
        "context": CONTEXT,
    })
    assert result["rendered_count"] == 2
    assert result["placeholders_after"] == []
    assert Path(result["output_path"]).exists()


def test_tool_append_row_not_implemented_pptx(tmp_path: Path) -> None:
    src = tmp_path / "template.pptx"
    _make_pptx(src)

    result = call_tool("append_row", {
        "path": str(src),
        "table_index": 0,
        "values": ["x"],
    })
    assert result["error"] == "not_implemented"
