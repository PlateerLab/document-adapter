"""Smoke test — 세 포맷에 대해 템플릿 생성 → 렌더 → 검증.

외부 리소스 없이 pytest로 돌아간다:
    pytest tests/test_smoke.py -v
"""
from __future__ import annotations

from pathlib import Path

import pytest
from docx import Document
from docx.shared import Pt as DocxPt
from hwpx.document import HwpxDocument
from pptx import Presentation
from pptx.util import Inches, Pt as PptxPt

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


# ---- regression: set_cell must keep run-level font formatting (issue #1) ----

def _make_docx_table_with_font(path: Path) -> None:
    doc = Document()
    table = doc.add_table(rows=1, cols=1)
    cell = table.cell(0, 0)
    run = cell.paragraphs[0].add_run("original")
    run.font.name = "Malgun Gothic"
    run.font.size = DocxPt(18)
    run.bold = True
    doc.save(path)


def _make_pptx_table_with_font(path: Path) -> None:
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    table_shape = slide.shapes.add_table(1, 1, Inches(1), Inches(1), Inches(6), Inches(1))
    cell = table_shape.table.cell(0, 0)
    cell.text = "original"
    run = cell.text_frame.paragraphs[0].runs[0]
    run.font.name = "Malgun Gothic"
    run.font.size = PptxPt(18)
    run.font.bold = True
    prs.save(path)


def test_docx_set_cell_preserves_font(tmp_path: Path) -> None:
    src = tmp_path / "formatted.docx"
    _make_docx_table_with_font(src)

    doc = load(src)
    old = doc.set_cell(0, 0, 0, "replaced")
    doc.save()
    doc.close()
    assert old == "original"

    verify = Document(src)
    run = verify.tables[0].cell(0, 0).paragraphs[0].runs[0]
    assert run.text == "replaced"
    assert run.font.name == "Malgun Gothic"
    assert run.font.size == DocxPt(18)
    assert run.bold is True


def test_pptx_set_cell_preserves_font(tmp_path: Path) -> None:
    src = tmp_path / "formatted.pptx"
    _make_pptx_table_with_font(src)

    doc = load(src)
    old = doc.set_cell(0, 0, 0, "replaced")
    doc.save()
    doc.close()
    assert old == "original"

    verify = Presentation(src)
    cell = next(
        shape.table.cell(0, 0)
        for slide in verify.slides
        for shape in slide.shapes
        if shape.has_table
    )
    run = cell.text_frame.paragraphs[0].runs[0]
    assert run.text == "replaced"
    assert run.font.name == "Malgun Gothic"
    assert run.font.size == PptxPt(18)
    assert run.font.bold is True


def test_docx_append_row_preserves_formatting(tmp_path: Path) -> None:
    """add_row가 템플릿 행을 복사하더라도 새 run을 쓰지 않고 첫 run을 재활용해야 한다."""
    src = tmp_path / "append.docx"
    doc = Document()
    table = doc.add_table(rows=1, cols=2)
    for i in range(2):
        run = table.cell(0, i).paragraphs[0].add_run("hdr")
        run.font.name = "Malgun Gothic"
        run.font.size = DocxPt(14)
    doc.save(src)

    adapter = load(src)
    adapter.append_row(0, ["A", "B"])
    adapter.save()
    adapter.close()

    verify = Document(src)
    new_row = verify.tables[0].rows[1]
    for col, expected in enumerate(["A", "B"]):
        run = new_row.cells[col].paragraphs[0].runs[0]
        assert run.text == expected
        # 복사된 행은 템플릿 속성을 이어받으므로 font가 비어 있지 않아야 한다
        assert run.font.name in {"Malgun Gothic", None}
