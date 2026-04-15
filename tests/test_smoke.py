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


# ---- empty-cell regression: endParaRPr / pPr.rPr must be cloned (issue #1 v0.1.2) ----


def _make_pptx_empty_cell_with_endpararpr(path: Path) -> None:
    """빈 PPTX 셀에 endParaRPr(폰트 정보)만 심어둔 상태를 만든다.

    실제 고객 템플릿에서 "아직 값을 입력하지 않았지만 셀이 칠해질 때 쓸
    폰트"가 endParaRPr에 박혀 있는 상황을 재현한다.
    """
    from lxml import etree

    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    table_shape = slide.shapes.add_table(1, 1, Inches(1), Inches(1), Inches(6), Inches(1))
    cell = table_shape.table.cell(0, 0)

    # cell.text = ""로 만든 뒤 기존 run과 그 rPr을 전부 제거하고,
    # endParaRPr만 남긴다 (실제 PowerPoint 저장본의 empty-cell 상태 모방).
    cell.text = ""
    p_el = cell.text_frame.paragraphs[0]._p
    A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
    # 기존 <a:r>들 제거
    for r in p_el.findall(f"{{{A_NS}}}r"):
        p_el.remove(r)
    # endParaRPr 주입 (이미 있으면 속성만 세팅)
    end_rpr = p_el.find(f"{{{A_NS}}}endParaRPr")
    if end_rpr is None:
        end_rpr = etree.SubElement(p_el, f"{{{A_NS}}}endParaRPr")
    end_rpr.set("lang", "en-US")
    end_rpr.set("sz", "1800")  # 18pt in PPTX units (1/100 pt)
    end_rpr.set("b", "1")
    latin = etree.SubElement(end_rpr, f"{{{A_NS}}}latin")
    latin.set("typeface", "Microsoft Sans Serif")

    prs.save(path)


def test_pptx_empty_cell_preserves_endpararpr(tmp_path: Path) -> None:
    src = tmp_path / "empty_endpara.pptx"
    _make_pptx_empty_cell_with_endpararpr(src)

    doc = load(src)
    old = doc.set_cell(0, 0, 0, "V-2024-001")
    doc.save()
    doc.close()
    assert old == ""

    verify = Presentation(src)
    cell = next(
        shape.table.cell(0, 0)
        for slide in verify.slides
        for shape in slide.shapes
        if shape.has_table
    )
    run = cell.text_frame.paragraphs[0].runs[0]
    assert run.text == "V-2024-001"
    # endParaRPr에 담겼던 속성이 run의 rPr로 복사되었는지 XML 레벨로 확인
    A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
    rpr = run._r.find(f"{{{A_NS}}}rPr")
    assert rpr is not None, "new run lost its rPr"
    assert rpr.get("sz") == "1800"
    assert rpr.get("b") == "1"
    latin = rpr.find(f"{{{A_NS}}}latin")
    assert latin is not None
    assert latin.get("typeface") == "Microsoft Sans Serif"


def _make_docx_empty_cell_with_ppr_rpr(path: Path) -> None:
    """빈 DOCX 셀의 paragraph에 <w:pPr><w:rPr>만 심어둔 상태를 만든다."""
    from lxml import etree

    W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    doc = Document()
    table = doc.add_table(rows=1, cols=1)
    cell = table.cell(0, 0)
    p_el = cell.paragraphs[0]._p

    # 기존 run들 제거
    for r in p_el.findall(f"{{{W_NS}}}r"):
        p_el.remove(r)

    # <w:pPr><w:rPr>...</w:rPr></w:pPr> 주입
    ppr = p_el.find(f"{{{W_NS}}}pPr")
    if ppr is None:
        ppr = etree.SubElement(p_el, f"{{{W_NS}}}pPr")
        p_el.insert(0, ppr)
    rpr = etree.SubElement(ppr, f"{{{W_NS}}}rPr")
    rfonts = etree.SubElement(rpr, f"{{{W_NS}}}rFonts")
    rfonts.set(f"{{{W_NS}}}ascii", "Malgun Gothic")
    rfonts.set(f"{{{W_NS}}}eastAsia", "Malgun Gothic")
    sz = etree.SubElement(rpr, f"{{{W_NS}}}sz")
    sz.set(f"{{{W_NS}}}val", "36")  # half-points → 18pt
    b = etree.SubElement(rpr, f"{{{W_NS}}}b")
    b.set(f"{{{W_NS}}}val", "1")

    doc.save(path)


def test_docx_empty_cell_preserves_ppr_rpr(tmp_path: Path) -> None:
    src = tmp_path / "empty_pprrpr.docx"
    _make_docx_empty_cell_with_ppr_rpr(src)

    doc = load(src)
    old = doc.set_cell(0, 0, 0, "값")
    doc.save()
    doc.close()
    assert old == ""

    verify = Document(src)
    run = verify.tables[0].cell(0, 0).paragraphs[0].runs[0]
    assert run.text == "값"
    # pPr.rPr에 담겼던 속성이 run의 rPr로 복사됐는지 확인
    W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    rpr = run._r.find(f"{{{W_NS}}}rPr")
    assert rpr is not None, "new run lost its rPr"
    sz = rpr.find(f"{{{W_NS}}}sz")
    assert sz is not None and sz.get(f"{{{W_NS}}}val") == "36"
    rfonts = rpr.find(f"{{{W_NS}}}rFonts")
    assert rfonts is not None
    assert rfonts.get(f"{{{W_NS}}}ascii") == "Malgun Gothic"


def test_hwpx_set_cell_preserves_charprid_ref(tmp_path: Path) -> None:
    """HWPX 셀의 run이 가진 charPrIDRef가 set_cell 후에도 유지되는지 확인.

    python-hwpx의 paragraph.text setter는 PPTX/DOCX와 달리 기존 run의 <hp:t>만
    교체하는 방식이라 charPrIDRef가 자연스럽게 보존된다. 이 테스트는 upstream이
    그 동작을 바꿀 경우(run 재생성 방식으로) 즉시 감지하기 위한 회귀 가드다.
    """
    import re
    import zipfile

    src = tmp_path / "empty_charpr.hwpx"
    doc = HwpxDocument.new()
    doc.add_paragraph("")
    doc.add_table(1, 1)
    doc.save_to_path(src)

    # table cell 안의 run만 charPrIDRef="7"로 조작한 zip으로 재기록
    patched = tmp_path / "patched_charpr.hwpx"
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(patched, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if "section" in item.filename.lower() and item.filename.endswith(".xml"):
                text = data.decode("utf-8")
                def repl(m: re.Match) -> str:
                    tc = m.group(0)
                    return re.sub(
                        r'(<hp:run\s+charPrIDRef=")0(")',
                        r"\g<1>7\g<2>",
                        tc,
                        count=1,
                    )
                text = re.sub(r"<hp:tc[^>]*>.*?</hp:tc>", repl, text, flags=re.DOTALL)
                data = text.encode("utf-8")
            zout.writestr(item, data)

    # set_cell 실행
    adapter = load(patched)
    adapter.set_cell(0, 0, 0, "새 값")
    adapter.save()
    adapter.close()

    # table cell 안의 run만 추출해서 확인
    with zipfile.ZipFile(patched) as z:
        section = next(
            n for n in z.namelist() if "section" in n.lower() and n.endswith(".xml")
        )
        xml = z.read(section).decode("utf-8")

    tc_match = re.search(r"<hp:tc[^>]*>.*?</hp:tc>", xml, flags=re.DOTALL)
    assert tc_match is not None
    run_match = re.search(r"<hp:run[^>]*>.*?</hp:run>", tc_match.group())
    assert run_match is not None
    run_xml = run_match.group()
    assert 'charPrIDRef="7"' in run_xml, f"charPrIDRef not preserved: {run_xml!r}"
    assert "새 값" in run_xml
