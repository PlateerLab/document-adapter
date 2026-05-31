"""시나리오 기반 회귀/보안 테스트.

test_smoke.py 가 "정상 동작"을 검증한다면, 이 파일은 발굴한 엣지케이스·보안
시나리오를 실패 테스트로 고정하고 개선을 추적한다.

    pytest tests/test_scenarios.py -v
"""
from __future__ import annotations

import zipfile
from pathlib import Path

import pytest
from lxml import etree
from hwpx.document import HwpxDocument

from document_adapter import load
from document_adapter.hwpx_core.package import HwpxPackage


# ---------------------------------------------------------------------------
# 공용 헬퍼
# ---------------------------------------------------------------------------

def _make_hwpx(path: Path, paragraphs: list[str]) -> None:
    doc = HwpxDocument.new()
    for p in paragraphs:
        doc.add_paragraph(p)
    doc.save_to_path(path)


def _poison_section0(src: Path, dst: Path, *, doctype: bytes, marker: bytes) -> None:
    """src.hwpx 의 Contents/section0.xml 에 DOCTYPE 엔티티를 주입한다.

    - doctype: <!DOCTYPE ...> 선언 (엔티티 정의 포함)
    - marker: 본문 어딘가의 바이트를 엔티티 참조로 치환할 (orig, ref) 쌍
    나머지 파트는 그대로 복사해 유효한 .hwpx 컨테이너를 유지한다.
    """
    with zipfile.ZipFile(src, "r") as zin:
        names = zin.namelist()
        raw = {n: zin.read(n) for n in names}

    sec = raw["Contents/section0.xml"]
    # XML 선언 뒤에 DOCTYPE 삽입
    assert sec.lstrip().startswith(b"<?xml"), "예상한 XML 선언이 없음"
    decl_end = sec.index(b"?>") + 2
    sec = sec[:decl_end] + b"\n" + doctype + sec[decl_end:]
    # 본문에 엔티티 참조를 끼워 넣을 자리: 첫 <hp:t> ... </hp:t> 사이에 marker 주입
    orig, ref = marker
    sec = sec.replace(orig, orig + ref, 1)
    raw["Contents/section0.xml"] = sec

    with zipfile.ZipFile(dst, "w") as zout:
        for n in names:
            zout.writestr(n, raw[n])


# ---------------------------------------------------------------------------
# 시나리오 A — XXE / 엔티티 확장 방어 (보안)
# ---------------------------------------------------------------------------

def test_hwpx_internal_entity_not_resolved(tmp_path: Path) -> None:
    """신뢰할 수 없는 .hwpx 의 내부 엔티티가 확장되면 안 된다.

    하드닝되지 않은 파서(resolve_entities=True 기본값)는 &pwn; 을 'PWNED'로
    확장한다. 보안상 untrusted 입력의 엔티티는 절대 확장하면 안 된다.
    """
    good = tmp_path / "good.hwpx"
    _make_hwpx(good, ["HELLO"])
    bad = tmp_path / "bad.hwpx"
    _poison_section0(
        good, bad,
        doctype=b'<!DOCTYPE hs:sec [ <!ENTITY pwn "PWNED"> ]>',
        marker=(b"HELLO", b"&pwn;"),
    )

    pkg = HwpxPackage.open(bad)
    text = pkg.export_text()
    pkg.close()
    assert "PWNED" not in text, "내부 엔티티가 확장됨 — 파서 하드닝 필요"


def test_hwpx_billion_laughs_not_expanded(tmp_path: Path) -> None:
    """billion-laughs 류 엔티티 확장 DoS 방어."""
    good = tmp_path / "good.hwpx"
    _make_hwpx(good, ["BOOM"])
    bad = tmp_path / "bad.hwpx"
    doctype = (
        b'<!DOCTYPE hs:sec [\n'
        b' <!ENTITY a "aaaaaaaaaa">\n'
        b' <!ENTITY b "&a;&a;&a;&a;&a;&a;&a;&a;&a;&a;">\n'
        b' <!ENTITY c "&b;&b;&b;&b;&b;&b;&b;&b;&b;&b;">\n'
        b' <!ENTITY d "&c;&c;&c;&c;&c;&c;&c;&c;&c;&c;">\n'
        b']>'
    )
    _poison_section0(good, bad, doctype=doctype, marker=(b"BOOM", b"&d;"))

    pkg = HwpxPackage.open(bad)
    text = pkg.export_text()
    pkg.close()
    # 확장됐다면 'a'가 수천 자. 하드닝되면 0자.
    assert text.count("a") < 100, "엔티티 확장 DoS 방어 필요"


# ---------------------------------------------------------------------------
# 시나리오 B — ABC 계약 일치 (allow_merge_redirect)
# ---------------------------------------------------------------------------

def test_abc_signature_supports_fill_form() -> None:
    """DocumentAdapter ABC 의 추상 시그니처만 보고 구현한 어댑터로도
    fill_form 이 동작해야 한다.

    base.fill_form → _fill_one_cell 은 set_cell/append_to_cell 을
    allow_merge_redirect= 키워드로 호출한다. 추상 시그니처에 이 인자가
    선언돼 있지 않으면, ABC 계약대로 구현한 어댑터에서 TypeError 가 난다.
    """
    from document_adapter.base import (
        DocumentAdapter,
        TableSchema,
        CellContent,
    )

    class FakeAdapter(DocumentAdapter):
        """ABC 추상 시그니처 '그대로' 구현한 최소 어댑터 (단일 라벨-값 표)."""
        format = "fake"

        def _open(self) -> None:
            # 2x2: (0,0)="성명" 라벨, (0,1)=빈 값셀
            self._cells = {(0, 0): "성명", (0, 1): "", (1, 0): "생년", (1, 1): ""}

        def save(self, path=None):
            return self.path

        def get_placeholders(self):
            return []

        def get_tables(self, min_rows=1, min_cols=1, preview_rows=4, max_cell_len=40):
            preview = [["성명", ""], ["생년", ""]]
            return [TableSchema(index=0, rows=2, cols=2, preview=preview)]

        def get_cell(self, table_index, row, col):
            text = self._cells.get((row, col), "")
            return CellContent(
                row=row, col=col, text=text, paragraphs=[text],
                is_anchor=True, anchor=(row, col), span=(1, 1),
            )

        def render_template(self, context):
            pass

        # 핵심: ABC 추상 시그니처를 '그대로' 따른다 (allow_merge_redirect 없음)
        def set_cell(self, table_index, row, col, value):
            old = self._cells.get((row, col), "")
            self._cells[(row, col)] = value
            return old

        def append_to_cell(self, table_index, row, col, value, separator="  "):
            old = self._cells.get((row, col), "")
            self._cells[(row, col)] = (old + separator + value) if old else value
            return old

        def append_row(self, table_index, values):
            pass

    fa = FakeAdapter(Path("/dev/null"))
    # 추상 시그니처대로 구현했는데 fill_form 이 깨지면 ABC 계약이 거짓.
    result = fa.fill_form({"성명": "홍길동"})
    assert result["filled"], f"fill_form 이 라벨을 채우지 못함: {result}"
    assert fa._cells[(0, 1)] == "홍길동"


# ---------------------------------------------------------------------------
# 시나리오 C — 발굴한 보장(guarantee)을 회귀 테스트로 고정
# ---------------------------------------------------------------------------

def _make_form_hwpx(path: Path, rows: list[tuple[str, str]]) -> None:
    """라벨-값 2열 표를 가진 HWPX 폼 생성."""
    doc = HwpxDocument.new()
    doc.add_paragraph("")
    doc.add_table(len(rows), 2)
    doc.save_to_path(path)
    d = HwpxDocument.open(path)
    try:
        sec = d.sections[0]
        tbl = next(p.tables[0] for p in sec.paragraphs if p.tables)
        for i, (label, value) in enumerate(rows):
            tbl.rows[i].cells[0].text = label
            tbl.rows[i].cells[1].text = value
        d.save_to_path(path)
    finally:
        d.close()


def test_set_cell_escapes_xml_special_chars(tmp_path: Path) -> None:
    """셀 값에 XML 특수문자/태그 문자열이 들어와도 이스케이프되어
    문서 구조를 깨뜨리지 않고 그대로 보존돼야 한다 (injection 방어)."""
    src = tmp_path / "f.hwpx"
    _make_form_hwpx(src, [("제목", ""), ("비고", "")])
    payload = '<hp:t>x</hp:t> & "q" < > 이모지😀'

    ad = load(src)
    ad.set_cell(0, 0, 1, payload)
    out = tmp_path / "out.hwpx"
    ad.save(out)
    ad.close()

    # well-formed XML 유지
    with zipfile.ZipFile(out) as z:
        etree.fromstring(z.read("Contents/section0.xml"))  # noqa: S320 (신뢰 입력)
    # 값 그대로 보존 + 인접 셀 무손상
    ad2 = load(out)
    assert ad2.get_cell(0, 0, 1).text == payload
    assert ad2.get_cell(0, 1, 0).text == "비고"
    ad2.close()


def test_save_preserves_mimetype_first_and_stored(tmp_path: Path) -> None:
    """save() 후에도 mimetype 이 첫 엔트리 + STORED(비압축)로 유지돼야 한다
    (OPC/HWPX 컨테이너 유효성). 안 건드린 파트는 byte-identical."""
    src = tmp_path / "src.hwpx"
    _make_form_hwpx(src, [("a", "b")])

    pkg = HwpxPackage.open(src)
    pkg.get_xml_root("Contents/section0.xml")
    pkg.mark_dirty("Contents/section0.xml")
    out = tmp_path / "out.hwpx"
    pkg.save(out)
    pkg.close()

    with zipfile.ZipFile(src) as z1, zipfile.ZipFile(out) as z2:
        n2 = z2.namelist()
        assert n2[0] == "mimetype", "mimetype 이 첫 엔트리가 아님"
        assert z2.getinfo("mimetype").compress_type == zipfile.ZIP_STORED
        assert z1.read("mimetype") == z2.read("mimetype")
        # header.xml 은 편집 안 했으므로 byte-identical
        assert z1.read("Contents/header.xml") == z2.read("Contents/header.xml")


@pytest.mark.parametrize("data,expect_filled,expect_notfound,expect_ambig", [
    ({}, 0, 0, 0),                          # 빈 입력
    ({"   ": "x", "***": "y"}, 0, 2, 0),     # 정규화하면 빈 라벨 → not_found
])
def test_fill_form_degenerate_inputs(tmp_path: Path, data, expect_filled,
                                     expect_notfound, expect_ambig) -> None:
    """빈/특수문자뿐인 라벨 등 비정상 입력에서도 crash 없이 분류만."""
    src = tmp_path / "f.hwpx"
    _make_form_hwpx(src, [("성명", ""), ("생년월일", "")])
    ad = load(src)
    r = ad.fill_form(data)
    ad.close()
    assert len(r["filled"]) == expect_filled
    assert len(r["not_found"]) == expect_notfound
    assert len(r["ambiguous"]) == expect_ambig


def test_fill_form_duplicate_label_is_ambiguous(tmp_path: Path) -> None:
    """같은 라벨이 여러 곳이면 채우지 않고 ambiguous 로 분류 (오기입 방지)."""
    src = tmp_path / "dup.hwpx"
    _make_form_hwpx(src, [("금액", ""), ("금액", "")])
    ad = load(src)
    r = ad.fill_form({"금액": "100"})
    ad.close()
    assert len(r["ambiguous"]) == 1
    assert len(r["filled"]) == 0


def test_get_cell_out_of_bounds_raises(tmp_path: Path) -> None:
    """경계를 벗어난 좌표는 CellOutOfBoundsError(IndexError 하위)."""
    from document_adapter.base import CellOutOfBoundsError

    src = tmp_path / "f.hwpx"
    _make_form_hwpx(src, [("a", "b")])
    ad = load(src)
    with pytest.raises(CellOutOfBoundsError):
        ad.get_cell(0, 99, 99)
    ad.close()
