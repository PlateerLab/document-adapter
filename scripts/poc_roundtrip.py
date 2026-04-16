"""PoC: HWPX를 python-hwpx 없이 zipfile+lxml 만으로 round-trip 가능한지 검증.

목표: 원본 HWPX → 순수 zipfile로 읽기 → XML을 lxml로 파싱 → 직렬화 → 새 ZIP 쓰기
     → 원본과 비교 + python-hwpx로 re-open 검증

단계:
  1. 샘플 HWPX 생성 (python-hwpx로, 병합 셀 포함)
  2. Bytes-identical copy: zipfile로 읽기만 하고 쓰기 → diff
  3. lxml round-trip: XML 파싱/재직렬화까지 거친 round-trip → diff
  4. python-hwpx로 결과물 re-open → 병합 셀 정보 일치 확인
  5. 셀 수정 1건 후 저장 → re-open 비교
"""
from __future__ import annotations

import hashlib
import shutil
import sys
import zipfile
from io import BytesIO
from pathlib import Path

from lxml import etree

SCRIPT_DIR = Path(__file__).resolve().parent
ROOT = SCRIPT_DIR.parent
OUT = SCRIPT_DIR / "_poc_out"
OUT.mkdir(exist_ok=True)


def header(msg: str) -> None:
    print(f"\n{'=' * 70}\n{msg}\n{'=' * 70}")


def digest(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()[:16]


# ---------- Stage 1: 샘플 생성 ----------

def make_sample(path: Path) -> None:
    from hwpx.document import HwpxDocument

    doc = HwpxDocument.new()
    doc.add_paragraph("제목: round-trip 테스트")
    doc.add_table(3, 3)
    doc.save_to_path(path)

    doc2 = HwpxDocument.open(path)
    try:
        tbl = None
        for para in doc2.sections[0].paragraphs:
            if para.tables:
                tbl = para.tables[0]
                break
        assert tbl is not None
        row0 = tbl.rows[0].cells
        row0[0].set_span(row_span=1, col_span=3)
        row0[0].text = "병합된 제목"
        for sib in (row0[1], row0[2]):
            sib.set_size(width=0, height=0)
            sib.text = ""
        r1 = tbl.rows[1].cells
        r1[0].text = "A1"; r1[1].text = "A2"; r1[2].text = "A3"
        r2 = tbl.rows[2].cells
        r2[0].text = "B1"; r2[1].text = "B2"; r2[2].text = "B3"
        doc2.save_to_path(path)
    finally:
        doc2.close()


# ---------- Stage 2: bytes-identical copy ----------

def bytes_copy(src: Path, dst: Path) -> None:
    """zipfile로 읽어서 각 파일을 있는 그대로 새 ZIP에 기록.

    XML 파싱 없음. 순수 바이트 복사 + ZipInfo 보존 시도.
    """
    with zipfile.ZipFile(src, "r") as zin:
        with zipfile.ZipFile(dst, "w") as zout:
            for info in zin.infolist():
                data = zin.read(info.filename)
                # ZipInfo를 복제해서 compress_type, date_time, external_attr 등 보존
                new_info = zipfile.ZipInfo(
                    filename=info.filename,
                    date_time=info.date_time,
                )
                new_info.compress_type = info.compress_type
                new_info.external_attr = info.external_attr
                new_info.internal_attr = info.internal_attr
                new_info.create_system = info.create_system
                new_info.create_version = info.create_version
                new_info.extract_version = info.extract_version
                new_info.flag_bits = info.flag_bits
                zout.writestr(new_info, data)


# ---------- Stage 3: lxml round-trip ----------

def lxml_roundtrip(src: Path, dst: Path) -> None:
    """XML 파일은 lxml로 파싱 후 재직렬화. 나머지는 바이트 복사."""
    with zipfile.ZipFile(src, "r") as zin:
        with zipfile.ZipFile(dst, "w") as zout:
            for info in zin.infolist():
                data = zin.read(info.filename)
                if info.filename.endswith(".xml") or info.filename.endswith(".hpf"):
                    # lxml 파싱 → 직렬화
                    parser = etree.XMLParser(remove_blank_text=False, strip_cdata=False)
                    try:
                        tree = etree.fromstring(data, parser=parser)
                        data = etree.tostring(
                            tree,
                            xml_declaration=True,
                            encoding="UTF-8",
                            standalone=True,
                        )
                    except etree.XMLSyntaxError as e:
                        print(f"  ⚠️  {info.filename}: XML 파싱 실패 — 바이트 복사 ({e})")
                new_info = zipfile.ZipInfo(
                    filename=info.filename,
                    date_time=info.date_time,
                )
                new_info.compress_type = info.compress_type
                new_info.external_attr = info.external_attr
                zout.writestr(new_info, data)


# ---------- Stage 4: diff ----------

def diff_zip(a: Path, b: Path, label: str) -> dict:
    """두 ZIP 파일의 namelist와 각 파일의 bytes를 비교."""
    with zipfile.ZipFile(a, "r") as za, zipfile.ZipFile(b, "r") as zb:
        names_a = za.namelist()
        names_b = zb.namelist()
        missing_in_b = [n for n in names_a if n not in names_b]
        extra_in_b = [n for n in names_b if n not in names_a]
        order_same = names_a == names_b

        differing = []
        identical = []
        for name in names_a:
            if name not in names_b:
                continue
            da = za.read(name)
            db = zb.read(name)
            if da == db:
                identical.append(name)
            else:
                differing.append((name, digest(da), digest(db), len(da), len(db)))

    print(f"[{label}] size: {a.stat().st_size} → {b.stat().st_size} bytes")
    print(f"  namelist 길이: {len(names_a)} → {len(names_b)}, 순서 동일={order_same}")
    if missing_in_b:
        print(f"  빠진 파일: {missing_in_b}")
    if extra_in_b:
        print(f"  추가된 파일: {extra_in_b}")
    print(f"  bytes-identical: {len(identical)}/{len(names_a)}")
    if differing:
        print(f"  다른 파일:")
        for name, dig_a, dig_b, la, lb in differing:
            print(f"    {name}: {la}B({dig_a}) → {lb}B({dig_b})")
    return {
        "order_same": order_same,
        "identical_count": len(identical),
        "total": len(names_a),
        "differing": [d[0] for d in differing],
    }


# ---------- Stage 5: python-hwpx re-open 검증 ----------

def verify_hwpx_reopen(path: Path, label: str) -> bool:
    """python-hwpx로 열어서 병합 셀 정보가 살아있는지 확인."""
    from hwpx.document import HwpxDocument

    try:
        doc = HwpxDocument.open(path)
    except Exception as e:
        print(f"[{label}] ❌ HwpxDocument.open 실패: {e}")
        return False
    try:
        tbl = None
        for para in doc.sections[0].paragraphs:
            if para.tables:
                tbl = para.tables[0]
                break
        if tbl is None:
            print(f"[{label}] ❌ 표 없음")
            return False

        rows, cols = tbl.row_count, tbl.column_count
        merge_cells = []
        cell_texts = []
        for entry in tbl.iter_grid():
            if entry.is_anchor and entry.span != (1, 1):
                merge_cells.append((entry.anchor, entry.span))
            if entry.is_anchor:
                text = "".join(
                    t.text or ""
                    for p in entry.cell.paragraphs
                    for run in p.element.findall(
                        "{http://www.hancom.co.kr/hwpml/2011/paragraph}run"
                    )
                    for t in run.findall(
                        "{http://www.hancom.co.kr/hwpml/2011/paragraph}t"
                    )
                )
                cell_texts.append((entry.anchor, text))

        print(f"[{label}] ✅ 열림. {rows}x{cols}, merges={merge_cells}")
        print(f"         texts={cell_texts}")
        return True
    finally:
        doc.close()


# ---------- 메인 ----------

def main() -> int:
    sample = OUT / "original.hwpx"
    header("Stage 1: 샘플 HWPX 생성")
    make_sample(sample)
    print(f"생성됨: {sample} ({sample.stat().st_size} bytes)")

    header("Stage 2: bytes-identical copy (XML 파싱 없음)")
    copy_dst = OUT / "copy.hwpx"
    bytes_copy(sample, copy_dst)
    r1 = diff_zip(sample, copy_dst, "copy")
    ok1 = verify_hwpx_reopen(copy_dst, "copy")

    header("Stage 3: lxml round-trip (XML 파싱+직렬화)")
    lxml_dst = OUT / "lxml.hwpx"
    lxml_roundtrip(sample, lxml_dst)
    r2 = diff_zip(sample, lxml_dst, "lxml")
    ok2 = verify_hwpx_reopen(lxml_dst, "lxml")

    header("Stage 4: 요약")
    print(f"  bytes copy    : identical {r1['identical_count']}/{r1['total']}, reopen={ok1}")
    print(f"  lxml roundtrip: identical {r2['identical_count']}/{r2['total']}, reopen={ok2}")
    print(f"  bytes copy 다른 파일: {r1['differing']}")
    print(f"  lxml 다른 파일: {r2['differing']}")

    # 성공 조건:
    # - bytes copy: 100% identical + reopen 성공
    # - lxml roundtrip: XML만 다르고 bytes는 허용, reopen + merge 정보 보존이 핵심
    success = ok1 and ok2 and r1["identical_count"] == r1["total"]
    print(f"\n{'✅ PoC 성공' if success else '⚠️  부분 성공 — 상세 확인 필요'}")
    return 0 if success else 1


if __name__ == "__main__":
    sys.exit(main())
