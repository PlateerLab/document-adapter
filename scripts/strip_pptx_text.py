"""PPTX 에서 텍스트 내용만 지우고 레이아웃/디자인은 유지.

공공 PPTX (디자인/폰트/색상/도형/표 구조) 를 재사용 가능한 빈 양식으로 변환.

전략: 모든 <a:t> (runtime text) 의 text 를 빈 문자열로 치환. run properties
(rPr, 폰트/색상) 와 paragraph properties (pPr, 정렬 등) 는 그대로 유지.
도형/이미지/배경/표 구조는 건드리지 않음.
"""
from __future__ import annotations

import shutil
import sys
import zipfile
from io import BytesIO
from pathlib import Path

from lxml import etree

A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"


def strip_text(src: Path, dst: Path) -> None:
    shutil.copy2(src, dst)

    with zipfile.ZipFile(dst, "r") as zin:
        entries = {n: zin.read(n) for n in zin.namelist()}
        infos = {n: zin.getinfo(n) for n in zin.namelist()}

    parser = etree.XMLParser(remove_blank_text=False, strip_cdata=False)
    changed = 0
    kept_unchanged = 0

    for name, data in list(entries.items()):
        # 슬라이드 본문만 대상 (master/layout 은 유지)
        if not name.startswith("ppt/slides/slide") or not name.endswith(".xml"):
            continue
        try:
            root = etree.fromstring(data, parser=parser)
        except etree.XMLSyntaxError:
            kept_unchanged += 1
            continue
        cleared = 0
        for t in root.iter(f"{A_NS}t"):
            if t.text:
                t.text = ""
                cleared += 1
        if cleared:
            entries[name] = etree.tostring(
                root, xml_declaration=True, encoding="UTF-8", standalone=True
            )
            changed += 1
            print(f"  {name}: {cleared}개 text 비움")

    # 재압축
    with zipfile.ZipFile(dst, "w") as zout:
        for n in [info.filename for info in zin.infolist()] if False else list(infos.keys()):
            data = entries[n]
            info = infos[n]
            new_info = zipfile.ZipInfo(filename=n, date_time=info.date_time)
            new_info.compress_type = info.compress_type
            new_info.external_attr = info.external_attr
            zout.writestr(new_info, data)

    print(f"\n총 {changed} 개 슬라이드 XML 에서 텍스트 제거 완료")


def main() -> int:
    src = Path(__file__).resolve().parent.parent / "tests" / "fixtures" / "pptx" / "real" / "gov_policy_2025.pptx"
    if not src.exists():
        print(f"원본 없음: {src}")
        return 2

    dst = Path.home() / "Desktop" / "pptx_report_demo" / "gov_policy_2025_blank.pptx"
    dst.parent.mkdir(exist_ok=True)

    print(f"원본: {src.name} ({src.stat().st_size:,} bytes)")
    strip_text(src, dst)
    print(f"\n결과: {dst} ({dst.stat().st_size:,} bytes)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
