"""HWPX 패키지: ZIP + XML 파트 관리.

python-hwpx의 ``HwpxDocument``를 대체한다. 핵심 전략:
  - 열 때: 모든 파트를 원본 bytes로 보존 (ZipInfo 메타 포함)
  - 편집: lxml 트리를 lazy 파싱, dirty 플래그 기록
  - 저장: dirty 파트만 재직렬화, 나머지는 bytes 그대로 복사

이 전략으로 수정하지 않은 파일은 원본과 byte-identical하게 유지된다.
"""
from __future__ import annotations

import re
import zipfile
from pathlib import Path
from typing import Iterator

from lxml import etree

from document_adapter.hwpx_core.constants import HS_SEC

_SECTION_RE = re.compile(r"^Contents/section(\d+)\.xml$")


class HwpxPackage:
    """HWPX 문서의 ZIP 컨테이너 표현."""

    def __init__(self, path: Path | str) -> None:
        self._path = Path(path)
        self._zip_infos: list[zipfile.ZipInfo] = []
        self._raw: dict[str, bytes] = {}
        self._trees: dict[str, etree._ElementTree] = {}
        self._dirty: set[str] = set()
        self._load()

    # ---- lifecycle ----

    @classmethod
    def open(cls, path: Path | str) -> "HwpxPackage":
        return cls(path)

    def _load(self) -> None:
        with zipfile.ZipFile(self._path, "r") as zf:
            for info in zf.infolist():
                self._zip_infos.append(info)
                self._raw[info.filename] = zf.read(info.filename)

    def close(self) -> None:
        self._trees.clear()
        self._raw.clear()
        self._zip_infos.clear()
        self._dirty.clear()

    # ---- 속성 ----

    @property
    def path(self) -> Path:
        return self._path

    def namelist(self) -> list[str]:
        return [info.filename for info in self._zip_infos]

    def has_part(self, name: str) -> bool:
        return name in self._raw

    # ---- XML 파트 접근 ----

    def get_xml_root(self, name: str) -> etree._Element:
        """XML 파트의 루트 Element 반환. 수정하려면 mark_dirty(name) 호출 필요."""
        if name not in self._trees:
            if name not in self._raw:
                raise KeyError(f"part not found: {name}")
            parser = etree.XMLParser(remove_blank_text=False, strip_cdata=False)
            root = etree.fromstring(self._raw[name], parser=parser)
            tree = root.getroottree()
            self._trees[name] = tree
        return self._trees[name].getroot()

    def mark_dirty(self, name: str) -> None:
        if name not in self._raw:
            raise KeyError(f"part not found: {name}")
        self._dirty.add(name)

    def is_dirty(self, name: str) -> bool:
        return name in self._dirty

    # ---- 섹션 순회 ----

    def list_section_names(self) -> list[str]:
        """Contents/section{N}.xml 파트를 번호 순서대로."""
        sections: list[tuple[int, str]] = []
        for name in self.namelist():
            m = _SECTION_RE.match(name)
            if m:
                sections.append((int(m.group(1)), name))
        sections.sort()
        return [name for _, name in sections]

    def iter_section_roots(self) -> Iterator[tuple[str, etree._Element]]:
        """(part_name, section_root) 순회."""
        for name in self.list_section_names():
            yield name, self.get_xml_root(name)

    def export_text(self) -> str:
        """모든 섹션의 <hp:t> 텍스트를 순서대로 이어붙여 반환."""
        from document_adapter.hwpx_core.constants import HP_RUN, HP_T

        parts: list[str] = []
        for _, root in self.iter_section_roots():
            for t in root.iter(HP_T):
                if t.text:
                    parts.append(t.text)
        return "".join(parts)

    # ---- 저장 ----

    def save(self, path: Path | str | None = None) -> Path:
        target = Path(path) if path else self._path
        with zipfile.ZipFile(target, "w") as zout:
            for info in self._zip_infos:
                name = info.filename
                if name in self._dirty and name in self._trees:
                    tree = self._trees[name]
                    data = etree.tostring(
                        tree,
                        xml_declaration=True,
                        encoding="UTF-8",
                        standalone=True,
                    )
                else:
                    data = self._raw[name]

                new_info = zipfile.ZipInfo(
                    filename=name,
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

        self._path = target
        return target
