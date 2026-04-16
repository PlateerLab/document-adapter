"""HWPX 어댑터: python-hwpx 기반.

버그 회피:
- set_cell_text()는 빈 셀에서 lxml/ElementTree 혼용 에러가 발생 (v2.9.0) →
  cell.paragraphs[0].text 직접 할당으로 우회
- replace_text_in_runs()는 한글 공백이 run으로 쪼개질 때 매칭 실패 →
  위치 기반 편집을 권장

표 구조:
- iter_grid()로 병합 셀(rowSpan/colSpan)을 인식해 logical grid를 구성
- 셀 내부 중첩 테이블은 flat DFS로 인덱싱, parent_path로 위치 표시

부가:
- manifest fallback 로그가 기본적으로 매우 시끄러움 → logging 레벨 조정
"""
from __future__ import annotations

import logging
import re
import warnings
from pathlib import Path
from typing import Any, Iterator

# 경고성 로그 억제 (manifest fallback 등)
logging.getLogger("hwpx").setLevel(logging.ERROR)

from hwpx.document import HwpxDocument

from .base import DocumentAdapter, MergeInfo, TableSchema

TAG_PATTERN = re.compile(r"\{\{\s*(\w+)\s*\}\}")

_HP_NS = "http://www.hancom.co.kr/hwpml/2011/paragraph"
_HP_T = f"{{{_HP_NS}}}t"
_HP_RUN = f"{{{_HP_NS}}}run"
_HP_TBL = f"{{{_HP_NS}}}tbl"


class HwpxAdapter(DocumentAdapter):
    format = "hwpx"

    def _open(self) -> None:
        self._doc = HwpxDocument.open(self.path)

    def save(self, path: Path | str | None = None) -> Path:
        target = Path(path) if path else self.path
        self._doc.save_to_path(target)
        self.path = target
        return target

    def close(self) -> None:
        self._doc.close()

    # ---- helpers ----
    def _iter_tables(self) -> Iterator[tuple[int, Any, str]]:
        """(flat_index, table, parent_path) 순회. 중첩 테이블까지 DFS."""
        idx_counter = [0]

        def walk(tbl, parent_path: str) -> Iterator[tuple[int, Any, str]]:
            current_idx = idx_counter[0]
            idx_counter[0] += 1
            yield current_idx, tbl, parent_path
            # 중첩 테이블: 각 앵커 셀의 tables만 내려간다 (같은 물리 셀 중복 방지)
            seen_cell_ids: set[int] = set()
            for entry in tbl.iter_grid():
                if not entry.is_anchor:
                    continue
                cell = entry.cell
                cell_key = id(cell.element)
                if cell_key in seen_cell_ids:
                    continue
                seen_cell_ids.add(cell_key)
                for child_tbl in cell.tables:
                    child_parent = (
                        f"{parent_path}.tables[{current_idx}].cell"
                        f"({entry.anchor[0]},{entry.anchor[1]})"
                    )
                    yield from walk(child_tbl, child_parent)

        for section in self._doc.sections:
            for para in section.paragraphs:
                for tbl in para.tables:
                    yield from walk(tbl, "")

    def _get_table(self, table_index: int):
        for idx, tbl, _ in self._iter_tables():
            if idx == table_index:
                return tbl
        raise IndexError(f"HWPX table index {table_index} not found")

    @staticmethod
    def _cell_text(cell) -> str:
        """셀의 직접 텍스트만 추출 (중첩 테이블의 텍스트는 제외).

        python-hwpx의 ``paragraph.text``는 ``.//hp:t``로 descendant를 훑어
        중첩 테이블 내부 텍스트까지 흡수한다. LLM에게 이게 그대로 노출되면
        외부 셀의 내용이 중첩 테이블 내용과 뒤섞인 것처럼 보인다.
        따라서 run의 직접 자식 ``<hp:t>``만 읽는다 (중첩된 ``<hp:tbl>`` 서브트리는 자연히 제외).
        """
        parts: list[str] = []
        for para in cell.paragraphs:
            for run in para.element.findall(_HP_RUN):
                for t in run.findall(_HP_T):
                    if t.text:
                        parts.append(t.text)
        return "".join(parts).strip()

    # ---- inspection ----
    def get_placeholders(self) -> list[str]:
        text = self._doc.export_text()
        return sorted(set(TAG_PATTERN.findall(text)))

    def get_tables(self, min_rows: int = 1, min_cols: int = 1,
                   preview_rows: int = 4, max_cell_len: int = 40) -> list[TableSchema]:
        schemas: list[TableSchema] = []
        for idx, tbl, parent_path in self._iter_tables():
            rows, cols = tbl.row_count, tbl.column_count
            if rows < min_rows or cols < min_cols:
                continue

            visible_rows = min(rows, preview_rows)
            # 기본 프리뷰 grid: None 채운 뒤 앵커 위치에만 텍스트 주입
            preview: list[list[str | None]] = [
                [None for _ in range(cols)] for _ in range(visible_rows)
            ]
            merges: list[MergeInfo] = []
            seen_anchors: set[tuple[int, int]] = set()

            for entry in tbl.iter_grid():
                if entry.anchor in seen_anchors:
                    # 같은 앵커는 한 번만
                    if entry.row < visible_rows and entry.is_anchor:
                        pass  # preview는 이미 채웠으므로 skip
                    continue
                if entry.is_anchor:
                    seen_anchors.add(entry.anchor)
                    if entry.row < visible_rows:
                        text = self._cell_text(entry.cell)
                        preview[entry.row][entry.column] = text[:max_cell_len]
                    if entry.span != (1, 1):
                        merges.append(MergeInfo(anchor=entry.anchor, span=entry.span))

            schemas.append(
                TableSchema(
                    index=idx,
                    rows=rows,
                    cols=cols,
                    preview=preview,
                    merges=merges,
                    parent_path=parent_path or None,
                )
            )
        return schemas

    # ---- editing ----
    def render_template(self, context: dict[str, Any]) -> None:
        """본문 + 표 셀의 {{key}}를 paragraph 단위로 치환.

        병합 셀의 경우 같은 앵커의 paragraph를 여러 logical 좌표에서 참조하게 되므로,
        is_anchor 위치만 방문해 중복 치환을 피한다.
        """

        def substitute(para) -> None:
            text = para.text
            if TAG_PATTERN.search(text):
                para.text = TAG_PATTERN.sub(
                    lambda m: str(context.get(m.group(1), m.group(0))), text
                )

        # 본문
        for section in self._doc.sections:
            for para in section.paragraphs:
                substitute(para)
        # 표 셀 (중첩 테이블 포함; _iter_tables가 DFS)
        for _, tbl, _ in self._iter_tables():
            for entry in tbl.iter_grid():
                if not entry.is_anchor:
                    continue
                for para in entry.cell.paragraphs:
                    substitute(para)

    def set_cell(
        self,
        table_index: int,
        row: int,
        col: int,
        value: str,
        *,
        allow_merge_redirect: bool = False,
    ) -> str:
        """셀 값 교체. 원래 값 반환.

        병합 셀(non-anchor) 좌표로 호출하면 기본적으로 ``ValueError``를 발생시킨다.
        이는 LLM이 병합 구조를 잘못 이해하고 엉뚱한 앵커를 덮어쓰는 것을 방지한다.
        ``allow_merge_redirect=True``를 주면 앵커로 자동 리디렉트하고 경고만 남긴다.

        set_cell_text 버그 우회: paragraph.text 직접 할당.
        """
        tbl = self._get_table(table_index)
        if row < 0 or col < 0 or row >= tbl.row_count or col >= tbl.column_count:
            raise IndexError(
                f"cell ({row},{col}) out of bounds for table {table_index} "
                f"({tbl.row_count}x{tbl.column_count})"
            )

        grid_entry = None
        for entry in tbl.iter_grid():
            if (entry.row, entry.column) == (row, col):
                grid_entry = entry
                break
        if grid_entry is None:
            raise IndexError(
                f"cell ({row},{col}) does not resolve to any physical cell"
            )

        if not grid_entry.is_anchor:
            anchor_r, anchor_c = grid_entry.anchor
            if not allow_merge_redirect:
                raise ValueError(
                    f"cell ({row},{col}) is part of a merged region anchored at "
                    f"({anchor_r},{anchor_c}) span={grid_entry.span}. "
                    f"Write to the anchor coordinate, or pass "
                    f"allow_merge_redirect=True."
                )
            warnings.warn(
                f"set_cell({row},{col}) redirected to merge anchor "
                f"({anchor_r},{anchor_c})",
                stacklevel=2,
            )

        cell = grid_entry.cell
        paragraphs = list(cell.paragraphs)
        old = self._cell_text(cell)
        if paragraphs:
            paragraphs[0].text = value
            for p in paragraphs[1:]:
                p.text = ""
        return old

    def append_row(self, table_index: int, values: list[str]) -> None:
        """python-hwpx에는 표준 add_row API가 없음.
        대안: 템플릿에 충분한 빈 행을 미리 만들고 set_cell로 채우는 전략."""
        raise NotImplementedError(
            "HWPX는 python-hwpx에 동적 행 추가 공식 API가 없음. "
            "템플릿에 여분 행을 두고 set_cell로 채우는 방식을 권장."
        )
