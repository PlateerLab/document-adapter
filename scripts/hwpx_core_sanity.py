"""hwpx_core의 grid 순회가 python-hwpx와 동일한 결과를 내는지 확인.

각 fixture에 대해:
  - python-hwpx의 Table.iter_grid() 결과
  - hwpx_core.iter_grid() 결과
둘을 비교 (tables, merges, cell text).
"""
from __future__ import annotations

import logging
import sys
import warnings
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from document_adapter.hwpx_core import (
    HP_P,
    HP_RUN,
    HP_T,
    HP_TBL,
    HwpxPackage,
    iter_grid,
)

FIXTURES = ROOT / "tests" / "fixtures" / "hwpx" / "real"


def core_extract(path: Path) -> dict:
    pkg = HwpxPackage.open(path)
    try:
        tables = 0
        merges = 0
        texts: list[tuple[int, int, str]] = []
        total_rows = 0
        total_cols = 0
        # iterate top-level tables in all sections (재귀는 나중에 확장)
        stack: list = []
        for _, root in pkg.iter_section_roots():
            for tbl in root.iter(HP_TBL):
                stack.append(tbl)

        for tbl in stack:
            tables += 1
            seen_anchors: set[tuple[int, int]] = set()
            col_max = 0
            row_max = 0
            for entry in iter_grid(tbl):
                if entry.is_anchor and entry.anchor in seen_anchors:
                    continue
                if entry.is_anchor:
                    seen_anchors.add(entry.anchor)
                    if entry.span != (1, 1):
                        merges += 1
                    # 앵커 셀의 직접 자식 텍스트만 (중첩 테이블 제외)
                    tc = entry.cell_element
                    for p in tc.iter(HP_P):
                        # 직접 자식 run만 훑는다 (중첩 tbl 제외)
                        for run in p.findall(HP_RUN):
                            for t in run.findall(HP_T):
                                if t.text:
                                    texts.append((entry.row, entry.column, t.text))
                row_max = max(row_max, entry.row + 1)
                col_max = max(col_max, entry.column + 1)
            total_rows += row_max
            total_cols += col_max

        return {
            "tables": tables,
            "merges": merges,
            "total_rows": total_rows,
            "total_cols": total_cols,
            "text_sample": texts[:5],
            "text_count": len(texts),
        }
    finally:
        pkg.close()


def hwpx_extract(path: Path) -> dict:
    warnings.filterwarnings("ignore")
    logging.getLogger("hwpx").setLevel(logging.ERROR)
    from hwpx.document import HwpxDocument

    doc = HwpxDocument.open(path)
    try:
        tables = 0
        merges = 0
        texts: list[tuple[int, int, str]] = []
        total_rows = 0
        total_cols = 0

        def walk(tbl) -> None:
            nonlocal tables, merges, total_rows, total_cols
            tables += 1
            total_rows += tbl.row_count
            total_cols += tbl.column_count
            seen_anchors: set[tuple[int, int]] = set()
            for entry in tbl.iter_grid():
                if entry.is_anchor:
                    if entry.anchor in seen_anchors:
                        continue
                    seen_anchors.add(entry.anchor)
                    if entry.span != (1, 1):
                        merges += 1
                    tc = entry.cell.element
                    for p in tc.iter(HP_P):
                        for run in p.findall(HP_RUN):
                            for t in run.findall(HP_T):
                                if t.text:
                                    texts.append((entry.row, entry.column, t.text))
                    for child in entry.cell.tables:
                        walk(child)

        for section in doc.sections:
            for para in section.paragraphs:
                for tbl in para.tables:
                    walk(tbl)
        return {
            "tables": tables,
            "merges": merges,
            "total_rows": total_rows,
            "total_cols": total_cols,
            "text_sample": texts[:5],
            "text_count": len(texts),
        }
    finally:
        doc.close()


def main() -> int:
    fixtures = sorted(FIXTURES.glob("*.hwpx"))
    all_ok = True
    for p in fixtures:
        print(f"\n--- {p.name} ---")
        try:
            core = core_extract(p)
        except Exception as e:
            print(f"  core ERROR: {type(e).__name__}: {e}")
            all_ok = False
            continue
        try:
            hwpx = hwpx_extract(p)
        except Exception as e:
            print(f"  hwpx ERROR: {type(e).__name__}: {e}")
            all_ok = False
            continue

        # core는 top-level 테이블만, hwpx는 중첩 포함. 비교 위해 필드별 명시.
        print(f"  core: tables={core['tables']}, merges={core['merges']}, text_count={core['text_count']}")
        print(f"  hwpx: tables={hwpx['tables']}, merges={hwpx['merges']}, text_count={hwpx['text_count']}")
        # text_sample 동일성 (동일 위치 기준)
        matches = sum(1 for t in core["text_sample"] if t in hwpx["text_sample"])
        print(f"  text sample overlap: {matches}/{len(core['text_sample'])}")

    return 0 if all_ok else 1


if __name__ == "__main__":
    sys.exit(main())
