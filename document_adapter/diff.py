"""문서 diff — 편집 전/후를 셀 단위로 비교해 변경 사항을 구조적으로 반환.

LLM 워크플로우의 **검증 도구**: 원본을 복사해 두고 fill_form/set_cell 로 편집한 뒤
``diff_documents(원본, 편집본)`` 으로 "무엇이 어디서 어떻게 바뀌었나 + 오버플로
위험은 없나" 를 확인한다. 사람이 스크린샷으로 잡던 검증을 자동화한다.

포맷 무관(docx/pptx/hwpx/xlsx) — get_tables 의 preview 그리드를 비교한다.
"""
from __future__ import annotations

from pathlib import Path
from typing import Any

from document_adapter import load
from document_adapter.base import _overflow_risk

# 비교 시 전체 셀을 보기 위한 충분히 큰 한계(폼·보고서 수준에서 안전).
_FULL_ROWS = 100_000
_FULL_LEN = 100_000


def _grid_get(preview: list[list[str | None]], r: int, c: int) -> str | None:
    if r < len(preview) and c < len(preview[r]):
        return preview[r][c]
    return None


def diff_documents(
    path_a: str | Path,
    path_b: str | Path,
    *,
    include_overflow: bool = True,
) -> dict[str, Any]:
    """두 문서를 셀 단위로 비교. (A=이전/원본, B=이후/편집본)

    Returns:
        {
          "changed": 변경 셀 수,
          "changes": [{table_index, location, row, col, before, after,
                       overflow_risk?}, ...],
          "tables_added": [...], "tables_removed": [...],
        }
    """
    a = load(path_a)
    b = load(path_b)
    try:
        ta = {t.index: t for t in
              a.get_tables(preview_rows=_FULL_ROWS, max_cell_len=_FULL_LEN)}
        tb = {t.index: t for t in
              b.get_tables(preview_rows=_FULL_ROWS, max_cell_len=_FULL_LEN)}

        changes: list[dict[str, Any]] = []
        for idx in sorted(set(ta) | set(tb)):
            sa, sb = ta.get(idx), tb.get(idx)
            if sa is None or sb is None:
                continue
            rows = max(sa.rows, sb.rows)
            cols = max(sa.cols, sb.cols)
            for r in range(rows):
                for c in range(cols):
                    va = _grid_get(sa.preview, r, c)
                    vb = _grid_get(sb.preview, r, c)
                    if (va or "") == (vb or ""):
                        continue
                    entry: dict[str, Any] = {
                        "table_index": idx,
                        "location": sb.location or sa.location,
                        "row": r, "col": c,
                        "before": va, "after": vb,
                    }
                    if include_overflow and vb:
                        try:
                            cell = b.get_cell(idx, r, c)
                            entry["overflow_risk"] = _overflow_risk(vb, cell.width_cm)
                        except Exception:
                            pass
                    changes.append(entry)

        return {
            "changed": len(changes),
            "changes": changes,
            "tables_added": sorted(set(tb) - set(ta)),
            "tables_removed": sorted(set(ta) - set(tb)),
        }
    finally:
        a.close()
        b.close()
