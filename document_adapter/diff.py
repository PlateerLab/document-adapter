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

        result: dict[str, Any] = {
            "changed": len(changes),
            "changes": changes,
            "tables_added": sorted(set(tb) - set(ta)),
            "tables_removed": sorted(set(ta) - set(tb)),
        }

        # 차트 비교 (PPTX — 다른 포맷은 get_charts 가 빈 리스트라 no-op).
        # 차트 없는 문서에서는 키 자체를 생략해 기존 반환 형태를 보존한다.
        ca = {(c.slide_index, c.shape_id): c for c in a.get_charts()}
        cb = {(c.slide_index, c.shape_id): c for c in b.get_charts()}
        chart_changes: list[dict[str, Any]] = []
        for key in sorted(set(ca) & set(cb)):
            ch_a, ch_b = ca[key], cb[key]
            cats = ch_b.categories if len(ch_b.categories) >= len(ch_a.categories) \
                else ch_a.categories
            for si in range(max(len(ch_a.series), len(ch_b.series))):
                ser_a = ch_a.series[si] if si < len(ch_a.series) else None
                ser_b = ch_b.series[si] if si < len(ch_b.series) else None
                name = (ser_b or ser_a or {}).get("name", f"Series {si + 1}")
                vals_a = ser_a["values"] if ser_a else []
                vals_b = ser_b["values"] if ser_b else []
                for ci in range(max(len(vals_a), len(vals_b))):
                    x = vals_a[ci] if ci < len(vals_a) else None
                    y = vals_b[ci] if ci < len(vals_b) else None
                    if x == y:
                        continue
                    chart_changes.append({
                        "slide_index": key[0],
                        "shape_id": key[1],
                        "series": name,
                        "category": cats[ci] if ci < len(cats) else str(ci),
                        "before": x,
                        "after": y,
                    })
        charts_added = sorted(set(cb) - set(ca))
        charts_removed = sorted(set(ca) - set(cb))
        if chart_changes or charts_added or charts_removed:
            result["chart_changes"] = chart_changes
            result["charts_added"] = [list(k) for k in charts_added]
            result["charts_removed"] = [list(k) for k in charts_removed]

        return result
    finally:
        a.close()
        b.close()
