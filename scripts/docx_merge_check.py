"""DOCX 복잡 병합 표 처리 명시적 검증.

regression Stage D 는 "첫 anchor 셀에 sentinel 쓰기" 만 체크하므로 병합 셀
non-anchor 거부, anchor 정확성, set_cell 후 병합 구조 보존 등은 미검증.

loan_products_detail.docx (48 merges), x2bee_checklist.docx (20 merges) 에서:
  1. get_tables preview 로 병합 구조 확인
  2. 병합 anchor 셀에 set_cell → 성공 기대
  3. 병합 non-anchor 좌표에 set_cell → MergedCellWriteError 기대
  4. allow_merge_redirect=True 로 재호출 → anchor 로 redirect
  5. save → reload → 구조 보존
"""
from __future__ import annotations

import shutil
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from document_adapter import load
from document_adapter.base import (
    CellOutOfBoundsError,
    MergedCellWriteError,
    TableIndexError,
)


def inspect_merges(path: Path, max_tables: int = 3) -> None:
    """처음 몇 개 테이블의 병합 구조를 출력."""
    a = load(path)
    try:
        tables = a.get_tables(preview_rows=8, max_cell_len=30)
        print(f"\n=== {path.name} — 처음 {max_tables}개 테이블 ===")
        for t in tables[:max_tables]:
            print(f"\nT{t.index}: {t.rows}x{t.cols}, merges={len(t.merges)}")
            for m in t.merges[:8]:
                print(f"  merge anchor={m.anchor} span={m.span}")
            if len(t.merges) > 8:
                print(f"  ... (+{len(t.merges) - 8} more merges)")
            # preview 일부
            for r, row in enumerate(t.preview[:4]):
                row_str = " | ".join(
                    "None" if v is None else f"{v!r}"[:20]
                    for v in row
                )
                print(f"  row {r}: {row_str}")
    finally:
        a.close()


def find_merged_cells(path: Path) -> list[tuple[int, int, int, int]]:
    """(table_idx, anchor_row, anchor_col, span_rows, span_cols) 중 병합 셀만.

    Returns: list of (t_idx, anchor_r, anchor_c, span_r, span_c).
    """
    merged = []
    a = load(path)
    try:
        for t in a.get_tables(preview_rows=10_000, max_cell_len=50):
            for m in t.merges:
                merged.append((t.index, m.anchor[0], m.anchor[1], m.span[0], m.span[1]))
    finally:
        a.close()
    return merged


def run_merge_tests(path: Path) -> None:
    print(f"\n{'=' * 70}")
    print(f"병합 처리 검증: {path.name}")
    print(f"{'=' * 70}")

    merged = find_merged_cells(path)
    if not merged:
        print("병합 셀 없음 — 스킵")
        return
    print(f"총 병합 셀 {len(merged)} 개")

    # 첫 병합 셀로 테스트
    t_idx, ar, ac, rs, cs = merged[0]
    print(f"\n선택: T{t_idx} anchor=({ar},{ac}) span=({rs},{cs})")

    # 임시 파일에 copy
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
        shutil.copy2(path, tmp.name)
        work = Path(tmp.name)

    # Test 1: anchor 에 쓰기 — 성공 기대
    print("\n[1] anchor 셀에 set_cell (성공 기대):")
    a = load(work)
    try:
        old = a.set_cell(t_idx, ar, ac, "__ANCHOR_EDIT__")
        a.save()
        print(f"    ✅ old={old!r}")
    except Exception as e:
        print(f"    ❌ {type(e).__name__}: {e}")
    finally:
        a.close()

    # Test 2: non-anchor 에 쓰기 — MergedCellWriteError 기대
    if rs > 1 or cs > 1:
        # span 내부의 non-anchor 좌표
        non_anchor_r = ar + (rs - 1 if rs > 1 else 0)
        non_anchor_c = ac + (cs - 1 if cs > 1 else 0)
        if (non_anchor_r, non_anchor_c) != (ar, ac):
            print(f"\n[2] non-anchor ({non_anchor_r},{non_anchor_c}) 에 set_cell (거부 기대):")
            a = load(work)
            try:
                a.set_cell(t_idx, non_anchor_r, non_anchor_c, "should_fail")
                print(f"    ❌ 예상: MergedCellWriteError, 실제: 성공 (버그!)")
            except MergedCellWriteError as e:
                print(f"    ✅ 정상 거부: {str(e)[:120]}")
            except Exception as e:
                print(f"    ⚠️  예상 외: {type(e).__name__}: {e}")
            finally:
                a.close()

            # Test 3: allow_merge_redirect=True
            print(f"\n[3] non-anchor ({non_anchor_r},{non_anchor_c}) + allow_merge_redirect=True "
                  f"(anchor 로 redirect 기대):")
            a = load(work)
            try:
                a.set_cell(
                    t_idx, non_anchor_r, non_anchor_c, "__REDIRECTED__",
                    allow_merge_redirect=True,
                )
                a.save()
                # anchor 셀 값 확인
                cell = a.get_cell(t_idx, ar, ac)
                if "__REDIRECTED__" in cell.text:
                    print(f"    ✅ anchor 로 redirect 되어 값 반영: {cell.text.strip()[:60]!r}")
                else:
                    print(f"    ❌ anchor text 에 redirected 값 없음: {cell.text.strip()[:60]!r}")
            except Exception as e:
                print(f"    ❌ {type(e).__name__}: {e}")
            finally:
                a.close()

    # Test 4: save 후 reload 했을 때 병합 구조 보존
    print("\n[4] save → reload 후 병합 구조 보존 확인:")
    merged_after = find_merged_cells(work)
    if len(merged_after) == len(merged):
        print(f"    ✅ merge count 보존: {len(merged)} → {len(merged_after)}")
    else:
        print(f"    ⚠️  merge count 변화: {len(merged)} → {len(merged_after)}")

    try:
        work.unlink()
    except OSError:
        pass


def main() -> int:
    fixtures_dir = ROOT / "tests" / "fixtures" / "docx" / "real"
    targets = [
        fixtures_dir / "loan_products_detail.docx",   # 48 merges
        fixtures_dir / "x2bee_checklist.docx",        # 20 merges
    ]
    for p in targets:
        if not p.exists():
            print(f"없음: {p}")
            continue
        inspect_merges(p, max_tables=2)
        run_merge_tests(p)
    return 0


if __name__ == "__main__":
    sys.exit(main())
