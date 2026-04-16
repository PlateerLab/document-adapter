"""양식 HWPX 구조 파악 — 어느 셀이 라벨이고 어느 셀이 값인지 preview."""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from document_adapter import load

FIXTURE = ROOT / "tests" / "fixtures" / "hwpx" / "real" / "stop_payment_blank.hwpx"


def main() -> int:
    adapter = load(FIXTURE)
    try:
        tables = adapter.get_tables(preview_rows=30, max_cell_len=30)
        for t in tables:
            print(f"\nTable #{t.index}: {t.rows}x{t.cols}, merges={len(t.merges)}")
            # merge anchor 표시
            merge_anchors = {m.anchor: m.span for m in t.merges}
            for r, row in enumerate(t.preview):
                for c, val in enumerate(row):
                    if val is None:
                        marker = "  ·"
                    else:
                        span = merge_anchors.get((r, c))
                        tag = f"[{span[0]}x{span[1]}]" if span else ""
                        marker = f"({r:>2},{c:>2}){tag} {val!r}"
                        print(f"    {marker}")
    finally:
        adapter.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
