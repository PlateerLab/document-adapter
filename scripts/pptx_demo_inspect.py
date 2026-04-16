"""PPTX 구조 파악 — 시연 fixture 선정용."""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from document_adapter import load


def inspect(path: Path) -> None:
    print(f"\n{'=' * 70}\n{path.name}\n{'=' * 70}")
    a = load(path)
    try:
        tables = a.get_tables(preview_rows=20, max_cell_len=60)
        print(f"표 {len(tables)}개")
        for t in tables[:8]:
            print(f"\n  T{t.index}: {t.rows}x{t.cols} (merges={len(t.merges)})"
                  f" parent={t.parent_path}")
            for r, row in enumerate(t.preview[:8]):
                for c, val in enumerate(row):
                    if val is not None:
                        print(f"    ({r:>2},{c:>2}) {val!r}")
    finally:
        a.close()


if __name__ == "__main__":
    candidates = [
        ROOT / "tests" / "fixtures" / "pptx" / "real" / "ai_plan_small.pptx",
        ROOT / "tests" / "fixtures" / "pptx" / "real" / "kubeflow_volumes.pptx",
    ]
    for p in candidates:
        inspect(p)
