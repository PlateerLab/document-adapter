"""분기 보고서 PPTX fill_form 시연.

quarterly_report_synth.pptx (5 슬라이드, 5 표, 2 병합) 에 LLM 처럼 fill_form
과 set_cell 혼합 호출로 보고서 채우기.

출력: ~/Desktop/pptx_report_demo/
  01_original.pptx
  02_filled.pptx       (표지 fill_form + KPI/실적/계획 set_cell)
"""
from __future__ import annotations

import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from document_adapter import load

FIXTURE = ROOT / "tests" / "fixtures" / "pptx" / "real" / "quarterly_report_synth.pptx"
OUT = Path.home() / "Desktop" / "pptx_report_demo"


def main() -> int:
    OUT.mkdir(exist_ok=True)
    for old in OUT.glob("*.pptx"):
        old.unlink()

    original = OUT / "01_original.pptx"
    shutil.copy2(FIXTURE, original)
    print(f"[1] 원본 → {original.name}")

    filled = OUT / "02_filled.pptx"
    shutil.copy2(FIXTURE, filled)
    a = load(filled)
    try:
        # 표지 양식 (T0) — fill_form auto (빈 값 셀에 채움)
        result = a.fill_form({
            "보고일자": "2026-04-17",
            "작성자": "홍길동",
            "작성부서": "경영기획팀",
            "승인자": "박부장",
        })
        print(f"\n[2-a] 표지 fill_form: {len(result['filled'])}/4")
        for f in result["filled"]:
            print(
                f"  {f['label']:8s} → T{f['table_index']}({f['row']},{f['col']}) "
                f"[{f['action']}]"
            )

        # KPI (T2, 6x4) — set_cell 반복. "목표"/"실적"/"달성률" 열 채움
        print("\n[2-b] KPI 표 (T2) set_cell:")
        kpi_data = [
            # (row, (목표, 실적, 달성률))
            (1, ("100억", "112억", "112%"), "매출"),
            (2, ("15억", "18억", "120%"), "영업이익"),
            (3, ("500명", "580명", "116%"), "신규고객"),
            (4, ("8.5/10", "9.1/10", "107%"), "고객만족도"),
            (5, ("3건", "4건", "133%"), "제품출시"),
        ]
        for row, values, name in kpi_data:
            for col, val in enumerate(values, start=1):
                a.set_cell(2, row, col, val)
            print(f"  T2 row {row} ({name}): 목표/실적/달성률 채움")

        # 부문별 실적 (T3, 6x6, 2 병합) — set_cell
        print("\n[2-c] 부문별 실적 (T3) set_cell:")
        # 부문 아래 빈칸 (row 1 col 0) 에 "부문" 라벨 추가
        a.set_cell(3, 1, 0, "부문")
        # 사업별 매출 (2Q/3Q/4Q/Q1목표/Q1실적)
        divisions = [
            ("사업A", "25억", "28억", "32억", "35억", "38억"),
            ("사업B", "18억", "20억", "22억", "23억", "25억"),
            ("사업C", "12억", "14억", "16억", "17억", "19억"),
            ("합계", "55억", "62억", "70억", "75억", "82억"),
        ]
        for row_idx, row_data in enumerate(divisions, start=2):
            for col_idx, val in enumerate(row_data):
                a.set_cell(3, row_idx, col_idx, val)
        print(f"  T3 사업A/B/C/합계 × (2Q/3Q/4Q/Q1목표/Q1실적) 20 셀 채움")
        print("  병합 유지: (0,1) 2025년 (3열), (0,4) 2026 Q1 (2열)")

        # 다음 분기 계획 (T4, 5x4)
        print("\n[2-d] 다음 분기 계획 (T4) set_cell:")
        plans = [
            ("1", "신제품 X 런칭", "홍길동", "2026-05-31"),
            ("2", "해외 진출 리서치", "박부장", "2026-06-15"),
            ("3", "CS 시스템 개선", "김차장", "2026-06-30"),
            ("4", "파트너 미팅 4건", "이과장", "2026-06-30"),
        ]
        for r, plan in enumerate(plans, start=1):
            for c, val in enumerate(plan):
                a.set_cell(4, r, c, val)
        print(f"  T4 우선순위 1~4 × (과제/담당자/완료일) 채움")

        a.save(filled)
    finally:
        a.close()

    print(f"\n결과: {OUT}")
    for p in sorted(OUT.glob("*.pptx")):
        print(f"  {p.name}: {p.stat().st_size:,} bytes")
    print(f"\nKeynote: open -a Keynote '{filled}'")
    return 0


if __name__ == "__main__":
    sys.exit(main())
