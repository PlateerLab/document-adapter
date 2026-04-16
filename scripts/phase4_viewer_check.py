"""Phase 4: Hancom Office HWP Viewer 호환성 수동 검증 준비 (v2).

stop_payment_blank.hwpx (40KB, 1 table 28x16, 57 merges) 양식을 실제 사용
시나리오로 채운다:

  시나리오 A — 별도 값 셀에 set_cell (라벨은 유지)
    (3,3) 접수번호 값 셀 ← "2026-0001"
    (3,7) 접수일자 값 셀 ← "2026-04-16"

  시나리오 B — 라벨+값이 한 셀인 경우 append_to_cell
    (5,1) "성 명" → "성 명  홍길동"
    (11,1) "금융회사" → "금융회사  국민은행"

결과물을 ~/Desktop/hwpx_viewer_check/ 에 저장. Viewer에서 원본과 비교.
"""
from __future__ import annotations

import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from document_adapter import load

FIXTURE = ROOT / "tests" / "fixtures" / "hwpx" / "real" / "stop_payment_blank.hwpx"
OUT = Path.home() / "Desktop" / "hwpx_viewer_check"
VIEWER_APP = "Hancom Office HWP Viewer"


def main() -> int:
    if not FIXTURE.exists():
        print(f"fixture 없음: {FIXTURE}")
        return 2

    OUT.mkdir(exist_ok=True)
    # 이전 회차 결과 청소
    for old in OUT.glob("*.hwpx"):
        old.unlink()

    # 1) 원본
    original_dst = OUT / "01_original.hwpx"
    shutil.copy2(FIXTURE, original_dst)

    # 2) adapter 무수정 round-trip (회귀 기준)
    roundtrip_dst = OUT / "02_adapter_roundtrip_nochange.hwpx"
    a = load(FIXTURE)
    try:
        a.save(roundtrip_dst)
    finally:
        a.close()

    # 3) 양식 실제 사용 시나리오
    edit_dst = OUT / "03_adapter_form_filled.hwpx"
    shutil.copy2(FIXTURE, edit_dst)
    a = load(edit_dst)
    try:
        # --- 시나리오 A: 별도 값 셀에 set_cell ---
        scenario_a = [
            (0, 3, 3, "2026-0001"),       # 접수번호 값
            (0, 3, 7, "2026-04-16"),      # 접수일자 값
        ]
        print("시나리오 A (set_cell on 값 셀):")
        for tidx, r, c, value in scenario_a:
            old = a.set_cell(tidx, r, c, value)
            print(f"  set_cell(T{tidx},{r},{c}): {old!r} → {value!r}")

        # --- 시나리오 B: 라벨+값 한 셀, append_to_cell ---
        scenario_b = [
            (0, 5, 1, "홍길동"),          # 성명
            (0, 11, 1, "국민은행"),       # 지급정지요청계좌 금융회사
        ]
        print("\n시나리오 B (append_to_cell on 라벨 셀):")
        for tidx, r, c, value in scenario_b:
            old = a.append_to_cell(tidx, r, c, value)
            print(f"  append_to_cell(T{tidx},{r},{c}): {old!r} + '  {value}'")

        a.save(edit_dst)
    finally:
        a.close()

    print(f"\n파일 크기:")
    for p in sorted(OUT.glob("*.hwpx")):
        print(f"  {p.name}: {p.stat().st_size:,} bytes")

    print(f"\nFinder로 열기:")
    print(f"  open '{OUT}'")
    print(f"\n또는 Viewer로 직접 (터미널에서):")
    print(f"  open -a '{VIEWER_APP}' '{edit_dst}'")
    return 0


if __name__ == "__main__":
    sys.exit(main())
