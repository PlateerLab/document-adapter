"""v0.7 fill_form 시각 시연 — 라벨 기반 일괄 채우기.

이전 v0.6 이하에서는 LLM 이 각 셀마다:
    adapter.set_cell(1, 0, 1, "2026.04.16")   # table 1, (0,1) 이 값 셀이라 계산
    adapter.set_cell(1, 1, 1, "홍길동")
    adapter.set_cell(1, 2, 1, "개발팀")
세 번 호출 + table_index/row/col 을 매번 추정.

v0.7 부터:
    adapter.fill_form({
        "보고일자": "2026.04.16",
        "작성자": "홍길동",
        "담당부서": "개발팀",
    })
한 번 호출, 라벨 이름만 알면 됨.

출력:
  ~/Desktop/document_adapter_demo/
    hwpx_01_original.hwpx
    hwpx_02_fill_form.hwpx
    pptx_01_original.pptx
    pptx_02_fill_form.pptx
"""
from __future__ import annotations

import json
import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from document_adapter import load

OUT = Path.home() / "Desktop" / "document_adapter_demo"

HWPX_FIXTURE = ROOT / "tests" / "fixtures" / "hwpx" / "real" / "stop_payment_blank.hwpx"
PPTX_FIXTURE = ROOT / "tests" / "fixtures" / "pptx" / "real" / "ai_plan_small.pptx"


def demo_hwpx() -> None:
    print("=" * 70)
    print("HWPX 양식: 지급정지요청서 (stop_payment_blank.hwpx)")
    print("=" * 70)

    original = OUT / "hwpx_01_original.hwpx"
    shutil.copy2(HWPX_FIXTURE, original)

    filled = OUT / "hwpx_02_fill_form.hwpx"
    shutil.copy2(HWPX_FIXTURE, filled)

    # 한 번에 관련 라벨 모두 넘김 — 인접 라벨 오염 방지 효과
    data = {
        "접수번호": "2026-0001",
        "접수일자": "2026.04.16",
        "성 명": "홍길동",
        "생년월일": "1990.01.01",
        "주 소": "서울시 강남구 역삼동 123-4",
        "전화번호": "02-1234-5678",
        "휴대전화번호": "010-1234-5678",
        "전자우편주소": "hong@example.com",
        "금융회사": "국민은행",
        "예금종별": "보통예금",
        "계좌번호 및 명의인": "123-456-789 홍길동",
        "금액": "5,000,000원",
    }

    a = load(filled)
    try:
        result = a.fill_form(data)
        a.save()
    finally:
        a.close()

    print(f"\n한 번의 fill_form 호출로 {len(data)}개 라벨 요청:")
    print(f"  ✅ filled:    {len(result['filled'])}")
    print(f"  ⚠️  not_found: {len(result['not_found'])} {result['not_found']}")
    print(f"  ⚠️  ambiguous: {len(result['ambiguous'])}")
    print(f"\n채워진 라벨 → 좌표 + action:")
    for f in result["filled"]:
        action_mark = "→" if f["action"] == "set_cell" else "+"
        print(
            f"  {f['label']:14s} T{f['table_index']}({f['row']:>2},{f['col']:>2}) "
            f"[{f['action']:16s}] '{f['old_value']}' {action_mark} '{f['new_value']}'"
        )
    if result["not_found"]:
        print(f"\n못 찾은 라벨 (양식에 없음): {result['not_found']}")


def demo_pptx() -> None:
    print()
    print("=" * 70)
    print("PPTX 양식: AI 활용 기획서 (ai_plan_small.pptx)")
    print("=" * 70)

    original = OUT / "pptx_01_original.pptx"
    shutil.copy2(PPTX_FIXTURE, original)

    filled = OUT / "pptx_02_fill_form.pptx"
    shutil.copy2(PPTX_FIXTURE, filled)

    data = {
        "보고일자": "2026.04.16",
        "작성자": "홍길동",
        "담당부서": "개발팀",
    }

    a = load(filled)
    try:
        result = a.fill_form(data)
        a.save()
    finally:
        a.close()

    print(f"\nfill_form({len(data)}개 라벨):")
    print(f"  ✅ filled:    {len(result['filled'])}")
    print(f"  ⚠️  not_found: {result['not_found']}")
    print(f"\n채워진 라벨 → 좌표:")
    for f in result["filled"]:
        print(
            f"  {f['label']:10s} T{f['table_index']}({f['row']:>2},{f['col']:>2}) "
            f"[{f['action']:16s}] '{f['old_value']}' → '{f['new_value']}'"
        )


def main() -> int:
    OUT.mkdir(exist_ok=True)
    for old in OUT.glob("*.hwpx"):
        old.unlink()
    for old in OUT.glob("*.pptx"):
        old.unlink()

    demo_hwpx()
    demo_pptx()

    print()
    print("=" * 70)
    print(f"결과 파일: {OUT}")
    print("=" * 70)
    for p in sorted(OUT.iterdir()):
        print(f"  {p.name}: {p.stat().st_size:,} bytes")
    print(f"\nFinder: open '{OUT}'")
    print(f"\nHWPX 확인: open -a 'Hancom Office HWP Viewer' '{OUT}/hwpx_02_fill_form.hwpx'")
    print(f"PPTX 확인: open -a 'Keynote' '{OUT}/pptx_02_fill_form.pptx'")
    return 0


if __name__ == "__main__":
    sys.exit(main())
