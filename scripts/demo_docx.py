"""DOCX adapter 시각 시연 — 원본 vs 편집본 비교용.

두 시나리오 준비:
  Part 1: employee_form.docx (합성 양식, 병합 없음) — fill_form 기본 시연
    01_form_original.docx
    02_form_roundtrip.docx       (adapter load → save, 무수정)
    03_form_filled.docx          (fill_form 7개 라벨)

  Part 2: loan_products_detail.docx (실전, 48 merges) — 병합 처리 시연
    04_merged_original.docx
    05_merged_anchor_edit.docx   (병합 anchor 편집)
    06_merged_redirect.docx      (non-anchor + allow_merge_redirect)
"""
from __future__ import annotations

import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from document_adapter import load

OUT = Path.home() / "Desktop" / "docx_demo"
FIXTURES = ROOT / "tests" / "fixtures" / "docx" / "real"


def part1_form() -> None:
    print("=" * 70)
    print("Part 1: 합성 양식 (employee_form.docx) — fill_form 시연")
    print("=" * 70)

    original = OUT / "01_form_original.docx"
    shutil.copy2(FIXTURES / "employee_form.docx", original)
    print(f"\n[1] 원본 → {original.name}")

    # adapter 무수정 round-trip
    roundtrip = OUT / "02_form_roundtrip.docx"
    a = load(original)
    try:
        a.save(roundtrip)
    finally:
        a.close()
    print(f"[2] adapter 무수정 round-trip → {roundtrip.name}")

    # fill_form 으로 채우기
    filled = OUT / "03_form_filled.docx"
    shutil.copy2(original, filled)
    a = load(filled)
    try:
        result = a.fill_form({
            "성명": "홍길동",
            "부서": "개발팀",
            "직급": "책임",
            "입사일": "2026-04-17",
            "전화번호": "010-1234-5678",
            "이메일": "hong@example.com",
            "주소": "서울시 강남구 역삼동 123-4",
        })
        a.save(filled)
    finally:
        a.close()
    print(f"[3] fill_form → {filled.name}")
    for f in result["filled"]:
        print(
            f"     {f['label']:7s} → T{f['table_index']}({f['row']},{f['col']}) "
            f"[{f['action']:16s}] '{f['old_value']}' → '{f['new_value']}'"
        )


def part2_merged() -> None:
    print()
    print("=" * 70)
    print("Part 2: 실전 복잡 병합 (loan_products_detail.docx, 48 merges)")
    print("=" * 70)

    original = OUT / "04_merged_original.docx"
    shutil.copy2(FIXTURES / "loan_products_detail.docx", original)
    print(f"\n[4] 원본 → {original.name} (48 merges, 53 tables)")

    # 병합 anchor 편집 (T4 anchor=(0,0) span=(2,1))
    anchor_edit = OUT / "05_merged_anchor_edit.docx"
    shutil.copy2(original, anchor_edit)
    a = load(anchor_edit)
    try:
        # 실제 병합 있는 표 찾기
        tables = a.get_tables(preview_rows=5, max_cell_len=20)
        target_table = None
        for t in tables:
            if t.merges:
                target_table = t
                break
        if target_table is None:
            print("    병합 있는 표 없음")
            return
        m = target_table.merges[0]
        ar, ac = m.anchor
        rs, cs = m.span
        print(f"\n[5] T{target_table.index} anchor=({ar},{ac}) span={m.span} — 병합 anchor 편집")
        old = a.set_cell(target_table.index, ar, ac, "★ 수정된 anchor (span 유지) ★")
        print(f"     anchor set_cell: '{old}' → '★ 수정된 anchor (span 유지) ★'")
        a.save(anchor_edit)
    finally:
        a.close()

    # non-anchor + allow_merge_redirect (같은 파일 이어서 편집)
    redirect = OUT / "06_merged_redirect.docx"
    shutil.copy2(original, redirect)
    a = load(redirect)
    try:
        tables = a.get_tables(preview_rows=5, max_cell_len=20)
        target_table = None
        for t in tables:
            if t.merges:
                target_table = t
                break
        m = target_table.merges[0]
        ar, ac = m.anchor
        rs, cs = m.span
        # non-anchor 좌표 (span 내부)
        non_r = ar + (rs - 1) if rs > 1 else ar
        non_c = ac + (cs - 1) if cs > 1 else ac
        print(
            f"\n[6] T{target_table.index} non-anchor ({non_r},{non_c}) + "
            f"allow_merge_redirect=True"
        )
        # 먼저 non-anchor 에 그냥 시도 → 에러 기대
        try:
            a.set_cell(target_table.index, non_r, non_c, "X")
            print(f"     ⚠️  non-anchor 쓰기가 에러 없이 통과 (예상 밖)")
        except Exception as e:
            print(f"     non-anchor 쓰기 거부 (예상): {type(e).__name__}")
        # allow_merge_redirect=True 로 재시도
        old = a.set_cell(
            target_table.index, non_r, non_c,
            "★ redirect 로 anchor 에 쓰임 ★",
            allow_merge_redirect=True,
        )
        print(f"     allow_merge_redirect=True: anchor 로 redirect, old={old!r}")
        a.save(redirect)
    finally:
        a.close()


def main() -> int:
    OUT.mkdir(exist_ok=True)
    for old in OUT.glob("*.docx"):
        old.unlink()

    part1_form()
    part2_merged()

    print()
    print("=" * 70)
    print("결과 파일 (Word, Pages, LibreOffice 로 열어 비교)")
    print("=" * 70)
    for p in sorted(OUT.glob("*.docx")):
        print(f"  {p.name:40s} {p.stat().st_size:>8,} bytes")
    print(f"\nFinder: open '{OUT}'")
    return 0


if __name__ == "__main__":
    sys.exit(main())
