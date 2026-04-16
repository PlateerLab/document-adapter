"""PPTX adapter 실사용 시연 v2 — v0.6.0 셀 크기 메타 활용.

ai_plan_small.pptx 를 LLM 관점에서 편집:
  Step 1. inspect_document 로 각 셀의 column_widths_cm, row_heights_cm 확인
  Step 2. 각 셀 스케일에 맞는 값 길이 선택
  Step 3. set_cell / append_row 실행

이전 v1 시연에서는 T0(1.7×0.7cm 배지, char=4)에 22자 긴 제목을 넣어 5줄 접힘,
T1 비고(3cm×0.6cm)에 37자 긴 텍스트 넣어 윗 행과 시각 겹침 발생.
v2 에서는 metadata 를 따라 값 길이를 조정.

단계:
  01_original                    원본
  02_adapter_roundtrip_nochange  무수정 round-trip
  03_adapter_fully_filled        크기-aware 편집 (모든 셀 적정 길이)
"""
from __future__ import annotations

import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from document_adapter import load

FIXTURE = ROOT / "tests" / "fixtures" / "pptx" / "real" / "ai_plan_small.pptx"
OUT = Path.home() / "Desktop" / "pptx_demo"


def main() -> int:
    if not FIXTURE.exists():
        print(f"fixture 없음: {FIXTURE}")
        return 2

    OUT.mkdir(exist_ok=True)
    for old in OUT.glob("*.pptx"):
        old.unlink()

    # Step 1: LLM 처럼 inspect 해서 각 셀 메타 확인
    print("=" * 70)
    print("Step 1: inspect_document — LLM 이 보는 메타")
    print("=" * 70)
    a = load(FIXTURE)
    try:
        tables = a.get_tables(preview_rows=2)
        for t in tables:
            widths = t.column_widths_cm or []
            heights = t.row_heights_cm or []
            # 각 셀의 크기 (간단히 col 기준 폭)
            total_w = sum(widths) if widths else None
            max_h = max(heights) if heights else None
            preview_vals = []
            for row in t.preview[:1]:
                for v in row:
                    if v is not None:
                        preview_vals.append(repr(v)[:30])
            print(
                f"T{t.index}: {t.rows}x{t.cols}  "
                f"widths={widths} heights={heights}  "
                f"→ 추정 총 폭 {total_w}cm, 행 높이 최대 {max_h}cm"
            )
            if preview_vals:
                print(f"   (0,*): {', '.join(preview_vals)}")
    finally:
        a.close()

    # 1) 원본
    original = OUT / "01_original.pptx"
    shutil.copy2(FIXTURE, original)

    # 2) round-trip
    roundtrip = OUT / "02_adapter_roundtrip_nochange.pptx"
    a = load(FIXTURE)
    try:
        a.save(roundtrip)
    finally:
        a.close()

    # 3) 크기-aware 편집
    filled = OUT / "03_adapter_fully_filled.pptx"
    shutil.copy2(FIXTURE, filled)

    print()
    print("=" * 70)
    print("Step 2-3: 크기에 맞춘 값 선택 + 편집")
    print("=" * 70)

    a = load(filled)
    try:
        # --- T0 (1.7×0.7cm, char=4 — 작은 배지): 6자 이내 ---
        c0 = a.get_cell(0, 0, 0)
        print(f"T0 (1.7cm 배지, char={c0.char_count}): '{c0.text}' → 'AI 기획서' (5자)")
        a.set_cell(0, 0, 0, "AI 기획서")

        # --- T1 (3x2, 각 값 셀 ≈3cm 폭): 10자 내외 ---
        c1 = a.get_cell(1, 0, 1)
        print(f"T1(0,1) (3cm 값 셀, char={c1.char_count}): '{c1.text}' → '2026.04.16' (10자)")
        a.set_cell(1, 0, 1, "2026.04.16")

        c1 = a.get_cell(1, 1, 1)
        print(f"T1(1,1) (3cm 값 셀, char={c1.char_count}): '{c1.text}' → '홍길동' (3자)")
        a.set_cell(1, 1, 1, "홍길동")

        c1 = a.get_cell(1, 2, 1)
        print(f"T1(2,1) (3cm 값 셀, char={c1.char_count}): '{c1.text}' → '개발팀' (3자)")
        a.set_cell(1, 2, 1, "개발팀")

        # --- T1 append_row: 3cm 값 셀이므로 짧게 ---
        print("T1 append_row: ['승인자', '박부장'] (각 3자 — 3cm 셀에 적합)")
        a.append_row(1, ["승인자", "박부장"])

        print("T1 append_row: ['비고', 'v0.6 시연'] (2자/8자 — 오버플로 없음)")
        a.append_row(1, ["비고", "v0.6 시연"])

        # --- T2 (16.1×9.0cm, char=58 — 넓은 자유서술): 길게 OK ---
        c2 = a.get_cell(2, 0, 0)
        print(f"T2 ({c2.width_cm}×{c2.height_cm}cm, char={c2.char_count}): 140자 긴 주제 설명")
        a.set_cell(2, 0, 0,
            "GenAI 기반 문서 자동 편집 에이전트 개발 — DOCX/PPTX/HWPX 양식 문서의 "
            "셀 / 표 / 행을 LLM 도구 호출로 채우고, 한글(HWPX) 네이티브 지원을 "
            "차별점으로 삼는 MCP 도구 스위트 (document-adapter) 구축.")

        # --- T3-T5 (17.4×6.3cm 자유서술): 120자 내외 ---
        c3 = a.get_cell(3, 0, 0)
        print(f"T3 ({c3.width_cm}×{c3.height_cm}cm): 120자 업무 활용 시나리오")
        a.set_cell(3, 0, 0,
            "업무 시작: 채팅 업로드 → inspect_document 로 표 구조 파악. "
            "업무 중간: set_cell / append_to_cell / append_row 반복 호출로 값 채움. "
            "업무 완료: [File: ...] [Path: ...] 마커로 편집본 전달.")

        c4 = a.get_cell(4, 0, 0)
        print(f"T4 ({c4.width_cm}×{c4.height_cm}cm): 운영 및 모니터링")
        a.set_cell(4, 0, 0,
            "세션 단위 threading.Lock 으로 병렬 호출 직렬화 → chain 누적 보장. "
            "저장 시 자동 suffix ( (2), (3) … ) 로 버전 보존. "
            "4-Stage regression harness 로 편집 전후 메트릭·invariance 자동 감지.")

        c5 = a.get_cell(5, 0, 0)
        print(f"T5 ({c5.width_cm}×{c5.height_cm}cm): 예외·리스크 대응")
        a.set_cell(5, 0, 0,
            "병합 셀 non-anchor 쓰기 → MergedCellWriteError 명시적 거부. "
            "확장자 불일치 → 지원 포맷 목록 반환. "
            "immutable 버킷 쓰기 시도 → storage 폴더 설정 안내. "
            "모든 에러는 LLM self-correct 가능 형식.")

        a.save(filled)
    finally:
        a.close()

    print()
    print("=" * 70)
    print("결과")
    print("=" * 70)
    for p in sorted(OUT.glob("*.pptx")):
        print(f"  {p.name}: {p.stat().st_size:,} bytes")
    print(f"\nFinder: open '{OUT}'")
    return 0


if __name__ == "__main__":
    sys.exit(main())
