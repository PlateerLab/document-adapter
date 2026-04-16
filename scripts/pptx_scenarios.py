"""PPTX 심화 시나리오 검증 — regression harness 보완.

기존 scripts/pptx_regression.py 가 bytes/lxml/adapter round-trip 4 스테이지로
빠른 drift 감지를 하는 반면, 여기는 LLM 이 실제로 만나는 편집 패턴 / 에러 흐름 /
invariance 를 심화 검증한다.

시나리오 매트릭스:

  E (Edit patterns)
    E1  first anchor 에 set_cell 1회
    E2  같은 표의 anchor 3개 연속 set_cell (chain 누적)
    E3  다른 표 2개에 set_cell (flat_index 순회)
    E4  값 있는 anchor 에 append_to_cell (라벨 보존)
    E5  같은 anchor 두 번 set_cell (last write wins)

  X (Error handling)
    X1  merged non-anchor 좌표 set_cell → MergedCellWriteError
    X2  범위 초과 좌표 set_cell → CellOutOfBoundsError
    X3  잘못된 table_index → TableIndexError

  I (Invariance, 1회 편집 후 part-level diff)
    I1  namelist 동일 (추가/삭제 없음)
    I2  ppt/slides/ 외 .xml 파트 변경 없음 (수정 대상 슬라이드만 변경)
    I3  ppt/media/* 모두 bytes-identical (이미지 무손상)
    I4  ppt/slideMasters/*, ppt/slideLayouts/* 모두 동일 (마스터/레이아웃 무손상)
    I5  ppt/theme/*, ppt/fonts/* 모두 동일

사용:
  python scripts/pptx_scenarios.py              # 전 fixture × 전 시나리오
  python scripts/pptx_scenarios.py --fixture NAME   # 특정 fixture만
"""
from __future__ import annotations

import argparse
import sys
import traceback
import zipfile
from dataclasses import dataclass, field
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
ROOT = SCRIPT_DIR.parent
FIXTURES = ROOT / "tests" / "fixtures" / "pptx" / "real"
WORK = SCRIPT_DIR / "_poc_out" / "pptx_scenarios"


# ---------- 결과 모델 ----------

@dataclass
class ScenarioResult:
    code: str
    description: str
    passed: bool
    notes: str = ""


@dataclass
class FixtureReport:
    name: str
    size_bytes: int
    scenarios: list[ScenarioResult] = field(default_factory=list)


# ---------- helpers ----------

def part_level_diff(original: Path, edited: Path) -> dict:
    """ZIP 파트별 bytes 비교. 어느 파트가 변경/추가/삭제됐는지 반환."""
    with zipfile.ZipFile(original) as zo, zipfile.ZipFile(edited) as ze:
        orig_names = set(zo.namelist())
        edit_names = set(ze.namelist())
        added = sorted(edit_names - orig_names)
        removed = sorted(orig_names - edit_names)
        changed = []
        for name in sorted(orig_names & edit_names):
            if zo.read(name) != ze.read(name):
                changed.append(name)
    return {"added": added, "removed": removed, "changed": changed}


def find_first_anchors(adapter, n: int = 3) -> list[tuple[int, int, int]]:
    """(table_index, row, col) 튜플 n개 반환. 첫 표에서 우선 찾되 모자라면 다음 표로."""
    out: list[tuple[int, int, int]] = []
    tables = adapter.get_tables(preview_rows=1000)
    for t in tables:
        for r, row in enumerate(t.preview):
            for c, val in enumerate(row):
                if val is not None:
                    out.append((t.index, r, c))
                    if len(out) >= n:
                        return out
    return out


def find_anchors_across_tables(adapter, n: int = 2) -> list[tuple[int, int, int]]:
    """n개 서로 다른 표에서 각 1개 anchor 반환."""
    out: list[tuple[int, int, int]] = []
    seen_tables: set[int] = set()
    tables = adapter.get_tables(preview_rows=1000)
    for t in tables:
        if t.index in seen_tables:
            continue
        for r, row in enumerate(t.preview):
            for c, val in enumerate(row):
                if val is not None:
                    out.append((t.index, r, c))
                    seen_tables.add(t.index)
                    break
            if t.index in seen_tables:
                break
        if len(out) >= n:
            break
    return out


def find_merged_non_anchor(adapter) -> tuple[int, int, int] | None:
    """병합 영역의 non-anchor 좌표 찾기. 없으면 None."""
    tables = adapter.get_tables(preview_rows=1000)
    for t in tables:
        if not t.merges:
            continue
        for m in t.merges:
            ar, ac = m.anchor
            rs, cs = m.span
            # (ar, ac+1), (ar+1, ac) 같은 non-anchor 좌표
            for dr in range(rs):
                for dc in range(cs):
                    if dr == 0 and dc == 0:
                        continue
                    return (t.index, ar + dr, ac + dc)
    return None


def find_table_dims(adapter) -> tuple[int, int, int] | None:
    """첫 표의 (index, rows, cols)."""
    tables = adapter.get_tables(preview_rows=1000)
    if not tables:
        return None
    t = tables[0]
    return (t.index, t.rows, t.cols)


# ---------- 시나리오 실행 ----------

def run_scenarios(fixture_path: Path) -> FixtureReport:
    from document_adapter import load
    from document_adapter.base import (
        CellOutOfBoundsError,
        MergedCellWriteError,
        TableIndexError,
    )

    report = FixtureReport(name=fixture_path.name, size_bytes=fixture_path.stat().st_size)
    WORK.mkdir(parents=True, exist_ok=True)

    def scen(code: str, desc: str, passed: bool, notes: str = "") -> None:
        report.scenarios.append(ScenarioResult(code, desc, passed, notes))

    # ---------- E1: single set_cell ----------
    dst = WORK / f"e1_{fixture_path.name}"
    try:
        a = load(fixture_path)
        try:
            targets = find_first_anchors(a, n=1)
            if not targets:
                scen("E1", "single set_cell", False, "no anchor")
                return report
            tidx, r, c = targets[0]
            sentinel = "__E1_SENTINEL__"
            a.set_cell(tidx, r, c, sentinel)
            a.save(dst)
        finally:
            a.close()
        a2 = load(dst)
        try:
            got = a2.get_cell(tidx, r, c).text.strip()
        finally:
            a2.close()
        ok = got == sentinel
        scen("E1", "single set_cell", ok, f"target=T{tidx}({r},{c}), got={got!r}")
    except Exception as e:
        scen("E1", "single set_cell", False, f"{type(e).__name__}: {e}")

    # ---------- I (Invariance): E1 편집을 기준으로 part-level diff ----------
    try:
        diff = part_level_diff(fixture_path, dst)

        # I1: namelist 동일
        i1_ok = not diff["added"] and not diff["removed"]
        scen("I1", "namelist 동일", i1_ok,
             f"added={diff['added']}, removed={diff['removed']}")

        # python-pptx save 가 관례적으로 덮어쓰는 파트 — 허용 대상
        # ([Content_Types]/docProps/notesSlides/notesMasters/_rels 는 매 저장마다 재생성)
        def _is_benign(name: str) -> bool:
            return (
                name == "[Content_Types].xml"
                or name.startswith("docProps/")
                or name.startswith("ppt/notesSlides/")
                or name.startswith("ppt/notesMasters/")
                or name.endswith(".rels")
            )

        # I2: slides/ 외 본문 XML 변경 없음
        changed_xml = [n for n in diff["changed"] if n.endswith(".xml")]
        non_slide_xml = [
            n for n in changed_xml
            if not n.startswith("ppt/slides/") and not _is_benign(n)
        ]
        slide_changes = [n for n in changed_xml if n.startswith("ppt/slides/")]
        i2_ok = len(non_slide_xml) == 0
        scen("I2", "slides/ 외 본문 XML 변경 없음 (benign 파트 제외)", i2_ok,
             f"slide_changes={len(slide_changes)}, other={non_slide_xml}")

        # I3: 미디어 무손상
        media = [n for n in diff["changed"] if n.startswith("ppt/media/")]
        scen("I3", "ppt/media/* bytes 동일", len(media) == 0,
             f"changed_media={media}" if media else "(미디어 없거나 전부 동일)")

        # I4: 마스터/레이아웃 **본문** 무손상 (_rels 제외, python-pptx가 재생성)
        master_layout = [
            n for n in diff["changed"]
            if (n.startswith("ppt/slideMasters/") or n.startswith("ppt/slideLayouts/"))
            and not n.endswith(".rels")
        ]
        scen("I4", "마스터/레이아웃 본문 XML 동일", len(master_layout) == 0,
             f"changed={master_layout}" if master_layout else "(본문 전부 동일)")

        # I5: 테마/폰트 무손상
        theme_fonts = [
            n for n in diff["changed"]
            if n.startswith("ppt/theme/") or n.startswith("ppt/fonts/")
        ]
        scen("I5", "theme/fonts 파트 동일", len(theme_fonts) == 0,
             f"changed={theme_fonts}" if theme_fonts else "(전부 동일)")

        # I6: 수정되지 않은 슬라이드의 XML도 동일해야 (정말 중요한 invariance)
        # set_cell 1회는 slide 1개만 건드려야 이상적
        slide_changes_count = len(slide_changes)
        i6_ok = slide_changes_count <= 1
        scen("I6", "수정 대상 외 slide XML 불변 (편집 1회 → 1 slide 변경)", i6_ok,
             f"slide_changes_count={slide_changes_count}")
    except Exception as e:
        scen("I*", "invariance diff 실행", False, f"{type(e).__name__}: {e}")

    # ---------- E2: 같은 표 3 anchor 연속 set_cell ----------
    dst = WORK / f"e2_{fixture_path.name}"
    try:
        a = load(fixture_path)
        try:
            targets = find_first_anchors(a, n=3)
            if len(targets) < 2:
                scen("E2", "chain set_cell 3회", False, f"anchors={len(targets)} < 2")
            else:
                for i, (tidx, r, c) in enumerate(targets):
                    a.set_cell(tidx, r, c, f"__E2_{i}__")
                a.save(dst)
        finally:
            a.close()
        if len(targets) >= 2:
            a2 = load(dst)
            try:
                all_ok = True
                details = []
                for i, (tidx, r, c) in enumerate(targets):
                    got = a2.get_cell(tidx, r, c).text.strip()
                    want = f"__E2_{i}__"
                    details.append(f"({r},{c})={got!r}")
                    if got != want:
                        all_ok = False
            finally:
                a2.close()
            scen("E2", "chain set_cell 3회", all_ok, "; ".join(details))
    except Exception as e:
        scen("E2", "chain set_cell 3회", False, f"{type(e).__name__}: {e}")

    # ---------- E3: 다른 표 anchor set_cell ----------
    dst = WORK / f"e3_{fixture_path.name}"
    try:
        a = load(fixture_path)
        try:
            targets = find_anchors_across_tables(a, n=2)
            if len(targets) < 2:
                scen("E3", "다른 표 2개 set_cell", False,
                     f"distinct tables={len(targets)} < 2")
            else:
                for i, (tidx, r, c) in enumerate(targets):
                    a.set_cell(tidx, r, c, f"__E3_T{tidx}__")
                a.save(dst)
        finally:
            a.close()
        if len(targets) >= 2:
            a2 = load(dst)
            try:
                all_ok = True
                details = []
                for tidx, r, c in targets:
                    got = a2.get_cell(tidx, r, c).text.strip()
                    want = f"__E3_T{tidx}__"
                    details.append(f"T{tidx}={got!r}")
                    if got != want:
                        all_ok = False
            finally:
                a2.close()
            scen("E3", "다른 표 2개 set_cell", all_ok, "; ".join(details))
    except Exception as e:
        scen("E3", "다른 표 2개 set_cell", False, f"{type(e).__name__}: {e}")

    # ---------- E4: append_to_cell ----------
    # 검증 기준: (1) old 텍스트가 결과에 그대로 포함 (2) suffix로 끝남
    # old trailing whitespace 처리는 어댑터별로 다를 수 있어 startswith/endswith 로 관대하게
    dst = WORK / f"e4_{fixture_path.name}"
    try:
        a = load(fixture_path)
        try:
            targets = find_first_anchors(a, n=1)
            tidx, r, c = targets[0]
            old = a.append_to_cell(tidx, r, c, "__E4_SUFFIX__", separator="  ")
            a.save(dst)
        finally:
            a.close()
        a2 = load(dst)
        try:
            got_raw = a2.get_cell(tidx, r, c).text
        finally:
            a2.close()
        old_stripped = (old or "").strip()
        got_stripped = got_raw.strip()
        has_old = old_stripped in got_stripped if old_stripped else True
        has_suffix = got_stripped.endswith("__E4_SUFFIX__")
        ok = has_old and has_suffix
        scen("E4", "append_to_cell (라벨 보존)", ok,
             f"old={old_stripped!r}, got={got_stripped!r}, has_old={has_old}, has_suffix={has_suffix}")
    except Exception as e:
        scen("E4", "append_to_cell (라벨 보존)", False, f"{type(e).__name__}: {e}")

    # ---------- E5: 같은 셀 두 번 set_cell ----------
    dst = WORK / f"e5_{fixture_path.name}"
    try:
        a = load(fixture_path)
        try:
            targets = find_first_anchors(a, n=1)
            tidx, r, c = targets[0]
            a.set_cell(tidx, r, c, "__E5_FIRST__")
            a.set_cell(tidx, r, c, "__E5_SECOND__")
            a.save(dst)
        finally:
            a.close()
        a2 = load(dst)
        try:
            got = a2.get_cell(tidx, r, c).text.strip()
        finally:
            a2.close()
        ok = got == "__E5_SECOND__"
        scen("E5", "same cell 2회 set_cell (last-write-wins)", ok, f"got={got!r}")
    except Exception as e:
        scen("E5", "same cell 2회 set_cell (last-write-wins)", False,
             f"{type(e).__name__}: {e}")

    # ---------- X1: merged non-anchor rejection ----------
    try:
        a = load(fixture_path)
        try:
            target = find_merged_non_anchor(a)
            if target is None:
                scen("X1", "merged non-anchor 거부", True, "(해당 fixture에 병합 없음)")
            else:
                tidx, r, c = target
                try:
                    a.set_cell(tidx, r, c, "should_fail")
                    scen("X1", "merged non-anchor 거부", False,
                         f"MergedCellWriteError 기대했으나 통과 T{tidx}({r},{c})")
                except MergedCellWriteError:
                    scen("X1", "merged non-anchor 거부", True,
                         f"정상 거부 T{tidx}({r},{c})")
        finally:
            a.close()
    except Exception as e:
        scen("X1", "merged non-anchor 거부", False, f"{type(e).__name__}: {e}")

    # ---------- X2: out of bounds ----------
    try:
        a = load(fixture_path)
        try:
            dims = find_table_dims(a)
            if not dims:
                scen("X2", "out of bounds 거부", True, "(표 없음)")
            else:
                tidx, rows, cols = dims
                try:
                    a.set_cell(tidx, rows + 10, cols + 10, "should_fail")
                    scen("X2", "out of bounds 거부", False, "예외 없이 통과")
                except CellOutOfBoundsError:
                    scen("X2", "out of bounds 거부", True, "정상 거부")
        finally:
            a.close()
    except Exception as e:
        scen("X2", "out of bounds 거부", False, f"{type(e).__name__}: {e}")

    # ---------- X3: invalid table_index ----------
    try:
        a = load(fixture_path)
        try:
            try:
                a.set_cell(9999, 0, 0, "should_fail")
                scen("X3", "invalid table_index 거부", False, "예외 없이 통과")
            except TableIndexError:
                scen("X3", "invalid table_index 거부", True, "정상 거부")
        finally:
            a.close()
    except Exception as e:
        scen("X3", "invalid table_index 거부", False, f"{type(e).__name__}: {e}")

    return report


def print_report(r: FixtureReport) -> None:
    passed = sum(1 for s in r.scenarios if s.passed)
    total = len(r.scenarios)
    header = f"{r.name}  [{passed}/{total}]"
    print(f"\n{'=' * 80}\n{header}\n{'=' * 80}")
    groups = {"E": "편집", "I": "Invariance", "X": "에러"}
    for prefix, label in groups.items():
        print(f"\n  [{label}]")
        for s in r.scenarios:
            if not s.code.startswith(prefix):
                continue
            mark = "✅" if s.passed else "❌"
            print(f"    {mark} {s.code:3}  {s.description}")
            if s.notes:
                print(f"           {s.notes[:200]}")


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--fixture", help="특정 fixture 이름만 실행")
    args = ap.parse_args()

    fixtures = sorted(FIXTURES.glob("*.pptx"))
    if args.fixture:
        fixtures = [p for p in fixtures if args.fixture in p.name]

    if not fixtures:
        print("fixture 없음")
        return 2

    reports: list[FixtureReport] = []
    for p in fixtures:
        try:
            r = run_scenarios(p)
        except Exception as e:
            r = FixtureReport(name=p.name, size_bytes=p.stat().st_size)
            r.scenarios.append(
                ScenarioResult("??", "run_scenarios crashed", False,
                               f"{type(e).__name__}: {e}\n{traceback.format_exc()[:300]}")
            )
        reports.append(r)
        print_report(r)

    # 요약
    print("\n" + "=" * 80)
    print("총 요약")
    print("=" * 80)
    all_codes = sorted({s.code for r in reports for s in r.scenarios})
    for code in all_codes:
        passed = sum(1 for r in reports for s in r.scenarios if s.code == code and s.passed)
        total = sum(1 for r in reports for s in r.scenarios if s.code == code)
        status = "✅" if passed == total else "❌"
        print(f"  {status} {code}: {passed}/{total}")

    grand_passed = sum(1 for r in reports for s in r.scenarios if s.passed)
    grand_total = sum(len(r.scenarios) for r in reports)
    print(f"\n전체: {grand_passed}/{grand_total}")
    return 0 if grand_passed == grand_total else 1


if __name__ == "__main__":
    sys.exit(main())
