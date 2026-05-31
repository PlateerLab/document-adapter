"""LLM 주도(outcome-based) 검증 하니스.

단위 테스트가 "도구가 안 깨진다"를 보장한다면, 이 하니스는 그 다음 질문에
답한다: **실제 모델이 이 도구들로 문서 작업을 올바르게 해내는가.**

설계:
  - ``ModelBackend`` : 모델 한 스텝(대화 → 다음 행동)을 추상화한 Protocol.
    실제 LLM(Claude/Ollama)을 꽂거나, 테스트용 스크립트 백엔드를 꽂는다 —
    그래서 하니스 로직 자체는 LLM 없이 결정적으로 검증할 수 있다.
  - ``run_scenario`` : 에이전트 루프를 돌린 뒤, **끝난 문서를 직접 열어**
    golden 기대값과 대조해 채점한다. "fill_form 을 호출했다"가 아니라
    "성명 칸에 홍길동이 실제로 들어갔다"를 본다.
  - ``evaluate`` : 같은 시나리오를 N 회 돌려 pass-rate 를 낸다 — LLM 의
    비결정성(temperature) 때문에 단일 실행은 신뢰할 수 없다.

이 모듈은 LLM SDK 에 의존하지 않는다(백엔드가 주입됨). 실제 백엔드 구현은
``examples/`` 의 Claude/Ollama 러너를 참고.
"""
from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Protocol

from document_adapter import load
from document_adapter.base import _overflow_risk
from document_adapter.tools import TOOL_DEFINITIONS, call_tool

# 세 러너(examples/scripts)에 흩어져 있던 에이전트 system 프롬프트의 정식본.
SYSTEM = """당신은 DOCX / PPTX / HWPX 양식 문서를 편집하는 에이전트입니다.

⚠ **반드시 tools API 로 호출**하세요. 응답 텍스트에 JSON 코드블록이나 함수 호출
문법을 직접 적지 마세요 — 그건 호출되지 않습니다.

워크플로우:
1. `inspect_document` 로 구조 파악.
2. `fill_form` 1 회 호출을 우선 — 여러 셀을 한 번에 채움.
3. direction 선택:
   - 값 셀이 비어있는 양식 → direction 생략 (auto).
   - 기존 예시값이 있는 양식 → direction="right" 명시.
4. 같은 라벨이 여러 섹션에 있어 ambiguous 반환받으면 dot-path 재호출:
     fill_form({"피해자.금액": "1,000,000", "지급정지.금액": "2,000,000"})
5. output_path / path 는 user 메시지에 주어진 문서 경로를 그대로 사용.
"""

_WS_RE = re.compile(r"\s+")


def _norm(s: str) -> str:
    """값 비교용 정규화: 공백 접기 + strip."""
    return _WS_RE.sub(" ", (s or "")).strip()


# ---------------------------------------------------------------------------
# 모델 추상화 (pluggable)
# ---------------------------------------------------------------------------

@dataclass
class ToolCall:
    name: str
    arguments: dict


@dataclass
class ModelStep:
    """모델 한 스텝의 산출: 도구 호출들 또는 종료 텍스트."""
    tool_calls: list[ToolCall] = field(default_factory=list)
    final_text: str = ""


class ModelBackend(Protocol):
    """대화 이력 + 도구 정의를 받아 다음 행동을 돌려준다.

    history 는 중립 형식의 메시지 리스트(role: system/user/assistant/tool,
    content: str). 실제 백엔드는 이를 자신의 wire 포맷으로 변환해 LLM 을
    호출하고, 스크립트 백엔드는 무시하고 미리 정한 시퀀스를 재생한다.
    """
    def next_step(self, history: list[dict], tools: list[dict]) -> ModelStep: ...


# ---------------------------------------------------------------------------
# 시나리오 / 채점 결과
# ---------------------------------------------------------------------------

@dataclass
class FieldExpectation:
    """채점 단위: (table, row, col) 셀이 expected 값을 담아야 한다.

    match='contains'(기본): expected 가 셀 텍스트에 포함되면 OK
    (라벨 접두사 "성명: 홍길동" 류 허용). match='exact': 정규화 후 완전일치.
    """
    table_index: int
    row: int
    col: int
    expected: str
    match: str = "contains"


@dataclass
class Scenario:
    """공개 합성 fixture + 자연어 지시 + golden 기대 최종상태."""
    name: str
    fmt: str                                    # "docx" | "pptx" | "hwpx"
    build_fixture: Callable[[Path], None]
    instruction: str
    expectations: list[FieldExpectation]
    # 덮어쓰면 안 되는 라벨 셀들. 실행 후에도 원본 텍스트가 유지돼야 한다.
    protected_cells: list[tuple[int, int, int]] = field(default_factory=list)

    @property
    def slug(self) -> str:
        return _WS_RE.sub("_", self.name.strip())[:40] or "scenario"


@dataclass
class FieldResult:
    expectation: FieldExpectation
    actual: str
    ok: bool
    overflow: bool = False


@dataclass
class ScenarioResult:
    name: str
    passed: bool
    score: float                                # 0.0 ~ 1.0 (맞춘 필드 비율)
    fields: list[FieldResult]
    corrupted_labels: list[tuple[int, int, int]]
    tool_sequence: list[str]
    error: str | None = None

    def summary(self) -> str:
        mark = "PASS" if self.passed else "FAIL"
        bits = [f"{mark} score={self.score:.2f}", f"tools={self.tool_sequence}"]
        if self.corrupted_labels:
            bits.append(f"corrupted={self.corrupted_labels}")
        if self.error:
            bits.append(f"error={self.error}")
        return " | ".join(bits)


@dataclass
class EvalSummary:
    name: str
    runs: int
    pass_rate: float
    mean_score: float
    results: list[ScenarioResult]


# ---------------------------------------------------------------------------
# 실행 + 채점
# ---------------------------------------------------------------------------

def _value_match(actual: str, exp: FieldExpectation) -> bool:
    a, e = _norm(actual), _norm(exp.expected)
    if not e:
        return True
    return a == e if exp.match == "exact" else e in a


def _snapshot_cells(path: Path, cells: list[tuple[int, int, int]]) -> dict:
    snap: dict[tuple[int, int, int], str] = {}
    if not cells:
        return snap
    ad = load(path)
    try:
        for (ti, r, c) in cells:
            try:
                snap[(ti, r, c)] = ad.get_cell(ti, r, c).text
            except Exception:
                snap[(ti, r, c)] = ""
    finally:
        ad.close()
    return snap


def run_scenario(
    scenario: Scenario,
    backend: ModelBackend,
    *,
    workdir: Path,
    max_turns: int = 8,
) -> ScenarioResult:
    """시나리오 1회 실행 후 결과 문서를 채점한다."""
    workdir.mkdir(parents=True, exist_ok=True)
    path = workdir / f"{scenario.slug}.{scenario.fmt}"
    scenario.build_fixture(path)

    # 실행 전 라벨 스냅샷 (덮어쓰기 오염 검출용)
    protected_before = _snapshot_cells(path, scenario.protected_cells)

    history: list[dict] = [
        {"role": "system", "content": SYSTEM},
        {"role": "user", "content": f"문서: {path}\n\n요청: {scenario.instruction}"},
    ]
    tool_sequence: list[str] = []
    error: str | None = None

    try:
        for _ in range(max_turns):
            step = backend.next_step(history, TOOL_DEFINITIONS)
            if not step.tool_calls:
                history.append({"role": "assistant", "content": step.final_text})
                break
            history.append({
                "role": "assistant",
                "content": step.final_text,
                "tool_calls": [
                    {"name": tc.name, "arguments": tc.arguments}
                    for tc in step.tool_calls
                ],
            })
            for tc in step.tool_calls:
                args = dict(tc.arguments)
                args.setdefault("path", str(path))   # 경로 누락 시 보정
                result = call_tool(tc.name, args)
                tool_sequence.append(tc.name)
                history.append({
                    "role": "tool",
                    "content": json.dumps(result, ensure_ascii=False),
                })
    except Exception as e:  # 백엔드/루프 크래시도 결과로 기록
        error = f"{type(e).__name__}: {e}"

    return _score(scenario, path, tool_sequence, protected_before, error)


def _score(
    scenario: Scenario,
    path: Path,
    tool_sequence: list[str],
    protected_before: dict,
    error: str | None,
) -> ScenarioResult:
    fields: list[FieldResult] = []
    corrupted: list[tuple[int, int, int]] = []

    ad = load(path)
    try:
        for exp in scenario.expectations:
            try:
                cell = ad.get_cell(exp.table_index, exp.row, exp.col)
                actual, wcm = cell.text, cell.width_cm
            except Exception:
                actual, wcm = "", None
            # placement-aware: 값이 맞아도 칸을 넘쳐 깨지면(overflow) 통과 아님.
            of = _overflow_risk(actual, wcm)
            ok = _value_match(actual, exp) and not of
            fields.append(FieldResult(exp, actual, ok, overflow=of))

        for (ti, r, c), before in protected_before.items():
            try:
                after = ad.get_cell(ti, r, c).text
            except Exception:
                after = ""
            if _norm(after) != _norm(before):
                corrupted.append((ti, r, c))
    finally:
        ad.close()

    n = len(fields)
    n_ok = sum(1 for f in fields if f.ok)
    score = (n_ok / n) if n else (0.0 if error else 1.0)
    passed = (n > 0 and n_ok == n) and not corrupted and error is None
    return ScenarioResult(
        name=scenario.name,
        passed=passed,
        score=score,
        fields=fields,
        corrupted_labels=corrupted,
        tool_sequence=tool_sequence,
        error=error,
    )


def evaluate(
    scenario: Scenario,
    make_backend: Callable[[], ModelBackend],
    *,
    workdir: Path,
    runs: int = 10,
    max_turns: int = 8,
) -> EvalSummary:
    """같은 시나리오를 runs 회 돌려 pass-rate / 평균 score 를 집계.

    make_backend 는 매 실행마다 새 백엔드를 만드는 팩토리 — 스크립트 백엔드처럼
    내부 상태(턴 카운터)를 가진 경우 실행 간 격리하기 위함.
    """
    results: list[ScenarioResult] = []
    for i in range(runs):
        results.append(run_scenario(
            scenario, make_backend(),
            workdir=workdir / f"run{i}", max_turns=max_turns,
        ))
    passed = sum(1 for r in results if r.passed)
    mean_score = sum(r.score for r in results) / len(results) if results else 0.0
    return EvalSummary(
        name=scenario.name,
        runs=runs,
        pass_rate=passed / len(results) if results else 0.0,
        mean_score=mean_score,
        results=results,
    )
