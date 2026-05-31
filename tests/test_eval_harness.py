"""LLM 평가 하니스의 결정적 검증 (실제 모델/API 불필요).

스크립트된 가짜 백엔드로 채점 로직을 검증한다:
  - 정답을 채우는 백엔드 → passed=True, score=1.0
  - 틀린 값 → score < 1.0, passed=False
  - 라벨 셀을 덮어쓰는 백엔드 → corrupted_labels 검출, passed=False
  - evaluate() 가 N회 실행의 pass-rate 를 올바로 집계

이 테스트가 통과하면, 키/Ollama 만 꽂으면 실제 모델 평가가 같은 채점으로
돌아간다는 것이 보장된다.

    pytest tests/test_eval_harness.py -v
"""
from __future__ import annotations

from pathlib import Path

import pytest

from document_adapter.eval import ModelStep, ToolCall, evaluate, run_scenario
from document_adapter.eval.scenarios import ALL_SCENARIOS, HARD_SCENARIOS

SCEN = {s.name: s for s in ALL_SCENARIOS}


class ScriptedBackend:
    """미리 정한 ModelStep 시퀀스를 순서대로 재생 (history 무시).

    실제 LLM 자리에 꽂아 하니스를 결정적으로 검증하는 용도.
    """
    def __init__(self, steps: list[ModelStep]) -> None:
        self._steps = list(steps)
        self._i = 0

    def next_step(self, history: list[dict], tools: list[dict]) -> ModelStep:
        if self._i >= len(self._steps):
            return ModelStep(final_text="done")
        step = self._steps[self._i]
        self._i += 1
        return step


def _fill_form_step(data: dict, **kw) -> ModelStep:
    args = {"data": data}
    args.update(kw)
    return ModelStep(tool_calls=[ToolCall("fill_form", args)])


# ---------------------------------------------------------------------------

def test_correct_backend_passes(tmp_path: Path) -> None:
    """정답 fill_form → 모든 필드 통과, 라벨 무손상."""
    scen = SCEN["A hwpx blank auto"]
    backend = ScriptedBackend([
        _fill_form_step({"접수번호": "2026-0001", "성명": "홍길동",
                         "주소": "서울시 강남구"}),
    ])
    res = run_scenario(scen, backend, workdir=tmp_path)
    assert res.passed, res.summary()
    assert res.score == 1.0
    assert not res.corrupted_labels
    assert "fill_form" in res.tool_sequence


def test_wrong_values_fail(tmp_path: Path) -> None:
    """틀린 값을 채우면 score<1, passed=False."""
    scen = SCEN["A hwpx blank auto"]
    backend = ScriptedBackend([
        _fill_form_step({"접수번호": "WRONG", "성명": "엉뚱이",
                         "주소": "서울시 강남구"}),
    ])
    res = run_scenario(scen, backend, workdir=tmp_path)
    assert not res.passed
    assert res.score < 1.0
    # 주소만 맞았으므로 1/3
    assert abs(res.score - 1 / 3) < 1e-9, res.summary()


def test_label_corruption_detected(tmp_path: Path) -> None:
    """라벨 셀(보호 대상)을 직접 덮어쓰면 corrupted 로 검출되고 fail."""
    scen = SCEN["A hwpx blank auto"]
    backend = ScriptedBackend([
        # set_cell 로 라벨 칸(0,0,0)을 파괴 + 값은 right 로 정상 기입
        ModelStep(tool_calls=[
            ToolCall("set_cell", {"table_index": 0, "row": 0, "col": 0,
                                  "value": "파괴됨"}),
            ToolCall("set_cell", {"table_index": 0, "row": 0, "col": 1,
                                  "value": "2026-0001"}),
            ToolCall("set_cell", {"table_index": 0, "row": 1, "col": 1,
                                  "value": "홍길동"}),
            ToolCall("set_cell", {"table_index": 0, "row": 2, "col": 1,
                                  "value": "서울시 강남구"}),
        ]),
    ])
    res = run_scenario(scen, backend, workdir=tmp_path)
    assert (0, 0, 0) in res.corrupted_labels, res.summary()
    assert not res.passed  # 값은 맞아도 라벨 파괴면 실패


def test_noop_backend_fails(tmp_path: Path) -> None:
    """아무 도구도 안 부르면 빈 칸 그대로 → 전 필드 실패."""
    scen = SCEN["A hwpx blank auto"]
    res = run_scenario(scen, ScriptedBackend([]), workdir=tmp_path)
    assert not res.passed
    assert res.score == 0.0
    assert res.tool_sequence == []


def test_dot_path_multi_section(tmp_path: Path) -> None:
    """중복 라벨 섹션을 dot-path 로 구분해 각각 올바른 셀에 채움."""
    scen = SCEN["C hwpx dup-section dot-path"]
    backend = ScriptedBackend([
        _fill_form_step({"피해자.금액": "1,000,000",
                         "지급정지.금액": "2,000,000"}),
    ])
    res = run_scenario(scen, backend, workdir=tmp_path)
    assert res.passed, res.summary()


def test_evaluate_aggregates_pass_rate(tmp_path: Path) -> None:
    """evaluate() 가 좋은/나쁜 백엔드 혼합 시 pass-rate 를 집계.

    make_backend 가 호출 횟수에 따라 정답/오답을 번갈아 내도록 해
    pass-rate=0.5 가 나오는지 확인 (집계 로직 검증).
    """
    scen = SCEN["A hwpx blank auto"]
    correct = {"접수번호": "2026-0001", "성명": "홍길동", "주소": "서울시 강남구"}
    wrong = {"접수번호": "X", "성명": "X", "주소": "X"}
    counter = {"n": 0}

    def make_backend():
        i = counter["n"]
        counter["n"] += 1
        data = correct if i % 2 == 0 else wrong
        return ScriptedBackend([_fill_form_step(data)])

    summary = evaluate(scen, make_backend, workdir=tmp_path, runs=4)
    assert summary.runs == 4
    assert summary.pass_rate == 0.5, summary
    # 정답 2회(score 1.0) + 오답 2회(3필드 모두 틀려 score 0.0) → 평균 0.5
    assert abs(summary.mean_score - 0.5) < 1e-9, summary


# ---------------------------------------------------------------------------
# HARD 시나리오 golden 도달 가능성 (오라클 백엔드로 정답 tool-call 재생)
# 모델 없이도 "이 fixture+golden 은 fill_form 으로 풀 수 있다"를 보장한다.
# ---------------------------------------------------------------------------

_ORACLE_FILL = {
    "H1 hwpx header-below":
        {"성명": "홍길동", "부서": "개발팀", "직급": "책임"},
    "H2 hwpx wide-grid":
        {"성명": "김철수", "생년월일": "1990-01-01",
         "주소": "서울시 강남구", "연락처": "010-1111-2222"},
    "H3 hwpx merged-section-header":
        {"성명": "홍길동", "연락처": "010-1234-5678"},
    "H4 hwpx three-section dot-path":
        {"신청인.성명": "홍길동", "대리인.성명": "김대리",
         "보증인.성명": "이보증"},
}


@pytest.mark.parametrize("scen_name", [s.name for s in HARD_SCENARIOS])
def test_hard_scenario_solvable_by_oracle(tmp_path: Path, scen_name: str) -> None:
    """정답 fill_form 한 번으로 HARD 시나리오의 golden 이 충족돼야 한다."""
    scen = SCEN[scen_name]
    backend = ScriptedBackend([_fill_form_step(_ORACLE_FILL[scen_name])])
    res = run_scenario(scen, backend, workdir=tmp_path)
    assert res.passed, res.summary()
    assert res.score == 1.0
    assert not res.corrupted_labels
