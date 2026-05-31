"""LLM 주도 평가 하니스.

도구가 안 깨지는지(단위 테스트)를 넘어, 실제 모델이 도구로 작업을 올바르게
해내는지를 결과 기반으로 채점한다. 백엔드(LLM)는 주입식이라 하니스 자체는
LLM 없이 결정적으로 검증 가능하다.
"""
from __future__ import annotations

from document_adapter.eval.harness import (
    SYSTEM,
    EvalSummary,
    FieldExpectation,
    FieldResult,
    ModelBackend,
    ModelStep,
    Scenario,
    ScenarioResult,
    ToolCall,
    evaluate,
    run_scenario,
)

__all__ = [
    "SYSTEM",
    "EvalSummary",
    "FieldExpectation",
    "FieldResult",
    "ModelBackend",
    "ModelStep",
    "Scenario",
    "ScenarioResult",
    "ToolCall",
    "evaluate",
    "run_scenario",
]
