#!/usr/bin/env python3
"""공개 시나리오를 실제 LLM 으로 돌려 pass-rate 를 측정하는 러너.

document_adapter.eval 하니스(결과 기반 채점) + 공개 합성 fixture 를 사용한다.
하니스 자체는 tests/test_eval_harness.py 에서 가짜 백엔드로 결정적으로
검증돼 있고, 여기서는 실제 모델 백엔드를 꽂는다.

사용:
    # Ollama (로컬, 무료) — ollama serve 가 떠 있어야 함
    python examples/eval_run.py --backend ollama --model qwen2.5:14b --runs 5

    # 특정 시나리오만
    python examples/eval_run.py --backend ollama --only "A hwpx" --runs 10

비결정성 때문에 runs>1 로 여러 번 돌려 pass-rate 를 보는 것을 권장한다.
"""
from __future__ import annotations

import argparse
import json
import tempfile
from pathlib import Path

from document_adapter.eval import ModelStep, ToolCall
from document_adapter.eval.harness import SYSTEM  # noqa: F401  (백엔드가 참조 가능)
from document_adapter.eval.scenarios import ALL_SCENARIOS


def _to_openai_tools(tool_defs: list[dict]) -> list[dict]:
    return [
        {"type": "function", "function": {
            "name": t["name"], "description": t["description"],
            "parameters": t["input_schema"],
        }}
        for t in tool_defs
    ]


def _history_to_ollama(history: list[dict]) -> list[dict]:
    """중립 history → Ollama(OpenAI 호환) 메시지 포맷."""
    out: list[dict] = []
    for m in history:
        role = m["role"]
        if role == "assistant" and m.get("tool_calls"):
            out.append({
                "role": "assistant",
                "content": m.get("content", ""),
                "tool_calls": [
                    {"function": {"name": tc["name"], "arguments": tc["arguments"]}}
                    for tc in m["tool_calls"]
                ],
            })
        else:
            out.append({"role": role, "content": m.get("content", "")})
    return out


class OllamaBackend:
    """ollama Client 를 ModelBackend 로 감싼 어댑터."""

    def __init__(self, model: str, host: str = "http://localhost:11434") -> None:
        from ollama import Client
        self._model = model
        self._client = Client(host=host)

    def next_step(self, history: list[dict], tools: list[dict]) -> ModelStep:
        resp = self._client.chat(
            model=self._model,
            messages=_history_to_ollama(history),
            tools=_to_openai_tools(tools),
            options={"num_predict": 1024, "temperature": 0.0},
        )
        msg = resp.message
        calls = getattr(msg, "tool_calls", None) or []
        if calls:
            return ModelStep(tool_calls=[
                ToolCall(c.function.name, dict(c.function.arguments or {}))
                for c in calls
            ])
        return ModelStep(final_text=getattr(msg, "content", "") or "")


def _history_to_openai(history: list[dict]) -> list[dict]:
    """중립 history → OpenAI chat 포맷 (tool_call_id 순차 페어링).

    assistant 의 tool_calls 에 call_{n} id 를 부여하고, 뒤따르는 tool 결과
    메시지에 같은 id 를 순서대로 매단다 (vLLM/OpenAI 가 요구).
    """
    out: list[dict] = []
    pending_ids: list[str] = []
    counter = 0
    for m in history:
        role = m["role"]
        if role == "assistant" and m.get("tool_calls"):
            tcs = []
            pending_ids = []
            for tc in m["tool_calls"]:
                cid = f"call_{counter}"
                counter += 1
                pending_ids.append(cid)
                tcs.append({
                    "id": cid, "type": "function",
                    "function": {"name": tc["name"],
                                 "arguments": json.dumps(
                                     tc["arguments"], ensure_ascii=False)},
                })
            out.append({"role": "assistant",
                        "content": m.get("content", "") or None,
                        "tool_calls": tcs})
        elif role == "tool":
            cid = pending_ids.pop(0) if pending_ids else "call_0"
            out.append({"role": "tool", "tool_call_id": cid,
                        "content": m.get("content", "")})
        else:
            out.append({"role": role, "content": m.get("content", "")})
    return out


class OpenAIBackend:
    """OpenAI 호환 /v1/chat/completions 백엔드 (vLLM·OpenAI·기타).

    vLLM 의 OpenAI 서버를 그대로 사용 — tool-calling 활성 서버 필요
    (예: --enable-auto-tool-choice --tool-call-parser ...).
    """

    def __init__(self, model: str, base_url: str,
                 api_key: str = "EMPTY") -> None:
        import requests
        self._model = model
        self._url = base_url.rstrip("/") + "/chat/completions"
        self._headers = {"Authorization": f"Bearer {api_key}",
                         "Content-Type": "application/json"}
        self._requests = requests

    def next_step(self, history: list[dict], tools: list[dict]) -> ModelStep:
        payload = {
            "model": self._model,
            "messages": _history_to_openai(history),
            "tools": _to_openai_tools(tools),
            "tool_choice": "auto",
            "temperature": 0.0,
            "max_tokens": 1024,
        }
        resp = self._requests.post(self._url, headers=self._headers,
                                   json=payload, timeout=120)
        resp.raise_for_status()
        msg = resp.json()["choices"][0]["message"]
        calls = msg.get("tool_calls") or []
        if calls:
            steps = []
            for c in calls:
                fn = c["function"]
                args = fn.get("arguments") or "{}"
                steps.append(ToolCall(
                    fn["name"],
                    json.loads(args) if isinstance(args, str) else args,
                ))
            return ModelStep(tool_calls=steps)
        return ModelStep(final_text=msg.get("content") or "")


def _make_backend(kind: str, model: str, base_url: str | None):
    if kind == "ollama":
        return lambda: OllamaBackend(model)
    if kind == "openai":
        if not base_url:
            raise SystemExit("--base-url 필요 (예: http://localhost:8012/v1)")
        return lambda: OpenAIBackend(model, base_url)
    raise SystemExit(f"unknown backend: {kind} (지원: ollama, openai)")


def main() -> int:
    ap = argparse.ArgumentParser(description="공개 시나리오 LLM 평가 러너")
    ap.add_argument("--backend", default="ollama", choices=["ollama", "openai"])
    ap.add_argument("--model", default="qwen2.5:14b")
    ap.add_argument("--base-url", default=None,
                    help="openai 백엔드용 (예: http://localhost:8012/v1)")
    ap.add_argument("--runs", type=int, default=5)
    ap.add_argument("--only", default=None, help="이름 prefix 필터")
    args = ap.parse_args()

    from document_adapter.eval import evaluate

    scenarios = [s for s in ALL_SCENARIOS
                 if not args.only or s.name.startswith(args.only)]
    make_backend = _make_backend(args.backend, args.model, args.base_url)

    print(f"backend={args.backend} model={args.model} runs={args.runs}\n")
    print(f"{'scenario':<34} {'pass-rate':>10} {'mean':>7}")
    print("-" * 54)
    with tempfile.TemporaryDirectory() as td:
        for scen in scenarios:
            summ = evaluate(scen, make_backend,
                            workdir=Path(td) / scen.slug, runs=args.runs)
            print(f"{scen.name:<34} {summ.pass_rate:>9.0%} {summ.mean_score:>7.2f}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
