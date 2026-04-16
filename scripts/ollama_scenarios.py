"""Ollama 로 여러 모델 × 여러 시나리오 매트릭스 실험.

UX 개선을 위한 실측 데이터 수집 — LLM 이 fill_form / direction / dot-path 를
어떻게 사용하는지 관찰.

실행:
    ssh -L 11434:localhost:11434 -N home &
    python scripts/ollama_scenarios.py
"""
from __future__ import annotations

import json
import shutil
import sys
import tempfile
import time
from dataclasses import dataclass, field
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from ollama import Client

from document_adapter.tools import TOOL_DEFINITIONS, call_tool


MODELS = ["qwen2.5:14b", "qwen3.5:4b"]

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
5. output_path 는 생략 (원본 덮어쓰기).
"""


SCENARIOS = [
    {
        "name": "A. HWPX blank 양식 — auto 보수적 append 기대",
        "fixture": "tests/fixtures/hwpx/real/stop_payment_blank.hwpx",
        "instruction": (
            "이 양식을 다음 정보로 채워줘: "
            "접수번호=2026-0001, 접수일자=2026-04-17, "
            "성명=홍길동, 주소=서울시 강남구"
        ),
    },
    {
        "name": "B. PPTX 예시값 있는 양식 — direction='right' 사용 기대",
        "fixture": "tests/fixtures/pptx/real/ai_plan_small.pptx",
        "instruction": (
            "이 보고서 양식의 기존 예시값을 다음으로 교체해줘: "
            "보고일자=2026-04-17, 작성자=홍길동, 담당부서=개발팀. "
            "기존 값은 예시일 뿐이라 덮어써야 해."
        ),
    },
    {
        "name": "C. HWPX 복잡 양식 — ambiguous 해소 기대 (dot-path)",
        "fixture": "tests/fixtures/hwpx/real/stop_payment_blank.hwpx",
        "instruction": (
            "이 양식의 피해자 정보 섹션과 지급정지요청계좌 섹션 각각에 "
            "금융회사=국민은행, 금액=1,000,000원 을 채워줘. "
            "두 섹션에 같은 라벨이 있으니 구분이 필요해."
        ),
    },
    {
        "name": "D. DOCX 양식 — fill_form 기본",
        "fixture": "tests/fixtures/docx/real/employee_form.docx",
        "instruction": (
            "이 직원 정보 양식을 다음 정보로 채워줘: "
            "성명=홍길동, 부서=개발팀, 직급=책임, 입사일=2026-04-17, "
            "전화번호=010-1234-5678, 이메일=hong@example.com, 주소=서울시 강남구"
        ),
    },
]


@dataclass
class Run:
    model: str
    scenario: str
    turns: int = 0
    elapsed_s: float = 0.0
    prompt_tokens: int = 0
    completion_tokens: int = 0
    tool_sequence: list[str] = field(default_factory=list)
    used_fill_form: bool = False
    used_direction_right: bool = False
    used_dot_path: bool = False
    errors: list[str] = field(default_factory=list)
    final_text: str = ""


def to_openai_tools(tool_defs):
    return [
        {
            "type": "function",
            "function": {
                "name": t["name"],
                "description": t["description"],
                "parameters": t["input_schema"],
            },
        }
        for t in tool_defs
    ]


def run_one(model: str, scenario: dict, verbose: bool = True) -> Run:
    run = Run(model=model, scenario=scenario["name"])
    src = ROOT / scenario["fixture"]

    # fixture 를 임시 파일에 copy — 원본 보호
    suffix = src.suffix
    with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
        shutil.copy2(src, tmp.name)
        working_path = tmp.name

    try:
        client = Client(host="http://localhost:11434")
        tools = to_openai_tools(TOOL_DEFINITIONS)
        messages = [
            {"role": "system", "content": SYSTEM},
            {
                "role": "user",
                "content": f"문서: {working_path}\n\n요청: {scenario['instruction']}",
            },
        ]

        t_start = time.time()
        for turn in range(8):
            resp = client.chat(
                model=model,
                messages=messages,
                tools=tools,
                options={"num_predict": 1024, "temperature": 0.2},
            )
            msg = resp.message
            content = getattr(msg, "content", None) or ""
            tool_calls = getattr(msg, "tool_calls", None) or []
            run.turns = turn + 1
            run.prompt_tokens = max(run.prompt_tokens, getattr(resp, "prompt_eval_count", 0) or 0)
            run.completion_tokens += getattr(resp, "eval_count", 0) or 0

            if verbose:
                print(f"    turn {turn+1}: calls={len(tool_calls)} content_len={len(content)}")

            if not tool_calls:
                run.final_text = content
                break

            messages.append({
                "role": "assistant",
                "content": content,
                "tool_calls": [
                    {
                        "function": {
                            "name": tc.function.name,
                            "arguments": tc.function.arguments or {},
                        }
                    }
                    for tc in tool_calls
                ],
            })
            for tc in tool_calls:
                name = tc.function.name
                args = tc.function.arguments or {}
                if isinstance(args, str):
                    try:
                        args = json.loads(args)
                    except json.JSONDecodeError:
                        args = {}
                run.tool_sequence.append(name)
                if name == "fill_form":
                    run.used_fill_form = True
                    if args.get("direction") == "right":
                        run.used_direction_right = True
                    data = args.get("data", {})
                    if any("." in str(k) for k in data.keys()):
                        run.used_dot_path = True
                try:
                    result = call_tool(name, args)
                except Exception as e:
                    run.errors.append(f"{name}: {type(e).__name__}: {e}")
                    result = {"error": str(e)}
                messages.append({
                    "role": "tool",
                    "content": json.dumps(result, ensure_ascii=False),
                })
        run.elapsed_s = time.time() - t_start
    finally:
        try:
            Path(working_path).unlink()
        except OSError:
            pass
    return run


def print_run(r: Run) -> None:
    print(f"\n  [{r.model}] tool seq: {r.tool_sequence}")
    print(
        f"    turns={r.turns} elapsed={r.elapsed_s:.1f}s "
        f"prompt={r.prompt_tokens} completion={r.completion_tokens}"
    )
    flags = []
    if r.used_fill_form:
        flags.append("fill_form ✓")
    else:
        flags.append("fill_form ✗")
    if r.used_direction_right:
        flags.append("direction=right ✓")
    if r.used_dot_path:
        flags.append("dot-path ✓")
    print(f"    flags: {', '.join(flags)}")
    if r.errors:
        print(f"    errors: {r.errors[:2]}")
    if r.final_text:
        print(f"    final: {r.final_text[:200]}")


def main() -> int:
    results: list[Run] = []
    for scenario in SCENARIOS:
        print(f"\n{'=' * 72}\n{scenario['name']}\n{'=' * 72}")
        for model in MODELS:
            print(f"\n→ {model}")
            try:
                r = run_one(model, scenario, verbose=True)
            except Exception as e:
                r = Run(model=model, scenario=scenario["name"])
                r.errors.append(f"run crashed: {type(e).__name__}: {e}")
            results.append(r)
            print_run(r)

    # 매트릭스 요약
    print(f"\n{'=' * 72}\n요약 매트릭스\n{'=' * 72}")
    print(f"{'Model':<14} {'Scen':<5} {'Turns':>5} {'Time':>7} {'fill_form':>10} {'right':>6} {'dot-path':>9}")
    print("-" * 72)
    scen_ids = ["A", "B", "C"]
    for r in results:
        scen_id = next((s for s in scen_ids if r.scenario.startswith(s)), "?")
        print(
            f"{r.model:<14} {scen_id:<5} {r.turns:>5} {r.elapsed_s:>6.1f}s "
            f"{str(r.used_fill_form):>10} {str(r.used_direction_right):>6} "
            f"{str(r.used_dot_path):>9}"
        )
    return 0


if __name__ == "__main__":
    sys.exit(main())
