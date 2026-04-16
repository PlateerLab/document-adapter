"""Ollama 네이티브 API 로 document-adapter 7 개 도구 에이전트 루프.

Ollama OpenAI-compat layer (/v1/chat/completions) 는 qwen2.5 같은 모델의 tool
call 태그를 tool_calls 로 파싱 못하는 경우가 있어 Ollama 공식 Python SDK 사용.
내부적으로 /api/chat 엔드포인트 호출.

실행 전제:
    SSH tunnel (또는 Ollama 네트워크 노출):
      ssh -L 11434:localhost:11434 -N home &

    pip install ollama

실행:
    python examples/ollama_example.py <doc_path> [model]
    # 기본 model: qwen2.5:14b (tool calling 지원)
"""
from __future__ import annotations

import json
import os
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from ollama import Client

from document_adapter.tools import TOOL_DEFINITIONS, call_tool


SYSTEM = """당신은 DOCX / PPTX / HWPX 양식 문서를 편집하는 에이전트입니다.

워크플로우:
1. 먼저 inspect_document 로 구조 파악 — placeholders, 표 preview, 병합 셀,
   column_widths_cm / row_heights_cm (오버플로 방지 힌트).
2. 여러 셀을 라벨로 채우는 경우 fill_form 1 회 호출을 우선. set_cell 반복보다
   효율적이고 라벨 오염이 덜함. direction 기본 auto 는 보수적이라 예시값
   덮어쓰기가 목적이면 direction="right" 명시.
3. 같은 라벨이 여러 섹션에 있어 ambiguous 반환받으면 "피해자.금액" 같은
   dot-path 로 재호출.
4. 셀 크기 (width_cm, char_count) 를 보고 좁은 셀에는 짧은 값만 넣기.

중요: inspect_document 는 세션당 1 회면 충분. 편집 도구 반환 문자열로 성공/실패
판단하고 재확인 목적 inspect 호출 금지.
"""


def to_openai_tools(tool_defs):
    """Anthropic 포맷 → OpenAI function calling 포맷."""
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


def run_agent(doc_path: str, user_instruction: str, model: str) -> None:
    host = os.environ.get("OLLAMA_HOST", "http://localhost:11434")
    client = Client(host=host)
    tools = to_openai_tools(TOOL_DEFINITIONS)

    messages = [
        {"role": "system", "content": SYSTEM},
        {"role": "user", "content": f"문서: {doc_path}\n\n요청: {user_instruction}"},
    ]

    for turn in range(10):
        print(f"\n{'=' * 70}\n[turn {turn + 1}] 모델 호출 중 ({model})...\n{'=' * 70}")
        t0 = time.time()
        resp = client.chat(
            model=model,
            messages=messages,
            tools=tools,
            options={"num_predict": 2048, "temperature": 0.2},
        )
        elapsed = time.time() - t0
        msg = resp.message
        print(f"[turn {turn + 1}] 응답 ({elapsed:.1f}s, done_reason={resp.done_reason})")

        content = getattr(msg, "content", None) or ""
        tool_calls = getattr(msg, "tool_calls", None) or []

        if content:
            print(f"\n[assistant]\n{content[:400]}")

        if not tool_calls:
            pt = getattr(resp, "prompt_eval_count", None)
            ct = getattr(resp, "eval_count", None)
            print(f"\n[usage] prompt={pt} completion={ct}")
            break

        messages.append({
            "role": "assistant",
            "content": content,
            "tool_calls": [
                {
                    "function": {
                        "name": tc.function.name,
                        "arguments": tc.function.arguments or {},  # dict 그대로
                    },
                }
                for tc in tool_calls
            ],
        })
        for i, tc in enumerate(tool_calls):
            args = tc.function.arguments or {}
            if isinstance(args, str):
                try:
                    args = json.loads(args)
                except json.JSONDecodeError:
                    args = {"_raw": args}
            print(f"\n[tool call] {tc.function.name}")
            print(f"  args={json.dumps(args, ensure_ascii=False)}")
            result = call_tool(tc.function.name, args)
            preview = json.dumps(result, ensure_ascii=False)
            print(f"[tool result] {preview[:600]}")
            if len(preview) > 600:
                print(f"  ... (+{len(preview)-600} chars)")
            messages.append({
                "role": "tool",
                "content": json.dumps(result, ensure_ascii=False),
            })


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("usage: python ollama_example.py <doc_path> [model] [instruction]")
        print()
        print("예시:")
        print("  python ollama_example.py tests/fixtures/hwpx/real/stop_payment_blank.hwpx")
        print("  python ollama_example.py form.hwpx qwen2.5:14b '접수번호 2026-0001 로 채워줘'")
        sys.exit(1)

    doc = sys.argv[1]
    model = sys.argv[2] if len(sys.argv) > 2 else "qwen2.5:14b"
    instruction = sys.argv[3] if len(sys.argv) > 3 else (
        "이 양식 문서를 다음 정보로 채워줘:\n"
        "- 접수번호: 2026-0001\n"
        "- 접수일자: 2026-04-17\n"
        "- 성명: 홍길동\n"
        "- 주소: 서울시 강남구 역삼동 123-4\n"
        "- 전화번호: 02-1234-5678\n"
        "양식에 없는 필드는 무시하고, 결과를 같은 파일에 저장해줘."
    )
    run_agent(doc, instruction, model)
