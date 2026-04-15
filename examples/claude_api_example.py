"""Claude API Tool Use로 document-adapter를 호출하는 예시.

실행:
    ANTHROPIC_API_KEY=xxx python examples/claude_api_example.py <path-to-doc>
"""
from __future__ import annotations

import json
import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import anthropic

from document_adapter.tools import TOOL_DEFINITIONS, call_tool


def run_agent(doc_path: str, user_instruction: str) -> None:
    client = anthropic.Anthropic()
    tools = [
        {
            "name": t["name"],
            "description": t["description"],
            "input_schema": t["input_schema"],
        }
        for t in TOOL_DEFINITIONS
    ]

    messages = [
        {
            "role": "user",
            "content": f"문서: {doc_path}\n\n요청: {user_instruction}",
        }
    ]

    for _ in range(10):  # 최대 10턴
        resp = client.messages.create(
            model="claude-opus-4-6",
            max_tokens=4096,
            tools=tools,
            messages=messages,
        )

        messages.append({"role": "assistant", "content": resp.content})

        if resp.stop_reason != "tool_use":
            for block in resp.content:
                if block.type == "text":
                    print("\n[Claude]", block.text)
            break

        tool_results = []
        for block in resp.content:
            if block.type == "tool_use":
                print(f"\n[tool call] {block.name}({json.dumps(block.input, ensure_ascii=False)[:120]}...)")
                result = call_tool(block.name, block.input)
                print(f"[tool result] {json.dumps(result, ensure_ascii=False)[:200]}...")
                tool_results.append({
                    "type": "tool_result",
                    "tool_use_id": block.id,
                    "content": json.dumps(result, ensure_ascii=False),
                })
        messages.append({"role": "user", "content": tool_results})


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("usage: python claude_api_example.py <doc_path> [instruction]")
        sys.exit(1)
    doc = sys.argv[1]
    instruction = sys.argv[2] if len(sys.argv) > 2 else (
        "이 문서의 구조를 inspect하고, 첫 번째 표의 첫 셀을 '[Claude 수정]'으로 바꿔줘. "
        "결과는 같은 폴더에 _edited 붙여서 저장해."
    )
    run_agent(doc, instruction)
