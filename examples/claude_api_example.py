"""Claude API Tool Use 로 document-adapter 7 개 도구 에이전트 루프 예시.

실행:
    ANTHROPIC_API_KEY=xxx python examples/claude_api_example.py <doc_path> [instruction]

기본 instruction 은 v0.7 의 `fill_form` 사용을 유도한다:
    "이 양식을 채워줘: 접수번호=2026-0001, 성명=홍길동, 담당부서=개발팀"
→ LLM 흐름 (예상):
    1. inspect_document → 표 구조 + 셀 크기 메타 확인
    2. fill_form({...}) 한 번에 다수 라벨 채움
    3. (필요 시) ambiguous 반환받으면 dot-path 로 재호출

Prompt cache (5분 TTL) 로 tool schema 를 캐시해 반복 호출 비용 절감.
"""
from __future__ import annotations

import json
import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import anthropic

from document_adapter.tools import TOOL_DEFINITIONS, call_tool


SYSTEM = """당신은 DOCX / PPTX / HWPX 양식 문서를 편집하는 에이전트입니다.

⚠ **반드시 tools API 로 호출**하세요. 응답 텍스트에 JSON 코드블록이나 함수
호출 문법을 직접 작성하지 마세요 — 그건 호출되지 않습니다.

워크플로우:
1. **먼저 inspect_document 로 구조 파악** — placeholders, 표 preview, 병합 셀,
   column_widths_cm / row_heights_cm (오버플로 방지 힌트).
2. **fill_form 1 회 호출로 여러 셀을 한 번에 채우는 것을 우선**. set_cell 반복보다
   iteration 효율이 높고 라벨 오염이 덜함.
3. direction 선택:
   - **값 셀이 비어있는 양식** (HWPX 공공 서식) → direction 생략 (auto).
   - **기존 예시값이 있는 양식** (PPTX 템플릿) → direction="right" 명시.
4. 같은 라벨이 여러 섹션에 있어 ambiguous 반환받으면 dot-path 로 재호출:
     fill_form({"피해자.금액": "1,000,000", "지급정지.금액": "2,000,000"})
5. output_path 는 생략 (원본에 덮어쓰기). 별도 저장 필요할 때만 지정.
6. 셀 크기 (width_cm, char_count) 를 보고 좁은 셀에는 짧은 값만.

**중요**:
- inspect_document 는 세션당 1 회면 충분 (구조는 편집 후 변하지 않음).
- 편집 도구 반환 문자열로 성공/실패 판단. 재확인 목적 inspect 호출 금지.
"""


def run_agent(doc_path: str, user_instruction: str) -> None:
    client = anthropic.Anthropic()
    tools = [
        {
            "name": t["name"],
            "description": t["description"],
            "input_schema": t["input_schema"],
            # Tool schema 는 반복 호출 동안 불변 → 마지막 도구에 cache 브레이크포인트.
            **({"cache_control": {"type": "ephemeral"}} if t is TOOL_DEFINITIONS[-1] else {}),
        }
        for t in TOOL_DEFINITIONS
    ]

    messages = [
        {
            "role": "user",
            "content": f"문서: {doc_path}\n\n요청: {user_instruction}",
        }
    ]

    for turn in range(10):  # 최대 10턴
        resp = client.messages.create(
            model="claude-opus-4-6",
            max_tokens=4096,
            system=SYSTEM,
            tools=tools,
            messages=messages,
        )

        messages.append({"role": "assistant", "content": resp.content})

        if resp.stop_reason != "tool_use":
            for block in resp.content:
                if block.type == "text":
                    print(f"\n[Claude / turn {turn+1}]", block.text)
            # 캐시 통계
            u = resp.usage
            cache_read = getattr(u, "cache_read_input_tokens", 0) or 0
            cache_write = getattr(u, "cache_creation_input_tokens", 0) or 0
            print(
                f"\n[usage] input={u.input_tokens} (cache_read={cache_read}, "
                f"cache_write={cache_write}) output={u.output_tokens}"
            )
            break

        tool_results = []
        for block in resp.content:
            if block.type == "tool_use":
                args_preview = json.dumps(block.input, ensure_ascii=False)[:150]
                print(f"\n[tool call / turn {turn+1}] {block.name}({args_preview}...)")
                result = call_tool(block.name, block.input)
                result_preview = json.dumps(result, ensure_ascii=False)[:250]
                print(f"[tool result] {result_preview}...")
                tool_results.append({
                    "type": "tool_result",
                    "tool_use_id": block.id,
                    "content": json.dumps(result, ensure_ascii=False),
                })
        messages.append({"role": "user", "content": tool_results})


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("usage: python claude_api_example.py <doc_path> [instruction]")
        print()
        print("예시:")
        print("  python claude_api_example.py form.hwpx")
        print("  python claude_api_example.py form.hwpx '접수번호 2026-0001, 성명 홍길동, 주소 서울시로 채워줘'")
        sys.exit(1)
    doc = sys.argv[1]
    instruction = sys.argv[2] if len(sys.argv) > 2 else (
        "이 양식 문서를 다음 정보로 채워줘:\n"
        "- 접수번호: 2026-0001\n"
        "- 접수일자: 2026-04-17\n"
        "- 성명: 홍길동\n"
        "- 담당부서: 개발팀\n"
        "양식에 없는 필드는 무시하고, 결과 파일을 _filled 붙여 저장해줘."
    )
    run_agent(doc, instruction)
