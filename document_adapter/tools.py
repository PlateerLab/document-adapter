"""LLM tool 정의 + 실행 함수.

동일 구현을 MCP 서버와 Claude API Tool Use 양쪽에서 재사용한다.
각 함수는 JSON-serializable dict를 반환.
"""
from __future__ import annotations

import shutil
from pathlib import Path
from typing import Any

from . import load

# -------- JSON schemas (Claude API tool use와 MCP 공용) --------

TOOL_DEFINITIONS: list[dict[str, Any]] = [
    {
        "name": "inspect_document",
        "description": (
            "문서(.docx/.pptx/.hwpx)의 구조를 분석한다. "
            "placeholders({{key}} 태그 목록)와 tables(각 표의 행/열/미리보기)를 반환한다. "
            "LLM이 어떤 필드를 채우거나 수정할지 판단할 때 먼저 호출해야 한다."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "path": {
                    "type": "string",
                    "description": "문서 절대경로",
                },
                "min_rows": {
                    "type": "integer",
                    "description": "표 필터: 최소 행 수 (기본 1)",
                    "default": 1,
                },
                "min_cols": {
                    "type": "integer",
                    "description": "표 필터: 최소 열 수 (기본 1)",
                    "default": 1,
                },
            },
            "required": ["path"],
        },
    },
    {
        "name": "render_template",
        "description": (
            "문서의 {{key}} placeholder를 context의 값으로 치환해 새 파일로 저장한다. "
            "DOCX는 docxtpl(Jinja2 loop/if 지원), PPTX/HWPX는 단순 {{key}} 치환."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "path": {"type": "string", "description": "템플릿 파일 경로"},
                "context": {
                    "type": "object",
                    "description": "{{key}}에 주입할 값 dict",
                    "additionalProperties": True,
                },
                "output_path": {
                    "type": "string",
                    "description": "결과 저장 경로 (생략 시 원본 옆에 _rendered 붙여 저장)",
                },
            },
            "required": ["path", "context"],
        },
    },
    {
        "name": "set_cell",
        "description": (
            "특정 표의 셀 값을 교체한다. table_index는 inspect_document의 tables 배열 인덱스. "
            "PPTX는 슬라이드 경계와 무관한 전역 index. "
            "HWPX 병합 셀 주의: inspect_document의 tables[i].merges에 나온 anchor 좌표로만 "
            "수정 가능. 병합 영역 내부의 non-anchor 좌표로 호출하면 ValueError가 발생하며, "
            "preview의 해당 슬롯은 null로 표시된다."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "path": {"type": "string"},
                "table_index": {"type": "integer"},
                "row": {"type": "integer"},
                "col": {"type": "integer"},
                "value": {"type": "string"},
                "output_path": {
                    "type": "string",
                    "description": "생략 시 원본 덮어쓰기",
                },
                "allow_merge_redirect": {
                    "type": "boolean",
                    "description": (
                        "HWPX 전용. true면 병합 영역 non-anchor 좌표 호출 시 "
                        "앵커로 자동 리디렉트(권장 X, 구조 잘못 이해한 호출을 숨김)."
                    ),
                    "default": False,
                },
            },
            "required": ["path", "table_index", "row", "col", "value"],
        },
    },
    {
        "name": "append_row",
        "description": (
            "표 끝에 새 행을 추가한다. **DOCX만 지원** — PPTX/HWPX는 API 미지원으로 에러 반환. "
            "그 경우 템플릿 단계에서 충분한 빈 행을 두고 set_cell로 채워야 한다."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "path": {"type": "string"},
                "table_index": {"type": "integer"},
                "values": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": "새 행의 각 셀 값. 열 수보다 적으면 나머지는 공백.",
                },
                "output_path": {"type": "string"},
            },
            "required": ["path", "table_index", "values"],
        },
    },
]


# -------- 실행 함수 --------

def _resolve_output(path: str, output_path: str | None, suffix: str = "_out") -> Path:
    if output_path:
        return Path(output_path)
    p = Path(path)
    return p.with_name(f"{p.stem}{suffix}{p.suffix}")


def inspect_document(path: str, min_rows: int = 1, min_cols: int = 1) -> dict[str, Any]:
    doc = load(path)
    try:
        schema = doc.get_schema()
        # min_rows/min_cols 필터 재적용
        filtered = [t for t in doc.get_tables(min_rows=min_rows, min_cols=min_cols)]
        result = schema.to_dict()
        result["tables"] = [t.to_dict() for t in filtered]
        return result
    finally:
        doc.close()


def render_template(path: str, context: dict[str, Any],
                    output_path: str | None = None) -> dict[str, Any]:
    out = _resolve_output(path, output_path, "_rendered")
    shutil.copy2(path, out)

    doc = load(out)
    try:
        before = doc.get_placeholders()
        doc.render_template(context)
        doc.save()
    finally:
        doc.close()

    # 검증 재로드
    doc2 = load(out)
    try:
        after = doc2.get_placeholders()
    finally:
        doc2.close()

    return {
        "output_path": str(out),
        "placeholders_before": before,
        "placeholders_after": after,
        "rendered_count": len(before) - len(after),
    }


def set_cell(path: str, table_index: int, row: int, col: int, value: str,
             output_path: str | None = None,
             allow_merge_redirect: bool = False) -> dict[str, Any]:
    target = Path(output_path) if output_path else Path(path)
    if output_path and Path(path) != target:
        shutil.copy2(path, target)

    doc = load(target)
    try:
        # allow_merge_redirect는 HWPX 어댑터만 지원하므로 키워드 인자로 전달 시도하고
        # 포맷이 지원 안 하면 무시.
        try:
            old = doc.set_cell(
                table_index, row, col, value,
                allow_merge_redirect=allow_merge_redirect,
            )
        except TypeError:
            old = doc.set_cell(table_index, row, col, value)
        doc.save()
    finally:
        doc.close()

    return {
        "output_path": str(target),
        "table_index": table_index,
        "row": row,
        "col": col,
        "previous_value": old,
        "new_value": value,
    }


def append_row(path: str, table_index: int, values: list[str],
               output_path: str | None = None) -> dict[str, Any]:
    target = Path(output_path) if output_path else Path(path)
    if output_path and Path(path) != target:
        shutil.copy2(path, target)

    doc = load(target)
    try:
        doc.append_row(table_index, values)
        doc.save()
        new_tables = doc.get_tables()
    finally:
        doc.close()

    target_schema = next((t for t in new_tables if t.index == table_index), None)
    return {
        "output_path": str(target),
        "table_index": table_index,
        "new_row_count": target_schema.rows if target_schema else None,
        "appended_values": values,
    }


# -------- 이름으로 dispatch --------

TOOL_HANDLERS = {
    "inspect_document": inspect_document,
    "render_template": render_template,
    "set_cell": set_cell,
    "append_row": append_row,
}


def call_tool(name: str, arguments: dict[str, Any]) -> dict[str, Any]:
    """이름으로 tool 실행. 예외도 dict로 직렬화."""
    handler = TOOL_HANDLERS.get(name)
    if handler is None:
        return {"error": f"unknown tool: {name}"}
    try:
        return handler(**arguments)
    except NotImplementedError as e:
        return {"error": "not_implemented", "message": str(e)}
    except (IndexError, ValueError, FileNotFoundError) as e:
        return {"error": type(e).__name__, "message": str(e)}
    except Exception as e:
        return {"error": "unexpected", "type": type(e).__name__, "message": str(e)}
