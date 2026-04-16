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
        "name": "get_cell",
        "description": (
            "셀 하나의 전체 내용과 병합/중첩 메타를 반환한다. "
            "inspect_document의 preview는 max_cell_len으로 잘리지만 get_cell은 전체 텍스트를 돌려준다. "
            "병합 영역 내부의 non-anchor 좌표로 호출해도 에러 없이 anchor 셀의 내용을 반환한다 "
            "(is_anchor=false, anchor/span 필드로 구조 확인 가능). "
            "nested_table_indices: 셀 안에 중첩 테이블이 있을 경우 그 flat index 목록 (DOCX/HWPX)."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "path": {"type": "string"},
                "table_index": {"type": "integer"},
                "row": {"type": "integer"},
                "col": {"type": "integer"},
            },
            "required": ["path", "table_index", "row", "col"],
        },
    },
    {
        "name": "set_cell",
        "description": (
            "특정 표의 셀 값을 교체한다. table_index는 inspect_document의 tables 배열 인덱스. "
            "3개 포맷(DOCX/PPTX/HWPX) 모두 병합 셀을 인지한다: inspect_document의 "
            "tables[i].merges에 나온 anchor 좌표로만 수정 가능. 병합 영역 내부의 non-anchor "
            "좌표로 호출하면 MergedCellWriteError(ValueError 호환)가 발생하며, preview의 해당 "
            "슬롯은 null로 표시된다."
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
                        "true면 병합 영역 non-anchor 좌표 호출 시 앵커로 자동 리디렉트 + 경고 "
                        "(권장 X, 구조 잘못 이해한 호출을 숨김)."
                    ),
                    "default": False,
                },
            },
            "required": ["path", "table_index", "row", "col", "value"],
        },
    },
    {
        "name": "append_to_cell",
        "description": (
            "기존 셀 텍스트 뒤에 separator + value를 덧붙인다. "
            "한국 관공서 폼처럼 '성  명' 같은 라벨 셀 뒤에 값을 붙이는 용도로 유용. "
            "빈 셀이면 separator 없이 value만 기록. set_cell과 동일하게 병합 anchor만 수정 가능."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "path": {"type": "string"},
                "table_index": {"type": "integer"},
                "row": {"type": "integer"},
                "col": {"type": "integer"},
                "value": {"type": "string", "description": "추가할 값"},
                "separator": {
                    "type": "string",
                    "description": "기존 텍스트와 값 사이 구분자 (기본 '  ')",
                    "default": "  ",
                },
                "output_path": {
                    "type": "string",
                    "description": "생략 시 원본 덮어쓰기",
                },
                "allow_merge_redirect": {
                    "type": "boolean",
                    "default": False,
                },
            },
            "required": ["path", "table_index", "row", "col", "value"],
        },
    },
    {
        "name": "append_row",
        "description": (
            "표 끝에 새 행을 추가한다. DOCX / PPTX / HWPX 모두 지원 (v0.5+). "
            "마지막 행을 deepcopy 해 스타일/서식 상속. "
            "제약: 마지막 행이 위 행의 rowSpan 영역에 걸쳐있으면 거부."
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
    {
        "name": "fill_form",
        "description": (
            "라벨 이름으로 값 셀을 자동 탐지해 **일괄 채우기**. 좌표 (table_index, row, col) "
            "계산 없이 '접수번호', '성명' 같은 라벨 key-value dict 로 양식 채움. "
            "auto 모드: 라벨 셀 오른쪽 → 아래 → 같은 셀 순으로 값 셀 탐색. "
            "오른쪽/아래 셀이 사용자 요청 라벨 중 하나이면 (서로 라벨 공간 보호) skip 후 다음 시도. "
            "같은 셀로 fallback 시 append_to_cell 로 라벨 뒤에 값 덧붙임. "
            "**팁**: 한 양식의 관련 라벨을 함께 넘기면 라벨끼리 서로 보호하여 덮어쓰기 방지. "
            "반환: {filled:[...], not_found:[...], ambiguous:[...]}."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "path": {"type": "string"},
                "data": {
                    "type": "object",
                    "description": "라벨: 값 dict (예: {'접수번호': '2026-0001', '성명': '홍길동'})",
                    "additionalProperties": {"type": "string"},
                },
                "direction": {
                    "type": "string",
                    "enum": ["auto", "right", "below", "same"],
                    "default": "auto",
                    "description": "값 셀 탐색 방향. auto 권장.",
                },
                "strict": {
                    "type": "boolean",
                    "default": False,
                    "description": "True 면 라벨 매칭 실패 시 에러. False 면 not_found 에 기록.",
                },
                "output_path": {"type": "string"},
            },
            "required": ["path", "data"],
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


def get_cell(path: str, table_index: int, row: int, col: int) -> dict[str, Any]:
    doc = load(path)
    try:
        cell = doc.get_cell(table_index, row, col)
        return cell.to_dict()
    finally:
        doc.close()


def set_cell(path: str, table_index: int, row: int, col: int, value: str,
             output_path: str | None = None,
             allow_merge_redirect: bool = False) -> dict[str, Any]:
    target = Path(output_path) if output_path else Path(path)
    if output_path and Path(path) != target:
        shutil.copy2(path, target)

    doc = load(target)
    try:
        old = doc.set_cell(
            table_index, row, col, value,
            allow_merge_redirect=allow_merge_redirect,
        )
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


def append_to_cell(path: str, table_index: int, row: int, col: int, value: str,
                   separator: str = "  ",
                   output_path: str | None = None,
                   allow_merge_redirect: bool = False) -> dict[str, Any]:
    target = Path(output_path) if output_path else Path(path)
    if output_path and Path(path) != target:
        shutil.copy2(path, target)

    doc = load(target)
    try:
        old = doc.append_to_cell(
            table_index, row, col, value,
            separator=separator,
            allow_merge_redirect=allow_merge_redirect,
        )
        doc.save()
    finally:
        doc.close()

    return {
        "output_path": str(target),
        "table_index": table_index,
        "row": row,
        "col": col,
        "previous_value": old,
        "appended_value": value,
        "separator": separator,
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


def fill_form(path: str, data: dict[str, str],
              direction: str = "auto", strict: bool = False,
              output_path: str | None = None) -> dict[str, Any]:
    target = Path(output_path) if output_path else Path(path)
    if output_path and Path(path) != target:
        shutil.copy2(path, target)

    doc = load(target)
    try:
        result = doc.fill_form(data, direction=direction, strict=strict)
        doc.save()
    finally:
        doc.close()

    result["output_path"] = str(target)
    return result


# -------- 이름으로 dispatch --------

TOOL_HANDLERS = {
    "inspect_document": inspect_document,
    "render_template": render_template,
    "get_cell": get_cell,
    "set_cell": set_cell,
    "append_to_cell": append_to_cell,
    "append_row": append_row,
    "fill_form": fill_form,
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
