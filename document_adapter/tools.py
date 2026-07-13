"""LLM tool 정의 + 실행 함수.

동일 구현을 MCP 서버와 Claude API Tool Use 양쪽에서 재사용한다.
각 함수는 JSON-serializable dict를 반환.
"""
from __future__ import annotations

import shutil
from pathlib import Path
from typing import Any, Callable

from . import load

# -------- JSON schemas (Claude API tool use와 MCP 공용) --------

TOOL_DEFINITIONS: list[dict[str, Any]] = [
    {
        "name": "inspect_document",
        "description": (
            "문서(.docx/.pptx/.hwpx)의 구조를 분석한다. "
            "placeholders({{key}} 태그 목록)와 tables(각 표의 행/열/미리보기)를 반환한다. "
            "LLM이 어떤 필드를 채우거나 수정할지 판단할 때 먼저 호출해야 한다. "
            "\n"
            "**PPTX 특화**: 반환 JSON 에 `shape_summary` (total_shapes / empty_shapes / "
            "hint) 포함. `hint` 가 있으면 그 지시 를 따라 get_shapes + set_shape_text "
            "를 호출해야 보고서가 완성됨. 표만 채우면 대부분의 보고서가 비어있는 상태로 "
            "남음."
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
            "DOCX는 docxtpl(Jinja2 loop/if 지원), PPTX/HWPX는 단순 {{key}} 치환. "
            "결과로 rendered_keys / missing_keys(context에 없던 키)를 반환. "
            "on_missing 으로 누락 키 처리 방식을 3포맷 동일하게 제어."
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
                "on_missing": {
                    "type": "string",
                    "enum": ["blank", "leave", "error"],
                    "description": (
                        "context에 없는 키 처리: blank(기본,빈칸)/leave({{key}} 유지)/"
                        "error(예외). 3포맷 동일 동작."
                    ),
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
        "name": "get_shapes",
        "description": (
            "**PPTX 전용**. 표 외 shape (textbox / placeholder / 도형 내 텍스트) 목록을 "
            "반환한다. PPTX 는 17 슬라이드에 걸쳐 표 4 개 외에 수십 개의 textbox / "
            "placeholder 로 내용이 채워져 있으므로, 전체 편집에는 get_tables 만으로 "
            "부족하다. "
            "반환: shapes[] — 각 항목에 slide_index (1-based), shape_id (슬라이드 내 고유 "
            "숫자), name, kind (placeholder/text_box), text/text_preview, placeholder_type. "
            "DOCX/HWPX 는 빈 리스트 반환 (표와 paragraph 중심이라 shape 편집 개념 약함)."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "path": {"type": "string"},
                "slide_index": {
                    "type": "integer",
                    "description": "1-based. 생략 시 전체 슬라이드.",
                },
                "min_text_len": {
                    "type": "integer",
                    "default": 1,
                    "description": "이 값 미만 길이의 텍스트 shape 는 제외.",
                },
            },
            "required": ["path"],
        },
    },
    {
        "name": "set_shape_text",
        "description": (
            "**PPTX 전용**. shape (textbox/placeholder) 의 텍스트를 교체한다. run-level "
            "포맷 (폰트/크기/색상) 은 보존. get_shapes 로 slide_index + shape_id 파악 후 "
            "호출. 표 셀은 이 도구 아닌 set_cell 을 쓸 것."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "path": {"type": "string"},
                "slide_index": {"type": "integer", "description": "1-based"},
                "shape_id": {"type": "integer"},
                "text": {"type": "string"},
                "output_path": {"type": "string"},
            },
            "required": ["path", "slide_index", "shape_id", "text"],
        },
    },
    {
        "name": "diff_documents",
        "description": (
            "두 문서(예: 원본 vs 편집본)를 셀 단위로 비교해 **무엇이 어디서 어떻게 "
            "바뀌었는지** 반환. 편집/채우기 후 검증용 — 변경 셀의 before/after 와 "
            "overflow_risk(값이 칸을 넘쳐 깨질 위험)를 함께 준다. 4 포맷 공통."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "path_a": {"type": "string", "description": "이전/원본 경로"},
                "path_b": {"type": "string", "description": "이후/편집본 경로"},
            },
            "required": ["path_a", "path_b"],
        },
    },
    {
        "name": "get_text_map",
        "description": (
            "**DOCX/HWPX**. 문서의 문단 지도(본문+표 셀+머리말/꼬리말)를 문서 "
            "순서로 반환 — 표 밖 텍스트 구조를 파악하는 '눈'. 각 항목: "
            "{para_index, scope(body/table/header/footer/shape), location, "
            "text[, truncated, char_count, is_heading]}. "
            "placeholder 없는 일반 텍스트를 replace_text/insert_text 로 "
            "편집하기 전에 이 도구(또는 find_text)로 위치를 파악할 것. "
            "긴 문서는 offset/limit 로 페이징."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "path": {"type": "string"},
                "max_para_len": {
                    "type": "integer",
                    "default": 80,
                    "description": "문단 텍스트 표시 길이 상한 (원길이는 char_count)",
                },
                "offset": {"type": "integer", "default": 0},
                "limit": {"type": "integer", "default": 200},
            },
            "required": ["path"],
        },
    },
    {
        "name": "find_text",
        "description": (
            "**DOCX/HWPX**. 문서 전역(본문+표+머리말)에서 텍스트를 검색. "
            "단어가 XML run 으로 쪼개져 있어도 찾는다. 매치마다 "
            "match_index(치환/삽입의 occurrences 로 지정 가능), scope, "
            "location, «매치» 표시된 context, nearest_heading(가까운 위쪽 "
            "제목), context_before(직전 문단)를 반환 — '~부분의 ~글' 같은 "
            "위치 참조를 해소하는 근거로 사용. whole_word=true 면 "
            "'홍길동님' 속 '홍길동' 같은 부분 일치 배제."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "path": {"type": "string"},
                "query": {"type": "string", "description": "찾을 리터럴 텍스트"},
                "whole_word": {"type": "boolean", "default": False},
                "scope": {
                    "type": "string",
                    "enum": ["all", "body", "table", "header", "footer", "shape"],
                    "default": "all",
                },
            },
            "required": ["path", "query"],
        },
    },
    {
        "name": "replace_text",
        "description": (
            "**DOCX/HWPX**. 문서 전역에서 old → new 치환. run-level 서식"
            "(폰트/크기/색/굵기) 보존 — 매치가 걸친 run 만 수정. 기본은 "
            "모든 등장 치환; 특정 등장만 바꾸려면 find_text 의 match_index "
            "를 occurrences 배열로 지정. 표 셀 안 텍스트도 함께 치환된다 "
            "(셀 전체 교체는 set_cell 이 더 적합). "
            "반환: {count, changes:[{match_index, location, paragraph_before/"
            "after}...], invalid_occurrences?}."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "path": {"type": "string"},
                "old": {"type": "string"},
                "new": {"type": "string"},
                "occurrences": {
                    "type": "array",
                    "items": {"type": "integer"},
                    "description": "치환할 match_index 목록 (생략 시 전부)",
                },
                "whole_word": {"type": "boolean", "default": False},
                "scope": {
                    "type": "string",
                    "enum": ["all", "body", "table", "header", "footer", "shape"],
                    "default": "all",
                },
                "output_path": {
                    "type": "string",
                    "description": "생략 시 원본 덮어쓰기",
                },
            },
            "required": ["path", "old", "new"],
        },
    },
    {
        "name": "insert_text",
        "description": (
            "**DOCX/HWPX**. 앵커 텍스트의 앞/뒤에 텍스트 삽입 (앵커 run 의 "
            "서식 상속). '성명 옆에 홍길동이라고 써줘' 류 요청 처리용 — "
            "anchor='성명', text='홍길동', position='after', separator=' '. "
            "앵커가 여러 곳이면 전부 삽입되므로 find_text 로 match_index 를 "
            "확인해 occurrences 로 좁힐 것."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "path": {"type": "string"},
                "anchor": {"type": "string", "description": "기준 텍스트"},
                "text": {"type": "string", "description": "삽입할 텍스트"},
                "position": {
                    "type": "string",
                    "enum": ["after", "before"],
                    "default": "after",
                },
                "occurrences": {
                    "type": "array",
                    "items": {"type": "integer"},
                },
                "whole_word": {"type": "boolean", "default": False},
                "scope": {
                    "type": "string",
                    "enum": ["all", "body", "table", "header", "footer", "shape"],
                    "default": "all",
                },
                "separator": {
                    "type": "string",
                    "default": "",
                    "description": "앵커와 삽입문 사이 구분자 (예: ' ')",
                },
                "output_path": {"type": "string"},
            },
            "required": ["path", "anchor", "text"],
        },
    },
    {
        "name": "get_form_controls",
        "description": (
            "**HWPX 전용**. 표가 아닌 폼 컨트롤(체크박스/라디오/콤보/에디트) 목록을 "
            "name·kind·caption·value 로 반환. set_form_control 로 채울 대상 파악용. "
            "표 셀이 아닌 인터랙티브 입력 필드를 다룰 때 사용."
        ),
        "input_schema": {
            "type": "object",
            "properties": {"path": {"type": "string"}},
            "required": ["path"],
        },
    },
    {
        "name": "set_form_control",
        "description": (
            "**HWPX 전용**. 이름(name)으로 폼 컨트롤 값을 설정. 체크박스/라디오는 "
            "value 가 truthy('Y'/'체크'/true) 면 체크, 에디트/콤보는 문자열을 기록. "
            "get_form_controls 로 name 을 먼저 확인."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "path": {"type": "string"},
                "name": {"type": "string"},
                "value": {"type": ["string", "boolean"]},
                "output_path": {"type": "string"},
            },
            "required": ["path", "name", "value"],
        },
    },
    {
        "name": "fill_form",
        "description": (
            "라벨 이름으로 값 셀을 자동 탐지해 **일괄 채우기**. 좌표 (table_index, row, col) "
            "계산 없이 '접수번호', '성명' 같은 라벨 key-value dict 로 양식 채움. "
            "\n"
            "**direction 선택 기준 (중요)**: "
            "(a) `auto` (기본, 보수적): 기존 값이 있는 셀은 다른 라벨로 간주하고 skip → 같은 "
            "셀에 append_to_cell. **값 셀이 비어있는 양식**에 적합 (HWPX 공공 서식 등). "
            "(b) `right` / `below` (명시): 라벨 오른쪽/아래 셀을 **덮어쓰기**. **기존에 "
            "예시값이 채워져 있는 양식**(PPTX 템플릿 등) 에는 반드시 direction='right' 명시. "
            "\n"
            "**Dot-path 섹션 지정**: 같은 라벨이 여러 섹션에 있어 ambiguous 로 반환되면 "
            "`{'피해자.금액': '...', '지급정지요청계좌.금액': '...'}` 처럼 섹션힌트.라벨 "
            "형태로 재호출. ambiguous 반환의 hint 필드에 예시 제공됨. "
            "\n"
            "**output_path**: 생략 시 **원본 파일에 덮어쓰기** (대부분의 경우 생략 권장). "
            "다른 위치에 저장이 필요할 때만 지정. "
            "\n"
            "**팁**: 한 양식의 관련 라벨을 **한 번에 dict 로** 넘기면 라벨끼리 서로 보호되어 "
            "인접 라벨 오염이 방지됩니다. "
            "\n"
            "반환: `{filled: [...], not_found: [...], ambiguous: [...]}`. "
            "ambiguous candidates 각각에 `context` 필드 포함 (어느 섹션인지 확인)."
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
    {
        "name": "create_document",
        "description": (
            "**처음부터 새 문서를 생성**한다 (.docx=markdown, .xlsx=sheets). "
            "역할 경계: `{{placeholder}}` 템플릿이 이미 있으면 render_template, "
            "기존 파일 수정은 set_cell/insert_row 등 편집 도구, 파일이 아예 없을 때만 이 도구. "
            "\n"
            "**markdown 서브셋 (.docx)** — 아래만 지원, 그 외는 일반 텍스트로 처리됨: "
            "`#`~`######` 헤딩(단일 `#` 제목으로 시작) · 문단 · `- ` 불릿 · `1. ` 번호 · "
            "`**굵게**` `*기울임*` `` `코드` `` · 파이프 표(`|---|` 구분행 필수) · "
            "`> ` 인용 · `---` 수평선 · ``` 코드펜스. "
            "금지: HTML / 이미지 / 각주 / 중첩 리스트 / 셀 병합. "
            "\n"
            "**sheets 스키마 (.xlsx)**: "
            '[{"name": "매출 요약", "headers": ["분기", "매출(억원)"], '
            '"rows": [["1분기", 120], ["합계", "=SUM(B2:B2)"]], '
            '"number_formats": {"B": "#,##0"}}] — 숫자는 숫자 타입으로, '
            "`=` 시작 문자열은 살아있는 수식이 됨. 스타일(헤더/테두리/틀고정)은 자동. "
            "\n"
            "**작성 규칙**: placeholder(TBD/내용추가) 금지 — 실제 완성된 내용을 쓸 것. "
            "생성 후 미세 수정은 inspect_document 좌표로 기존 편집 도구를 사용 "
            "(반환값의 table_shapes 가 그 좌표계). 렌더 실패 메시지를 받으면 규칙에 "
            "맞춰 **1회만** 재작성."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "path": {
                    "type": "string",
                    "description": "출력 절대경로 (.docx 또는 .xlsx — 확장자가 렌더러 선택)",
                },
                "markdown": {
                    "type": "string",
                    "description": ".docx 용 본문 (제약된 markdown 서브셋)",
                },
                "sheets": {
                    "type": "array",
                    "description": ".xlsx 용 sheet spec 리스트",
                    "items": {"type": "object"},
                },
                "lang": {
                    "type": "string",
                    "default": "ko",
                    "description": "기본 폰트 스택 (ko → 맑은 고딕)",
                },
                "overwrite": {
                    "type": "boolean",
                    "default": False,
                    "description": "기존 파일 덮어쓰기 허용",
                },
            },
            "required": ["path"],
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

        # PPTX: shape 요약을 추가해 LLM 에게 set_shape_text 필요성을 신호.
        # (표 4 개만 보고 set_cell 만 호출하는 실패 패턴 관찰됨 — 실제론 수십 개
        # textbox 가 비어있어 보고서 완성이 안 됨.)
        if schema.format == "pptx":
            try:
                all_shapes = doc.get_shapes(min_text_len=0)
                total = len(all_shapes)
                empty = sum(1 for s in all_shapes if not s.has_text)
                summary: dict[str, Any] = {
                    "total_shapes": total,
                    "empty_shapes": empty,
                    "filled_shapes": total - empty,
                }
                if total > 0 and empty / total > 0.5:
                    summary["hint"] = (
                        f"⚠ 이 PPTX 는 {total} 개 shape (textbox / placeholder / 도형 "
                        f"텍스트) 중 {empty} 개가 비어있는 '빈 양식' 상태입니다. "
                        f"표({len(filtered)}개) 만 set_cell 로 채우면 보고서의 대부분이 "
                        f"여전히 비어있습니다. get_shapes 로 전체 shape 목록을 가져와 "
                        f"set_shape_text 로 시나리오 내용에 맞게 각 shape 를 채워야 "
                        f"완성본이 됩니다."
                    )
                elif total > 0 and empty > 0:
                    summary["hint"] = (
                        f"{total} 개 shape 중 {empty} 개가 비어있음. 필요 시 "
                        f"get_shapes / set_shape_text 사용."
                    )
                result["shape_summary"] = summary
            except NotImplementedError:
                pass  # DOCX / HWPX 는 shape 개념 약함 — skip
        result["tables"] = [t.to_dict() for t in filtered]

        # 중복 라벨 힌트: 여러 곳에 같은 라벨이 있으면 fill_form 에서 ambiguous
        # 가 되므로, 미리 dot-path 를 쓰라고 신호한다 (재시도 라운드 절감).
        dups = [f for f in doc.get_form_fields() if f["ambiguous"]]
        if dups:
            result["duplicate_labels"] = [
                {"label": f["label"],
                 "count": len(f["occurrences"]),
                 "locations": f["occurrences"]}
                for f in dups
            ]
            result["duplicate_labels_hint"] = (
                "같은 라벨이 여러 곳에 있습니다. fill_form 에서 이 라벨들은 "
                "dot-path(예: '피해자.금액')로 섹션을 구분해 채우세요."
            )
        return result
    finally:
        doc.close()


def render_template(path: str, context: dict[str, Any],
                    output_path: str | None = None,
                    on_missing: str = "blank") -> dict[str, Any]:
    out = _resolve_output(path, output_path, "_rendered")
    shutil.copy2(path, out)

    doc = load(out)
    try:
        before = doc.get_placeholders()
        report = doc.render_template(context, on_missing=on_missing)
        doc.save()
    finally:
        doc.close()

    return {
        "output_path": str(out),
        "placeholders_before": before,
        "rendered_keys": report["used"],
        "missing_keys": report["missing"],
        "rendered_count": len(report["used"]),
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


def get_shapes(path: str, slide_index: int | None = None,
               min_text_len: int = 1) -> dict[str, Any]:
    doc = load(path)
    try:
        shapes = doc.get_shapes(slide_index=slide_index, min_text_len=min_text_len)
    finally:
        doc.close()
    return {
        "format": doc.format,
        "source": str(path),
        "shape_count": len(shapes),
        "shapes": [s.to_dict() for s in shapes],
    }


def set_shape_text(path: str, slide_index: int, shape_id: int, text: str,
                   output_path: str | None = None) -> dict[str, Any]:
    target = Path(output_path) if output_path else Path(path)
    if output_path and Path(path) != target:
        shutil.copy2(path, target)

    doc = load(target)
    try:
        old = doc.set_shape_text(slide_index, shape_id, text)
        doc.save()
    finally:
        doc.close()

    return {
        "output_path": str(target),
        "slide_index": slide_index,
        "shape_id": shape_id,
        "previous_text": old,
        "new_text": text,
    }


def get_text_map(path: str, max_para_len: int = 80,
                 offset: int = 0, limit: int = 200) -> dict[str, Any]:
    doc = load(path)
    try:
        result = doc.get_text_map(
            max_para_len=max_para_len, offset=offset, limit=limit,
        )
    finally:
        doc.close()
    result["source"] = str(path)
    return result


def find_text(path: str, query: str, whole_word: bool = False,
              scope: str = "all") -> dict[str, Any]:
    doc = load(path)
    try:
        matches = doc.find_text(query, whole_word=whole_word, scope=scope)
    finally:
        doc.close()
    return {
        "source": str(path),
        "query": query,
        "match_count": len(matches),
        "matches": [m.to_dict() for m in matches],
    }


def replace_text(path: str, old: str, new: str,
                 occurrences: list[int] | None = None,
                 whole_word: bool = False, scope: str = "all",
                 output_path: str | None = None) -> dict[str, Any]:
    target = Path(output_path) if output_path else Path(path)
    if output_path and Path(path) != target:
        shutil.copy2(path, target)

    doc = load(target)
    try:
        result = doc.replace_text(
            old, new,
            occurrences=occurrences, whole_word=whole_word, scope=scope,
        )
        doc.save()
    finally:
        doc.close()

    result["output_path"] = str(target)
    return result


def insert_text(path: str, anchor: str, text: str,
                position: str = "after",
                occurrences: list[int] | None = None,
                whole_word: bool = False, scope: str = "all",
                separator: str = "",
                output_path: str | None = None) -> dict[str, Any]:
    target = Path(output_path) if output_path else Path(path)
    if output_path and Path(path) != target:
        shutil.copy2(path, target)

    doc = load(target)
    try:
        result = doc.insert_text(
            anchor, text,
            position=position, occurrences=occurrences,
            whole_word=whole_word, scope=scope, separator=separator,
        )
        doc.save()
    finally:
        doc.close()

    result["output_path"] = str(target)
    return result


def get_form_controls(path: str) -> dict[str, Any]:
    doc = load(path)
    try:
        ctrls = doc.get_form_controls()
    finally:
        doc.close()
    return {
        "format": doc.format,
        "source": str(path),
        "control_count": len(ctrls),
        "controls": ctrls,
    }


def set_form_control(path: str, name: str, value: Any,
                     output_path: str | None = None) -> dict[str, Any]:
    target = Path(output_path) if output_path else Path(path)
    if output_path and Path(path) != target:
        shutil.copy2(path, target)

    doc = load(target)
    try:
        old = doc.set_form_control(name, value)
        doc.save()
    finally:
        doc.close()

    return {
        "output_path": str(target),
        "name": name,
        "previous_value": old,
        "new_value": value,
    }


def diff_documents(path_a: str, path_b: str) -> dict[str, Any]:
    from document_adapter.diff import diff_documents as _diff
    return _diff(path_a, path_b)


def create_document(path: str, markdown: str | None = None,
                    sheets: list[dict[str, Any]] | None = None,
                    lang: str = "ko", overwrite: bool = False) -> dict[str, Any]:
    from .generate import create_document as _create

    out = _create(path, markdown=markdown, sheets=sheets,
                  lang=lang, overwrite=overwrite)

    # 생성 직후 좌표 피드백 — LLM 이 후속 편집(set_cell/insert_row)을
    # inspect_document 재호출 없이 바로 이어갈 수 있게 표 요약을 돌려준다.
    doc = load(out)
    try:
        tables = doc.get_tables()
        summary = {
            "tables": len(tables),
            "table_shapes": [f"{t.rows}x{t.cols}" for t in tables],
        }
    finally:
        doc.close()

    return {
        "output_path": str(out),
        "format": out.suffix.lstrip("."),
        **summary,
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

TOOL_HANDLERS: dict[str, Callable[..., dict[str, Any]]] = {
    "inspect_document": inspect_document,
    "render_template": render_template,
    "get_cell": get_cell,
    "set_cell": set_cell,
    "append_to_cell": append_to_cell,
    "append_row": append_row,
    "fill_form": fill_form,
    "get_shapes": get_shapes,
    "set_shape_text": set_shape_text,
    "get_text_map": get_text_map,
    "find_text": find_text,
    "replace_text": replace_text,
    "insert_text": insert_text,
    "get_form_controls": get_form_controls,
    "set_form_control": set_form_control,
    "diff_documents": diff_documents,
    "create_document": create_document,
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
