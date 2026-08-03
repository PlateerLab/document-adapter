# document-adapter

[![PyPI version](https://img.shields.io/pypi/v/document-adapter.svg)](https://pypi.org/project/document-adapter/)
[![Python versions](https://img.shields.io/pypi/pyversions/document-adapter.svg)](https://pypi.org/project/document-adapter/)
[![License](https://img.shields.io/pypi/l/document-adapter.svg)](https://github.com/PlateerLab/document-adapter/blob/main/LICENSE)
[![MCP](https://img.shields.io/badge/MCP-compatible-blue.svg)](https://modelcontextprotocol.io)

**LLM이 DOCX / PPTX / HWPX 문서를 직접 편집할 수 있게 해주는 통합 어댑터 + MCP 서버.**

세 가지 오피스 포맷을 하나의 파이썬 인터페이스로 추상화하고, Claude Desktop / Claude Code / Anthropic API Tool Use에서 바로 호출할 수 있는 MCP 도구로 노출합니다. 양식 문서의 빈 셀을 자동으로 채우거나, 템플릿의 `{{key}}`를 치환하거나, 기존 표의 내용을 수정하는 작업을 LLM 에이전트가 수행할 수 있습니다.

- 📦 PyPI: https://pypi.org/project/document-adapter/
- 🔗 Repo: https://github.com/PlateerLab/document-adapter

## 지원 포맷

| 포맷 | 백엔드 | 템플릿 렌더 | 표 읽기 | 병합 셀 인지 | 중첩 테이블 | 셀 수정 | 행 추가 |
|---|---|---|---|---|---|---|---|
| `.docx` | `docxtpl` + `python-docx` | Jinja2 (`{%tr%}` loop 포함) | ✅ | ✅ | ✅ | ✅ | ✅ |
| `.pptx` | `python-pptx` + 자체 lxml 확장 | `{{key}}` 치환 | ✅ (슬라이드 위치 포함) | ✅ | — (포맷 미지원) | ✅ | ✅ (v0.5+) |

- **PPTX 차트 (v0.17+)**: `get_charts` / `set_chart_data`(수치 편집, 서식 보존) / `add_chart`(column/bar/line/pie/doughnut/area/radar 계열). scatter·bubble·날짜축·콤보 차트는 읽기 전용(`editable=false`).
- **PPTX 슬라이드 (v0.17+)**: `get_slides`(페이지 개요) / `duplicate_slide`(양식 페이지 복제 — 차트 포함 시 chart part + 내장 워크북까지 독립 복제, 노트는 미복제).
| `.hwpx` | 자체 `hwpx_core` (lxml + zipfile) | `{{key}}` 치환 | ✅ | ✅ | ✅ | ✅ | ✅ |
| `.xlsx` | `openpyxl` | `{{key}}` 치환 | ✅ (시트 = 표) | ✅ | — | ✅ | ✅ |

- HWPX는 한컴오피스 설치가 **불필요**합니다 (macOS/Linux 서버에서 그대로 동작).
- 구버전 `.hwp`(바이너리 포맷)는 지원하지 않습니다 — `.hwpx`로 변환 후 사용하세요.
- **XLSX (v0.10+)**: 각 워크시트를 하나의 표로 매핑(`table_index` = 시트 인덱스, `location` = 시트명). 병합 셀·`fill_form`·`render_template` 동일 인터페이스로 동작.
  - 셀 타입: 날짜는 날짜만 표시, `set_cell` 은 깔끔한 숫자(금액)를 숫자형으로 기록(전화·우편번호는 문자 유지).
  - **수식 셀**: 캐시된 계산값이 있으면(Excel 저장본) 그 값을, 없으면(openpyxl 생성본 등) 수식 문자열(`=A1+B1`)을 표시. 수식 자체는 편집/저장 시 보존됩니다.
- 병합 셀: 3개 포맷 모두 preview에 `null` 슬롯 + `merges` 메타로 구조 노출. non-anchor 좌표에 쓰기는 `MergedCellWriteError`로 거부.
- **셀 크기 메타 (v0.6+)**: `get_tables`는 `column_widths_cm` / `row_heights_cm`, `get_cell`은 `width_cm` / `height_cm` / `char_count`를 반환합니다. LLM이 좁은 셀(예: 1.7×0.7cm 배지)에 긴 텍스트를 넣어 오버플로 되는 것을 사전에 판단할 수 있습니다.

## 라이선스 (상용 사용 가능)

- 본 프로젝트: **Apache License 2.0** — 재배포·파생물 배포 시 저작권·`NOTICE` 출처 표시 유지 필요(라이선스 §4). 상용 사용 가능, 특허 사용권 포함.
- 런타임 의존성(`python-docx`, `docxtpl`, `python-pptx`, `lxml`, `mcp`): 전부 **허용형 OSS** (MIT/BSD/Apache-2.0/LGPL-2.1). 상용·내부 서비스에 그대로 포함 가능.
- v0.3 이하에서 사용했던 `python-hwpx` (Non-Commercial License) 는 v0.4.0부터 **dev 환경(테스트 fixture 생성) 전용**으로 이동. HWPX 편집은 자체 `hwpx_core` 모듈이 수행합니다.

## 설치

```bash
pip install document-adapter
```

Claude API 예시 스크립트까지 포함:

```bash
pip install "document-adapter[claude]"
```

개발 환경에서 소스로 설치:

```bash
git clone https://github.com/PlateerLab/document-adapter.git
cd document-adapter
pip install -e ".[dev]"
```

Python 3.10+ 필요.

## 빠른 시작 — 파이썬 API

```python
from document_adapter import load

doc = load("report_template.docx")

# 1. 구조 파악
schema = doc.get_schema()
print(schema.placeholders)   # ['author', 'date', 'title']
print(schema.tables)         # [TableSchema(index=0, rows=7, cols=2, ...), ...]

# 2. 템플릿 렌더
doc.render_template({
    "title": "Q1 운영 리포트",
    "author": "손성준",
    "date": "2026-04-15",
})
doc.save("report_filled.docx")

# 3. 기존 양식 파일의 표 셀 수정
doc = load("checklist.docx")

# 빈 셀 값 교체
old = doc.set_cell(table_index=1, row=1, col=1, value="○○전자")

# 라벨이 있는 셀 ("성 명")에 값 추가 → "성 명  홍길동"
doc.append_to_cell(table_index=2, row=0, col=0, value="홍길동")

# 셀 전체 텍스트 + 병합 메타 조회 (preview의 40자 잘림 없이)
cell = doc.get_cell(table_index=1, row=3, col=2)
print(cell.text, cell.is_anchor, cell.span, cell.nested_table_indices)

# DOCX/PPTX/HWPX 전부 행 추가 지원 (v0.5+)
doc.append_row(1, ["새 항목", "값"])

doc.save("checklist_filled.docx")
doc.close()
```

### 라벨 기반 일괄 채우기 (v0.7+)

LLM 이 좌표 `(table_index, row, col)` 를 직접 계산하지 않고 "접수번호", "성명" 같은 사람이 읽는 라벨 key-value 로 양식을 채울 수 있습니다.

```python
doc = load("form.hwpx")
result = doc.fill_form({
    "접수번호": "2026-0001",
    "성 명":   "홍길동",
    "주 소":   "서울시 강남구",
    "금융회사": "국민은행",
})
# → {"filled": [...], "not_found": [...], "ambiguous": [...]}
doc.save()
doc.close()
```

- `auto` (기본): 라벨 셀 오른쪽 → 아래 → 같은 셀 순으로 값 셀 탐색. 보수적이라 기존 값 있는 셀은 다른 라벨로 간주하고 skip.
- `direction="right"` 명시: 라벨 오른쪽 셀을 **덮어쓰기** (예시값 있는 PPTX 템플릿 등).
- **Dot-path 섹션 지정**: 동일 라벨이 여러 섹션에 있으면 `"피해자.금액"`, `"지급정지요청계좌.금액"` 처럼 섹션 힌트 부여. `ambiguous` 반환 시 `hint` 필드에 예시 제공.
- **팁**: 한 양식의 관련 라벨을 한 번에 dict 로 넘기면 라벨끼리 서로 보호되어 오염을 방지합니다.

### 셀 크기 메타 (v0.6+)

`get_tables()`가 `column_widths_cm` / `row_heights_cm` 를, `get_cell()`이 `width_cm` / `height_cm` / `char_count` 를 반환해 LLM이 **좁은 셀에 긴 텍스트를 넣어 오버플로 되는 것을 사전에 판단**할 수 있습니다.

```python
cell = doc.get_cell(table_index=0, row=0, col=0)
print(cell.width_cm, cell.char_count)  # 1.7cm, 4자 — 작은 배지
```

DOCX/PPTX는 EMU → cm, HWPX는 HU → cm 자동 환산 (1자리 반올림).

확장자로 자동 분기되므로 `.pptx` / `.hwpx`도 동일한 API를 사용합니다.

### 병합 셀 인지 동작 (v0.2+)

```python
schema = doc.get_schema()
t = schema.tables[0]

# preview는 logical grid. 병합된 non-anchor 슬롯은 None.
# [['HEADER', None, None], ['A1', 'A2', 'A3']]
print(t.preview)

# merges는 span>1x1인 anchor 목록
# [MergeInfo(anchor=(0,0), span=(1,3))]
print(t.merges)

# non-anchor에 쓰기 시도하면 MergedCellWriteError (ValueError 서브클래스)
try:
    doc.set_cell(0, 0, 2, "X")
except ValueError as e:
    print(e)  # "cell (0,2) is part of a merged region anchored at (0,0)..."

# 의도적으로 앵커로 리디렉트하고 싶다면
doc.set_cell(0, 0, 2, "X", allow_merge_redirect=True)  # 경고 + 실제로 (0,0) 수정
```

## MCP 서버로 사용 — Claude Desktop / Claude Code

### 실행

```bash
python -m document_adapter.mcp_server
# 또는 설치 후
document-adapter-mcp
```

### Claude Desktop 설정

`~/Library/Application Support/Claude/claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "document-adapter": {
      "command": "/absolute/path/to/venv/bin/python",
      "args": ["-m", "document_adapter.mcp_server"]
    }
  }
}
```

재시작하면 Claude Desktop에서 아래 7개 도구를 사용할 수 있습니다.

### Claude Code 설정

```bash
claude mcp add document-adapter \
  /absolute/path/to/venv/bin/python -m document_adapter.mcp_server
```

## Anthropic API Tool Use로 사용

`document_adapter.tools`가 Claude API의 tool schema 형식과 그대로 호환됩니다.

```python
import anthropic
from document_adapter.tools import TOOL_DEFINITIONS, call_tool

client = anthropic.Anthropic()

resp = client.messages.create(
    model="claude-opus-4-6",
    max_tokens=4096,
    tools=[{
        "name": t["name"],
        "description": t["description"],
        "input_schema": t["input_schema"],
    } for t in TOOL_DEFINITIONS],
    messages=[{
        "role": "user",
        "content": "report_template.docx의 표 구조를 확인하고 빈 셀을 적절히 채워줘",
    }],
)

# tool_use 블록을 받으면 call_tool(name, args)로 실행 후 결과 반환
```

전체 agent loop 예시는 [`examples/claude_api_example.py`](examples/claude_api_example.py) 참고.

## 노출되는 도구

총 **17 개** 도구 (DOCX/PPTX/HWPX/XLSX 공통, 확장자로 자동 디스패치):

| 도구 | 설명 |
|---|---|
| `inspect_document` | 문서 구조(placeholders, tables + `column_widths_cm`/`row_heights_cm`)를 JSON으로 반환. 같은 라벨이 여러 곳이면 `duplicate_labels` 힌트도 포함. **항상 첫 호출로 사용** |
| `render_template` | `{{key}}` 치환 + 조건/표현식(`{% if %}`, `{{ a*b }}`). `on_missing`(blank/leave/error)로 누락 키 처리, `rendered_keys`/`missing_keys` 반환 |
| `get_cell` | 셀 전체 텍스트 + 병합/중첩 메타 + `width_cm`/`height_cm`/`char_count` 반환 |
| `set_cell` | 특정 표의 `(row, col)` 셀 값 교체 (병합 anchor만) |
| `append_to_cell` | 기존 텍스트 뒤에 값 덧붙임 (라벨 유지용, 예: `"성 명"` → `"성 명  홍길동"`) |
| `fill_form` (v0.7+) | **라벨 이름**으로 일괄 채우기. 좌표 계산 없이 `{"접수번호": "...", "성명": "..."}` dict. dot-path 섹션 해소 + `overflow_warnings` |
| `append_row` | 표에 새 행 추가 (전 포맷, v0.5+). v0.14: `insert_row`/`insert_column` 어댑터 API 로 위치 지정 삽입 + 서식 상속 |
| `create_document` (v0.15) | **새 문서 생성** — .docx 는 제약된 markdown, .xlsx 는 sheet spec 을 결정적 렌더러가 스타일 잡힌 문서로 변환. 생성 직후 기존 편집 도구와 같은 좌표계로 이어짐 |
| `get_text_map` / `find_text` / `replace_text` / `insert_text` (v0.13, DOCX/HWPX) | 표 밖 본문 텍스트 지도·검색·치환·삽입 (run 분할 무관, 서식 보존) |
| `get_shapes` / `set_shape_text` (v0.8, PPTX) | 표 외 shape(textbox/placeholder/도형) 텍스트 조회·편집 |
| `get_charts` / `set_chart_data` / `add_chart` (v0.17, PPTX) | **차트** 데이터 조회·수치 편집·새 차트 추가. 차트는 tables/shapes 에 안 잡히므로 `get_charts` 가 유일한 진입점. `set_points` 로 이름 기반 부분 수정, 서식(색/축/범례) 보존 |
| `get_slides` / `duplicate_slide` (v0.17, PPTX) | 슬라이드(페이지) 개요 조회 + **양식 슬라이드 복제**(표/차트/이미지 유지, 차트 데이터 독립 복제). 반환 좌표로 곧장 셀/텍스트/차트 편집 |
| `copy_shape` (v0.18, PPTX) | 특정 **표/차트/텍스트박스 하나만** 다른 슬라이드로 복사 (서식 유지, 같은 파일 내). `clear_values` 로 값만 비운 "빈 양식" 복사 가능, 차트는 독립 복제 후 `set_chart_data` 로 수치 변경 |
| `get_form_controls` / `set_form_control` (v0.10, HWPX) | 폼 컨트롤(체크박스·라디오·에디트·콤보) 조회·설정 |
| `diff_documents` (v0.11) | 두 문서(편집 전/후)를 셀 단위 비교 → 변경 셀 before/after + `overflow_risk`. **편집 후 검증용** |

### 새 문서 생성 (v0.15+)

```python
from document_adapter import create_document, load

# DOCX — LLM 은 제약된 markdown 만 쓰면 된다 (OOXML 대비 토큰 ~96% 절감)
create_document("회의록.docx", lang="ko", markdown="""
# 주간 회의록

| 이름 | 부서 |
|---|---|
| 김경윤 | AI플랫폼 |

- 결정: v0.15 배포
""")

# XLSX — sheet spec (숫자는 숫자로, `=` 는 살아있는 수식)
create_document("매출.xlsx", sheets=[
    {"name": "매출 요약", "headers": ["분기", "매출(억원)"],
     "rows": [["1분기", 120], ["합계", "=SUM(B2:B2)"]],
     "number_formats": {"B": "#,##0"}},
])

# 생성-편집 왕복: 생성 직후 같은 좌표계로 편집 가능
doc = load("회의록.docx")
doc.insert_row(0, ["유지수", "기획"], at_row=2)   # 표에 행 추가
doc.save(); doc.close()
```

지원 markdown 서브셋: `#`~`######` 헤딩 · 문단 · `-` 불릿 · `1.` 번호 ·
`**굵게**` `*기울임*` `` `코드` `` · 파이프 표 · `>` 인용 · `---` 수평선 ·
코드펜스. (HTML/이미지/각주/중첩 리스트는 미지원 — 일반 텍스트로 관용 처리.
정교한 양식은 생성이 아니라 **기존 템플릿 편집**이 이 라이브러리의 철학입니다.)
HWPX 생성은 v0.16 예정.

### `inspect_document` 반환 예시 (v0.2+)

```json
{
  "format": "hwpx",
  "source": "/path/to/form.hwpx",
  "placeholders": [],
  "tables": [
    {
      "index": 0,
      "rows": 28,
      "cols": 16,
      "location": null,
      "parent_path": null,
      "preview": [
        ["포상금 지급신청서", null, null, null, null, null, null, null, null, null, null, null, null, null, null, null],
        ["접수번호", null, null, "", "접수일자", null, null, "", ...]
      ],
      "merges": [
        {"anchor": [0, 0], "span": [1, 16]},
        {"anchor": [1, 0], "span": [1, 3]}
      ]
    }
  ]
}
```

LLM은 이 preview를 보고 **"빈 셀이 어디 있는지 / 어떤 값을 넣어야 하는지"** 를 판단하여 `set_cell` / `append_to_cell`을 호출합니다. `null` 슬롯은 병합된 영역이며 `merges`의 anchor 좌표로만 쓸 수 있습니다.

## 템플릿 작성 규칙

### DOCX — Jinja2 전체 문법 사용 가능

```
{{ report_title }}
작성자: {{ author }}

{% for item in items %}- {{ item.name }}: {{ item.value }}
{% endfor %}
```

**표 행 반복은 `{%tr for ... %}` / `{%tr endfor %}`를 각각 별도 행에 두어야 합니다.**
같은 행에 두 태그를 넣으면 `<w:tr>` 전체가 `{% for %}`로 교체되어 `endfor`가 손실됩니다.

```
┌─────────────────────┬─────┬─────┐
│ 항목                │ 목표 │ 실적 │    <- 헤더
├─────────────────────┼─────┼─────┤
│ {%tr for r in rows %}          │    <- for 행
├─────────────────────┼─────┼─────┤
│ {{ r.name }}        │ {{ r.target }} │ {{ r.actual }} │  <- 반복 본문
├─────────────────────┼─────┼─────┤
│ {%tr endfor %}                 │    <- endfor 행
└─────────────────────┴─────┴─────┘
```

### PPTX / HWPX / XLSX — `{{key}}` + 조건/표현식 (v0.11+)

단순 `{{key}}` 치환에 더해 **조건(`{% if %}`)·표현식(`{{ price * qty }}`)·필터(`{{ x|length }}`)** 를 지원합니다 (셀/문단 단위 jinja2 렌더 — docx 의 Jinja 와 통일). 단순 `{{key}}` 만 있을 때는 기존 빠른 치환 경로를 그대로 씁니다.

- **제약**: 표 행 반복(`{%tr for%}` 같은 구조적 row 루프)은 DOCX(docxtpl)만 지원합니다. PPTX/HWPX/XLSX 의 조건/표현식은 *한 셀/문단 블록 내부* 범위입니다.
- PPTX는 placeholder가 여러 `run`으로 쪼개질 수 있어, 어댑터가 paragraph 전체 텍스트를 재조립한 뒤 첫 `run`에 다시 담는 방식으로 처리합니다 (서식 일부 손실 가능).

### 누락 키 처리 — 3포맷 동일 (`on_missing`)

`render_template(context, on_missing=...)` 은 **context에 없는 키**를 3포맷 동일하게 처리합니다:

- `"blank"`(기본): 빈 문자열로 — 출력에 `{{key}}` 리터럴을 남기지 않음
- `"leave"`: `{{key}}` 를 그대로 둠
- `"error"`: 누락 키가 있으면 `ValueError` (문서 미변경)

반환값은 `{"used": [...], "missing": [...]}` 로, 어떤 키가 치환됐고 누락됐는지 알려줍니다 (`fill_form` 의 `filled`/`not_found` 와 대칭).

### save() 의 byte 안정성 — 포맷별 차이

- **HWPX**: 수정한 파트만 재직렬화하고 나머지는 원본 bytes 그대로 복사 → **편집 안 한 부분은 byte-identical** (최소 diff). 원본 서식/메타 보존에 유리.
- **DOCX / PPTX**: `python-docx` / `python-pptx` 가 저장 시 패키지 전체를 재작성 → byte-identical 보장 안 됨 (내용은 보존되나 XML 직렬화가 달라질 수 있음).

원본과의 최소 diff가 중요한 워크플로우(버전 관리·서명 등)에서는 이 차이를 고려하세요.

## 내장된 버그 회피 / 백엔드 선택

| 포맷 | 문제 | 어댑터의 처리 |
|---|---|---|
| HWPX | `python-hwpx` 가 Non-Commercial License → 상용 배포 블로커 | **v0.4.0 부터 자체 `hwpx_core` 모듈** (zipfile + lxml) 로 교체. 런타임에 `python-hwpx` 불필요. 테스트 fixture 생성에만 사용 (dev extras) |
| PPTX | `python-pptx` 에 공식 `add_row` API 없음 (issue #86, 2014년부터 open) | **v0.5.0 부터 자체 lxml 구현** (`<a:tr>` deepcopy 패턴) |
| PPTX | placeholder가 여러 `run`으로 쪼개져 단순 `run.text` 치환이 실패 | paragraph 전체 재조립 |
| DOCX | `docxtpl`의 `{%tr%}`를 같은 행에 두면 파싱 에러 | README에 배치 규칙 명시 |

## 프로젝트 구조

```
document_adapter/
├── __init__.py        # load() dispatcher
├── base.py            # DocumentAdapter ABC + fill_form + dataclasses
├── docx_adapter.py    # DocxAdapter
├── pptx_adapter.py    # PptxAdapter (append_row 자체 구현 포함)
├── hwpx_adapter.py    # HwpxAdapter (hwpx_core 기반)
├── hwpx_core/         # 자체 HWPX 패키지 (v0.4+)
│   ├── constants.py
│   ├── package.py     # ZIP + dirty XML 관리
│   ├── grid.py        # iter_grid, table_shape
│   └── paragraph.py   # run-level 편집 헬퍼
├── tools.py           # 23개 MCP 도구 정의 + call_tool dispatcher
└── mcp_server.py      # MCP stdio server

examples/
└── claude_api_example.py    # Claude API Tool Use 에이전트 루프
```

## 릴리스 (PyPI 배포)

리포에 write 권한이 있으면 **누구나** 배포할 수 있다. PyPI 토큰은 필요 없다 —
`.github/workflows/release.yml` 이 GitHub Actions의 Trusted Publishing(OIDC)으로 업로드한다.

```bash
# 1. pyproject.toml 의 version 상향 → CHANGELOG.md 갱신 → main 에 머지
# 2. 태그를 밀면 끝
git tag v0.19.0
git push origin v0.19.0
```

태그 push → 테스트 → 빌드 → PyPI 업로드 → GitHub Release 생성까지 자동 진행된다.

- 태그와 `pyproject.toml` 의 버전이 다르면 첫 단계에서 실패한다 (오배포 방지).
- 배포 단계는 GitHub Environment `pypi` 를 거친다. 승인자가 지정돼 있으면 승인 후 업로드된다.
- 진행 상황: [Actions 탭](https://github.com/PlateerLab/document-adapter/actions) 의 `release` 워크플로우.
- 이미 PyPI에 있는 버전은 건너뛴다 (`skip-existing`). 재배포하려면 버전을 올려야 한다.

## 라이선스

**Apache License 2.0** — [`LICENSE`](LICENSE) / [`NOTICE`](NOTICE) 참조.
이 소프트웨어나 그 파생물을 **재배포**할 때는 저작권·라이선스 고지와 `NOTICE`
파일의 출처 표시를 유지해야 합니다(라이선스 §4). 단순히 의존성으로 사용만 하는
경우엔 별도 표시 의무가 없습니다(표준 OSS 공통).

## Credits

**런타임 의존성** (전부 허용형 OSS):
- [`python-docx`](https://github.com/python-openxml/python-docx) — MIT
- [`docxtpl`](https://github.com/elapouya/python-docx-template) — LGPL-2.1
- [`python-pptx`](https://github.com/scanny/python-pptx) — MIT
- [`lxml`](https://lxml.de/) — BSD
- [`mcp`](https://github.com/modelcontextprotocol/python-sdk) — MIT

**코드 참조**:
- [`xgen-doc2chunk`](https://github.com/PlateerLab/xgen-doc2chunk) (Apache-2.0) — HWPX table grid 파싱 로직 차용 (`NOTICE` 참조)

**Dev 전용** (fixture 생성에만 사용):
- [`python-hwpx`](https://github.com/airmang/python-hwpx) — Non-Commercial License (v0.4.0 부터 런타임 의존성 제거)
