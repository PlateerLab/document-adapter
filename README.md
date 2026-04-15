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

| 포맷 | 백엔드 | 템플릿 렌더 | 표 읽기 | 셀 수정 | 행 추가 |
|---|---|---|---|---|---|
| `.docx` | `docxtpl` + `python-docx` | Jinja2 (`{%tr%}` loop 포함) | ✅ | ✅ | ✅ |
| `.pptx` | `python-pptx` | `{{key}}` 치환 | ✅ (슬라이드 위치 포함) | ✅ | ❌ (미지원) |
| `.hwpx` | `python-hwpx` (Pure Python) | `{{key}}` 치환 | ✅ | ✅ | ❌ (미지원) |

- HWPX는 한컴오피스 설치가 **불필요**합니다 (macOS/Linux 서버에서 그대로 동작).
- 구버전 `.hwp`(바이너리 포맷)는 지원하지 않습니다 — `.hwpx`로 변환 후 사용하세요.

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
old = doc.set_cell(table_index=1, row=1, col=1, value="○○전자")
doc.append_row(1, ["새 항목", "값"])  # DOCX만 지원
doc.save("checklist_filled.docx")
doc.close()
```

확장자로 자동 분기되므로 `.pptx` / `.hwpx`도 동일한 API를 사용합니다.

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

재시작하면 Claude Desktop에서 아래 4개 도구를 사용할 수 있습니다.

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

## 노출되는 4개 도구

| 도구 | 설명 |
|---|---|
| `inspect_document` | 문서 구조(placeholders, tables)를 JSON으로 반환. **항상 첫 호출로 사용** |
| `render_template` | `{{key}}`를 context dict 값으로 치환해 새 파일 저장 |
| `set_cell` | 특정 표의 `(row, col)` 셀 값 교체 |
| `append_row` | 표 끝에 새 행 추가 (DOCX 전용) |

### `inspect_document` 반환 예시

```json
{
  "format": "docx",
  "source": "/path/to/checklist.docx",
  "placeholders": [],
  "tables": [
    {
      "index": 1,
      "rows": 7,
      "cols": 2,
      "location": null,
      "preview": [
        {"row": 0, "cells": ["항목", "기입 내용"]},
        {"row": 1, "cells": ["고객사 / 조직", ""]},
        {"row": 2, "cells": ["현업 담당부서 / 책임자", ""]}
      ]
    }
  ]
}
```

LLM은 이 preview를 보고 **"빈 셀이 어디 있는지 / 어떤 값을 넣어야 하는지"** 를 판단하여 `set_cell`을 호출합니다.

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

### PPTX / HWPX — 단순 `{{key}}` 치환만

loop / if / filter는 지원하지 않습니다. PPTX는 placeholder가 여러 `run`으로 쪼개질 수 있어, 어댑터가 paragraph 전체 텍스트를 재조립한 뒤 첫 `run`에 다시 담는 방식으로 처리합니다 (서식 일부 손실 가능).

## 내장된 버그 회피

| 포맷 | 문제 | 어댑터의 처리 |
|---|---|---|
| HWPX | `python-hwpx 2.9.0`의 `set_cell_text()`가 빈 셀에서 lxml/ElementTree 혼용 `TypeError` 발생 | `paragraphs[0].text = value` 직접 할당으로 우회 |
| HWPX | `replace_text_in_runs()`가 한글 공백이 run으로 쪼개진 경우 매칭 실패 | 위치 기반 API만 사용 |
| HWPX | `manifest fallback` 경고 로그가 과도하게 출력됨 | `logging.getLogger("hwpx")` 레벨을 `ERROR`로 조정 |
| PPTX | placeholder가 여러 `run`으로 쪼개져 단순 `run.text` 치환이 실패 | paragraph 전체 재조립 |
| DOCX | `docxtpl`의 `{%tr%}`를 같은 행에 두면 파싱 에러 | README에 배치 규칙 명시 |

## 프로젝트 구조

```
document_adapter/
├── __init__.py        # load() dispatcher
├── base.py            # DocumentAdapter ABC, TableSchema, DocumentSchema
├── docx_adapter.py    # DocxAdapter
├── pptx_adapter.py    # PptxAdapter
├── hwpx_adapter.py    # HwpxAdapter (버그 회피 포함)
├── tools.py           # Tool 정의 + call_tool dispatcher
└── mcp_server.py      # MCP stdio server

examples/
└── claude_api_example.py
```

## 라이선스

MIT

## Credits

- [`python-docx`](https://github.com/python-openxml/python-docx)
- [`docxtpl`](https://github.com/elapouya/python-docx-template)
- [`python-pptx`](https://github.com/scanny/python-pptx)
- [`python-hwpx`](https://github.com/airmang/python-hwpx)
- [`mcp`](https://github.com/modelcontextprotocol/python-sdk)
