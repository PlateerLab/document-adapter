# Changelog

All notable changes to this project will be documented in this file.

Format: [Keep a Changelog](https://keepachangelog.com/en/1.1.0/). Versioning:
[SemVer](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.18.0] — 2026-07-30

**PPTX shape 단위 복사** — 슬라이드 전체(duplicate_slide)가 아니라 특정
표/차트/텍스트박스 **하나만** 골라 다른 슬라이드로 복사한다. "이 차트를 저
페이지에도" / "이 표 양식만 베껴서" 류 요청 처리.

### Added
- **`copy_shape(target_slide_index, *, table_index | source_slide_index+shape_id,
  x_cm, y_cm, clear_values)`** — shape 단위 복사 (같은 파일 내, 서식·스타일 유지).
  - 표는 `table_index`, 차트/텍스트박스는 `source_slide_index`+`shape_id` 로 지정.
  - `clear_values=True`: 복사본의 표 셀/텍스트 run 값을 비움 (서식·구조 유지 —
    "양식만 베끼기"). 차트는 데이터 유지 복사 (수치는 set_chart_data 로).
  - 차트 part + 내장 워크북 독립 복제 (duplicate_slide 와 동일 — 복사본 편집이
    원본에 영향 없음).
  - placeholder 복사 시 위치/크기를 실측값으로 고정한 일반 shape 로 전환
    (`<p:ph>` 제거) — 대상 슬라이드의 placeholder 와 idx 충돌 방지.
  - cNvPr id 를 대상 슬라이드 내 유일값으로 재부여 (그룹 내부 포함).
  - 반환: 새 shape_id / (표) 재계산된 table_index + preview(행 라벨) /
    (차트) chart_type + set_chart_data 힌트 — 곧장 후속 편집 가능.
- MCP/Claude 도구 **`copy_shape`** 추가 (총 23개).

### Notes
- 같은 파일 내 복사만 지원 — 다른 pptx 파일에서 가져오는 크로스 파일 복사는
  테마/tableStyles 병합이 필요해 미지원 (시도 시 별도 파일을 열어야 하므로
  API 상 표현 불가).
- 그룹 내부 shape 는 단독 복사 불가 — 그룹 전체 shape_id 로 복사.

## [0.17.1] — 2026-07-30

### Fixed
- `duplicate_slide` 반환의 `tables[]` 에 **`preview`(행 라벨 포함) 추가** —
  LLM 이 복제본 표를 채울 때 행 매핑을 기억/추측하다 값이 한 행씩 밀리는
  off-by-one 실패 패턴 방지 (실측: "프로젝트명 행에 담당자 값" 오기입).
  도구 설명에도 "preview 라벨을 확인해 (row, col) 결정 — 추측 금지" 명시.

## [0.17.0] — 2026-07-30

**PPTX 차트 편집 + 슬라이드 복제** — 표/텍스트만 다루던 PPTX 어댑터에
차트 계층과 슬라이드(페이지) 계층을 추가. "덱 안의 양식 페이지를 복제해
새 페이지를 만들고, 표·텍스트·차트 수치를 채우는" 보고서 워크플로가
도구만으로 완결된다. 신규 의존성 없음.

### Added
- **`get_charts(slide_index=None)`** — 차트 목록 + 카테고리/시리즈 수치.
  차트는 `get_tables`/`get_shapes` 에 나타나지 않으므로 차트 작업의 유일한
  진입점. 그룹 shape 안의 차트도 수집. 편집 불가 구조(scatter/bubble·
  날짜축·다중레벨 카테고리·콤보)는 `editable=false` + `warning` 으로 표시.
- **`set_chart_data(slide_index, shape_id, ...)`** — 차트 수치 편집
  (read → 수정 → `replace_data`, 서식/색/축/범례 보존). 두 모드:
  `set_points`(시리즈/카테고리를 **이름**으로 지정하는 부분 수정) /
  `categories`+`series`(전체 교체 — 카테고리·시리즈 개수 변경 포함).
  `title` 단독 지정으로 제목만 변경 가능. 반환에 before/after 스냅샷.
- **`add_chart(slide_index, chart_type, ...)`** — 새 차트 추가 (column/bar/
  line/pie/doughnut/area/radar 계열 11종). 위치·크기(cm) 생략 시 결정적
  기본 배치, 시리즈 2개 이상이면 하단 범례 자동.
- **`get_slides()`** — 슬라이드 개요(레이아웃/제목/표·차트·텍스트 개수) —
  복제할 '양식 페이지' 를 고르는 눈.
- **`duplicate_slide(source_slide_index, at=None)`** — 양식 슬라이드 복제 +
  위치 지정 삽입. 서식/표/이미지 유지, **차트는 chart part + 내장 워크북을
  독립 복제**해 복제본 편집이 원본을 오염시키지 않음. 반환에 새 슬라이드의
  표/차트/텍스트 좌표(삽입 후 재계산) — 곧장 후속 편집 가능. 중간 삽입 시
  전역 `table_index` 시프트 warning 포함. 노트 슬라이드는 미복제.
- MCP/Claude 도구 5종 추가 (총 22개): `get_charts` / `set_chart_data` /
  `add_chart` / `get_slides` / `duplicate_slide`.
- `inspect_document`(PPTX): `chart_summary`(차트 존재 + set_chart_data 유도
  힌트, 수치는 미포함 — 토큰 절약) 와 `slide_count` 추가.
- `diff_documents`: 차트 수치 변경을 `chart_changes` / `charts_added` /
  `charts_removed` 로 보고 (차트 없는 문서는 키 생략 — 기존 반환 형태 보존).

### Notes
- `set_chart_data` 는 python-pptx `replace_data` 기반 — 차트 XML 캐시와
  내장 워크북을 함께 갱신하지만, 내장 워크북의 차트 외 시트/수식은 보존되지
  않는다 (python-pptx 의 알려진 동작).
- 3D/stock/surface 차트 생성은 미지원 (python-pptx 한계). scatter/bubble
  편집은 v2 후보.
- `diff_documents` 의 표/차트 매칭은 위치 기준(flat index / slide_index) —
  슬라이드·표 **삽입/삭제를 동반한** 비교에서는 밀린 위치끼리 비교될 수 있다
  (셀/수치 편집 전후 검증 용도로 설계됨. 삽입 후 검증은 duplicate_slide
  반환 좌표를 사용할 것).
- eval 하니스의 차트 채점(FieldExpectation 확장)은 known gap — v0.18 후보.

## [0.16.0] — 2026-07-13

**생성 대상 확장 — .pptx / .hwpx** — v0.15 의 docx/xlsx 생성 계층을
프레젠테이션·한글 문서로 넓혔다. markdown 렌더러 3종(.docx/.pptx/.hwpx)이
같은 `markdown_parser`(블록 IR)를 공유해, 하나의 markdown 입력이 포맷만
바꿔 재사용된다. 생성 직후 `load()` 왕복은 신규 포맷에도 동일하게 성립한다.

### Added
- **`.pptx` 생성** (`generate/pptx_writer.py`) — `create_document("x.pptx",
  markdown=...)`. `---` 또는 레벨 1~2 헤딩이 슬라이드 경계, 나머지 블록
  (불릿/문단/표)은 본문 플레이스홀더로 배치. python-pptx 내장 레이아웃 사용.
- **`.hwpx` 생성** (`generate/hwpx_writer.py`) — `create_document("x.hwpx",
  markdown=...)`. v0.15 에서 "v0.16 예정" 안내 에러였던 경로가 실제 렌더러로
  대체됨. HWPX(OWPML) 패키지를 직접 조립, `load()` 가 곧바로 성립.
- `create_document` 디스패처의 markdown 대상이 `{.docx, .pptx, .hwpx}` 로 확장.

## [0.15.0] — 2026-07-09

**문서 생성** — 편집 전용이던 엔진에 "무에서 생성" 계층 추가. LLM 은 경량
중간 산출물(제약된 markdown / sheet spec dict)만 쓰고 결정적 렌더러가
스타일 잡힌 문서로 변환한다 (OOXML 직접 생성 대비 토큰 ~96% 절감).
생성 직후 `load()` 가 성립해 기존 편집 도구(set_cell / insert_row / ...)와
같은 좌표계로 이어진다 — 생성-편집 왕복 보장.

### Added
- **`create_document(path, *, markdown=None, sheets=None, lang="ko",
  overwrite=False)`** — 확장자 디스패치 생성 진입점 (`load()` 와 대칭).
  `.docx`=markdown, `.xlsx`=sheet spec. `.hwpx` 는 v0.16 예정 안내 에러.
- **`generate/` 서브패키지** — `markdown_parser`(블록 IR: 포맷 무관 공용
  파서 — 이후 HWPX writer 가 공유), `docx_writer`(python-docx 내장 스타일만
  사용, CJK `w:eastAsia` 폰트 명시, hr=pBdr), `xlsx_writer`(헤더 스타일·
  틀고정·자동 열폭·`=` 수식 통과·`number_formats`).
- MCP/Claude 도구 **`create_document`** (총 17개) — 반환값에 생성 직후 표
  좌표 요약(`tables`/`table_shapes`) 포함, LLM 이 후속 편집을 바로 이어감.
- markdown 서브셋: 헤딩/문단/불릿/번호/굵게/기울임/코드/파이프표/인용/
  수평선/코드펜스. 지원 외 문법은 에러 없이 일반 텍스트 관용 처리.
- 모든 스펙 검증 실패는 이중어(EN/KO) `ValueError` — 호출 레이어의 1회
  재시도 계약용 (렌더러가 곧 검증기).

## [0.14.0] — 2026-07-07

표에 **위치를 지정해** 행/열을 삽입하는 계층 추가 — "2026년 행을 2025년
위에 넣어줘", "3분기 옆에 4분기 열 추가" 류 요청을 처리한다. 위치 결정은
호출자(LLM)의 몫이고, 엔진은 서식 상속과 병합 안전성만 보장한다 (엔진이
표 내용을 해석해 위치를 추론하지 않음 — 정렬 오판 방지 설계).

### Added
- **`insert_row(table_index, values, at_row=None)`** — 지정 위치에 행 삽입
  (DOCX/HWPX/PPTX). 인접 행 deepcopy 로 음영·테두리·행높이·폰트 상속.
  세로 병합이 삽입 경계를 가로지르면 `NotImplementedForFormat`.
  `at_row=None` 은 맨 끝(append 동일). values 는 논리 grid 열 기준.
- **`insert_column(table_index, values, at_col=None)`** — 지정 위치에 열
  삽입 (DOCX/HWPX/PPTX, 신규 능력). 행별 인접 셀 deepcopy 로 서식 상속
  (헤더 행은 헤더 서식, 데이터 행은 데이터 서식). 표 전체 폭 유지를 위해
  기존 열 폭 비례 축소(tblGrid/tcW·cellSz·gridCol). 가로 병합 경계 교차
  시 `NotImplementedForFormat`. 미구현 포맷(XLSX)은 base 기본값이 명시적
  에러.

### Fixed
- **DOCX `append_row` 서식 미상속** — python-docx `add_row()` 가 음영·
  테두리·행높이·폰트 없는 기본 스타일 빈 행을 만들던 문제. 마지막 행
  deepcopy 방식(insert_row 위임)으로 HWPX 와 동작 통일.

## [0.13.0] — 2026-07-06

표 좌표(set_cell)·`{{placeholder}}`(render_template) 없이도 **문서 전역의
임의 텍스트**를 찾고/치환하고/삽입하는 계층 추가 — "홍길동을 유지수로
바꿔줘", "결재란 부분의 담당자 옆에 유지수라고 써줘" 류 요청을 LLM 이
직접 처리할 수 있게 한다. (DOCX/HWPX)

### Added
- **`textops.splice_runs` / `find_spans`** — run 분할("홍길"+"동")에도
  concat 오프셋으로 매치하고, 매치에 걸친 run 만 수정해 run-level 서식
  (rPr/charPr)을 보존하는 포맷 독립 순수 알고리즘. 삽입은 앵커 구간
  치환으로 구현돼 앵커 run 의 서식을 자동 상속.
- **`get_text_map()`** — 본문+표 셀+머리말/꼬리말 문단 지도(문서 순서).
  scope/location/is_heading/truncation, offset/limit 페이징. LLM 이 표 밖
  텍스트 구조를 파악하는 '눈'.
- **`find_text(query, whole_word, scope)`** — 전역 검색. 매치마다
  «» context, nearest_heading, context_before 를 제공해 "~부분의" 위치
  참조를 해소. `whole_word` 로 "홍길동님" 속 부분 일치 배제.
- **`replace_text(old, new, occurrences, whole_word, scope)`** — 전역
  치환 (기본 전부, `occurrences` 로 특정 등장만). 문단 내 다중 매치는
  뒤→앞 splice 로 오프셋 안정.
- **`insert_text(anchor, text, position, separator, ...)`** — 앵커
  앞/뒤 삽입 (앵커 서식 상속).
- MCP/Claude 도구 4 종 추가: `get_text_map` / `find_text` /
  `replace_text` / `insert_text` (총 16 개).
- 어댑터 훅 `_iter_text_paragraphs()` — DOCX 는 `iter_inner_content` 로
  본문·표를 문서 순서 그대로(제목-표 연관 유지), hyperlink 내부 run 포함,
  linked 머리말/꼬리말 제외. HWPX 는 `root.iter(hp:p)` 단일 경로에 표
  flat index 좌표계 location + 변경 섹션 `mark_dirty`. PPTX/XLSX 는
  `NotImplementedForFormat`.

### Fixed
- HWPX 표 flat index 매핑 시 lxml 프록시 GC 로 `id()` 가 불안정해지는
  문제 — 프록시를 순회 수명 동안 유지해 좌표계 일치 보장.

## [0.12.0] — 2026-06-02

### Changed
- **라이선스를 MIT → Apache License 2.0 으로 변경.** 재배포·파생물 배포 시
  저작권·`NOTICE` 출처 표시 유지를 명시적으로 요구(라이선스 §4)하고, 특허
  사용권을 포함한다. `LICENSE`(Apache 2.0 전문)·`NOTICE`·pyproject classifier
  갱신. (이전 0.x 릴리스는 MIT 로 배포된 상태 그대로 유지 — 0.12.0 부터 Apache-2.0.)

## [0.11.1] — 2026-06-02

### Docs
- README '노출되는 도구' 표를 **12 개 전부**로 갱신(누락됐던 get_shapes/
  set_shape_text·get_form_controls/set_form_control·diff_documents 반영) +
  render_template on_missing/조건식·inspect duplicate_labels 등 최신 반환 필드.
  (코드 변경 없음 — PyPI 랜딩 페이지 정합성 패치.)

## [0.11.0] — 2026-06-02

리서치 기반 궁극 로드맵의 **신뢰성 해자 + 생성급 parity 1차** 를 구현.

### Added
- **PPTX/HWPX/XLSX 조건·표현식·필터** — 단순 `{{key}}` 치환에 더해
  `{% if %}` 조건, `{{ price * qty }}` 표현식, `{{ x|length }}` 필터를 지원
  (셀/문단 단위 jinja2, docx 의 Jinja 와 통일). 단순 `{{key}}` 는 기존 빠른
  경로 보존. 구조적 표 행 루프(`{%tr%}`)는 여전히 DOCX 전용.
- **`get_form_fields()` + inspect 중복 라벨 힌트** — 라벨 후보와 *중복(dot-path 필요)*
  여부를 미리 보여준다. `inspect_document` 응답에 `duplicate_labels` + 힌트를 추가해
  LLM 이 fill_form 전에 dot-path 필요를 인지(ambiguous 재시도 라운드 절감).
- **`diff_documents(path_a, path_b)`** — 편집 전/후 문서를 셀 단위로 비교해 변경된
  셀의 before/after 와 overflow_risk 를 반환하는 **검증 도구**(MCP 도구 → 총 12 개).
  fill 후 원본과 diff 해 "무엇이 어디서 바뀌었고 깨짐 위험은 없나" 를 LLM 이 스스로
  확인·자가교정할 수 있다. 4 포맷 공통.
- **xlsx 수식 계산값**: 캐시된 계산값이 있으면 그 값을, 없으면 수식 문자열을 표시
  (`data_only` lazy 폴백). 수식은 편집/저장 시 보존.

## [0.10.0] — 2026-06-01

코드 감사로 포맷별(docx/pptx/xlsx) 격차를 점검해 **Excel 지원을 신규 추가**하고,
inspect 가 병합 docx 에서 깨지던 버그와 머리말/꼬리말·노트 플레이스홀더 누락을 수정.

### Added
- **Excel(`.xlsx`) 지원 — `XlsxAdapter`** (openpyxl 기반). 각 워크시트를 하나의 표로
  매핑(`table_index`=시트 인덱스, `location`=시트명). `get_tables`/`get_cell`/
  `set_cell`/`append_to_cell`/`append_row`/`render_template` 구현, `fill_form` 은
  base 구현으로 자동 동작. 병합 셀 anchor/span 인지(비-anchor 쓰기는
  `MergedCellWriteError`), 셀 크기(cm) 메타 제공. `load("*.xlsx")` 자동 디스패치.
  MCP 도구는 확장자 디스패치로 그대로 동작.
  - **셀 타입 처리**: 날짜는 시간 없이 표시(`2026-06-01`), `set_cell` 은 깔끔한
    숫자(금액 등)를 숫자형으로 기록해 Excel 수식/합계를 유지하되, 전화·사번·
    우편번호(대시·선행 0)는 문자로 보존. 실제 xlsx 파일 5종(다중 시트 포함)으로 검증.

### Fixed
- **docx `get_placeholders` 병합표 크래시**: `row.cells` 가 가로+세로 병합 docx 에서
  `ValueError` 로 깨져 `inspect_document`/`get_schema` 가 실패하던 문제 —
  `_build_grid` anchor 셀 순회로 수정(`get_tables` 와 동일 견고 경로).
- **머리말/꼬리말·노트 플레이스홀더 누락**: docx `get_placeholders` 가 본문만 보고
  머리말/꼬리말을, pptx 가 슬라이드 노트를 놓쳐 `render` 는 채우는데 `inspect`/
  `used`/`missing` 에 안 잡히던 불일치 수정.

### Changed
- 런타임 의존성에 `openpyxl>=3.1` 추가.

### Verified
- xlsx 폼(병합 헤더·라벨-값·템플릿) inspect/fill_form/render/round-trip + MCP 경로,
  docx 머리말/꼬리말·pptx 노트 커버리지 회귀 테스트. 테스트 77 종, ruff·mypy 클린.

## [0.9.0] — 2026-06-01

실제 공공서식(지급정지요청서 등)과 다운로드한 docx/hwpx 폼들로 검증하며 드러난
결함을 수정하고, 폼 컨트롤·LLM 평가 하니스·render_template 일관화를 추가했다.
대부분의 변경이 **실증 또는 실제 폼이 드러낸 결함**에 기반한다.

### Added
- **폼 컨트롤 지원** (`get_form_controls` / `set_form_control`) — 표가 아닌
  인터랙티브 필드(체크박스·라디오·에디트·콤보·리스트)를 읽고 채운다.
  comboBox/listBox 는 옵션(`items`)과 현재값을 노출. MCP 도구 2 개 신규 → **총 11 개**.
- **LLM 주도 평가 하니스** (`document_adapter.eval`) — `ModelBackend`(pluggable),
  `run_scenario`/`evaluate`(결과 기반 채점: 필드 정합 + 라벨 보호 + 오버플로),
  공개 합성 양식 시나리오. 가짜 백엔드로 채점 로직을 API 없이 결정적으로 검증.
  실제 모델 러너 예시(`examples/eval_run.py`, Ollama/OpenAI·vLLM 호환).
- **`render_template(context, on_missing=...)`** — 누락 키 처리 정책
  (`blank`/`leave`/`error`)을 3 포맷 동일하게 제어. `{"used", "missing"}` 반환.
- **`fill_form` 오버플로 인지** — 값이 칸 너비를 넘겨 깨질 위험을 `overflow_risk`
  플래그 + `overflow_warnings` 로 보고 (`width_cm` 활용).

### Fixed
- **보안**: `hwpx_core` XML 파서에 XXE / billion-laughs 하드닝
  (`resolve_entities=False`, `no_network`, `load_dtd=False`).
- **DOCX 병합 표 크래시**: 가로(gridSpan)+세로(vMerge) 병합이 섞인 표에서
  `python-docx` 의 `row.cells` 가 `ValueError` 로 깨지던 문제 — `_build_grid` 를
  OOXML 레이어에서 직접 계산하도록 재작성. (실제 docx 폼에서 발견)
- **`render_template` 포맷 불일치**: pptx/hwpx 가 누락 키를 `{{missing}}` 리터럴로
  출력에 노출하던 문제 — `on_missing="blank"`(기본)로 3 포맷 동일하게 정렬.
- **`fill_form` 값셀 선택**: 라벨과 값 영역 사이의 얇은 스페이서 칸에 값이 들어가
  세로로 깨지던 문제(실제 지급정지요청서 접수일자) — `width_cm` 기준 가장 넓은
  값칸 선택, `below` 방향도 대칭. 스캔은 `ncols`/`nrows` 로 bound(무한루프 방지).
- **ABC 계약 정합**: `set_cell`/`append_to_cell` 추상 시그니처에
  `allow_merge_redirect` 누락 — ABC 만 보고 구현한 어댑터에서 `fill_form` 이
  `TypeError` 로 깨지던 문제 수정.
- **타입 정직화**: `TableSchema.column_widths_cm` / `row_heights_cm` 를
  `list[float | None]` 로 정정(병합 컬럼은 `None` 가능).
- **hwpx comboBox/listBox**: 현재값을 `<text>` 자식에서, 옵션을 `listItem` 에서
  올바로 처리(기존엔 존재하지 않는 `value` 속성을 읽어 항상 빈 값).

### Changed
- **CI**: `ruff` + `mypy` 잡 추가, `tests/` 전체 실행(이전엔 smoke 만).
- **`render_template` 반환 타입**: `None` → `{"used", "missing"}`.
  MCP `render_template` 응답: `placeholders_after` 제거, `rendered_keys` /
  `missing_keys` 추가.
- **`save()` byte 안정성** 문서화: HWPX 는 미편집 파트 byte-identical,
  DOCX/PPTX 는 패키지 전체 재작성(README 참조).

### Verified
- H100 vLLM(Qwen3.6-27B)로 실제 지급정지요청서(28×16, 병합 57) end-to-end
  **5/5 PASS** (필드 정합·overflow 0·라벨 무손상, dot-path 자동).
- 다운로드한 실제 docx/hwpx 폼 다수 + 7.3MB/110 표 정부보고서 round-trip(byte 안정).
- 테스트 72 종, `ruff`·`mypy` 클린.

### Migration
- `render_template` 가 이제 dict 를 반환한다(기존엔 `None`). 반환값을 무시하던
  호출 코드는 영향 없음. MCP `render_template` 응답에서 `placeholders_after` 를
  쓰던 코드는 `missing_keys` 로 변경.

## [0.8.1] — 2026-04-17

### Added (LLM UX — inspect_document PPTX shape 가시성)
- **`inspect_document` 응답에 `shape_summary`** (PPTX 전용). 실전 로그 분석 결과
  에이전트가 `set_shape_text` 를 한 번도 호출하지 않고 표만 채운 뒤 "완성" 으로
  종결하는 패턴 관찰. `inspect_document` 가 표 위주 정보만 주어서 에이전트가
  수십 개의 빈 textbox 존재를 **인지조차 못하는 상태** 가 원인.
  - `total_shapes` / `empty_shapes` / `filled_shapes` 항상 포함
  - 빈 shape 비율 > 50% 인 경우 `hint` 필드에 **강한 경고** 추가: "표만
    set_cell 로 채우면 보고서 대부분이 비어있다. get_shapes + set_shape_text
    를 호출하라."
  - 빈 shape 가 일부만 있으면 온건한 힌트.
- `inspect_document` MCP tool description 에 shape_summary/hint 언급 추가.

### Notes
- DOCX / HWPX 는 `get_shapes()` 가 빈 리스트를 반환하므로 shape_summary 자동 생략.
- 기존 API 100% 호환 — JSON payload 에 필드가 추가될 뿐.

## [0.8.0] — 2026-04-17

### Added (PPTX shape-level 편집 — 표 외 텍스트 직접 편집)
- **`get_shapes(slide_index=None, min_text_len=1)`** — PPTX 의 표 외 shape
  (textbox / placeholder / 도형 내 텍스트) 목록. 각 항목에 `slide_index` (1-based),
  `shape_id` (슬라이드 내 고유 숫자), `name`, `kind`, `text` / `text_preview`,
  `placeholder_type` 포함.
- **`set_shape_text(slide_index, shape_id, text)`** — shape 의 텍스트를 교체.
  run-level 포맷 (폰트/크기/색상) 보존.
- MCP 도구 2 개 신규 등록 → 총 **9 개 도구**.
- `ShapeInfo` dataclass 추가 (`base.py`).

### Motivation
실전 공공 PPTX (양산상공회의소 17슬라이드 정책 발표) 에 LLM 이 fill_form 만
호출하면 표 4 개만 채워지고 **나머지 13 슬라이드의 shape 텍스트 (textbox /
placeholder / 도형 라벨)** 는 건드릴 수 없었음. 보고서형 PPTX 편집 완성도를
위해 shape-level API 추가. 17 슬라이드 × 90 개 shape 까지 편집 가능.

### Notes
- DOCX / HWPX 는 `get_shapes()` 빈 리스트, `set_shape_text()` 는
  `NotImplementedForFormat` — 해당 포맷은 표와 paragraph 중심이라 shape 편집
  개념이 약함.

## [0.7.3] — 2026-04-17

### Fixed
- **HWPX `<hp:ctrl>` 내부 테이블 누락 수정** — 기존 `_iter_tables()` 가
  `root > hp:p > hp:run > hp:tbl` 직접 경로만 훑어서 `<hp:ctrl>` (header /
  footer / footNote / endNote / 도형 등) **내부에 포함된 테이블을 놓쳤음**.
  이제 섹션의 전체 `<hp:tbl>` descendant 중 cell 내부가 아닌 것을 top-level 로
  인식. nested 테이블은 기존 방식대로 anchor cell 재귀로 처리.
- 실전 fixture 10 개에는 해당 케이스가 없었으나 합성 테스트로 재현/수정 확인.

### Added
- smoke test `test_hwpx_ctrl_embedded_table_is_found` — 회귀 방지 (기본 표 +
  `<hp:ctrl>` 로 감싼 표 총 2 개 전부 발견 + 편집 가능).

## [0.7.2] — 2026-04-17

### Changed (Docs only)
- **`fill_form` tool description 강화** — 실측 LLM 실패 패턴 기반. Ollama
  qwen3.5:4b (4B 모델) 가 복잡한 시나리오에서 tool call 대신 응답 텍스트에
  `"""json {"fill_form": {...}}"""` JSON 코드블록을 적어 실행되지 않던 문제 관찰.
  description 에 다음 추가로 극적 개선 확인 (tool call 실패 → 1회 호출로 해소):
  - "⚠ 반드시 이 도구를 호출" 경고문
  - direction 선택 기준 (빈 양식 → auto / 예시값 있는 양식 → right) 명확화
  - `output_path` 기본 동작 ("생략 시 원본 덮어쓰기") 명시
  - dot-path 예시 `{'피해자.금액': '...', '지급정지.금액': '...'}`
- `examples/claude_api_example.py`, `examples/ollama_example.py` SYSTEM 프롬프트
  도 동일 원칙 반영.

### Added
- `examples/ollama_example.py` — Ollama 네이티브 SDK 로 document-adapter 7 도구
  에이전트 루프 (qwen2.5:14b 등, 로컬 OSS 스택).
- `scripts/ollama_scenarios.py` — 모델 × 시나리오 매트릭스 실험 runner.
- `.github/workflows/tests.yml` — Python 3.10/3.11/3.12 smoke + build CI.

## [0.7.1] — 2026-04-16

### Added
- **fill_form ambiguous UX 개선**
  - 반환 `ambiguous[].candidates` 가 `(t, r, c)` 튜플 대신 `{table_index, row, col,
    context}` dict. `context` 는 candidate 셀의 섹션 컨텍스트 (같은 표의 가까운
    col=0 anchor 라벨 — HWPX 에서 "피해자정 보", "지급정지요청계좌" 같은 rowSpan
    섹션 헤더 자동 추출).
  - `ambiguous[].hint` 필드 추가 — dot-path 재호출 예시 템플릿 제공.
- **Dot-path 섹션 해소** — `fill_form({"재무.금액": "...", "영업.금액": "..."})`.
  section hint 로 candidate 중 매칭되는 것만 선택. normalize 후 substring 매칭.

### Changed
- **`fill_form` auto 기본값이 보수적으로 변경** — 기존 값이 있는 셀은 다른 라벨로
  간주하고 skip, 최종 same append 로 fallback. **이전 v0.7.0 에서 예시값 있는 셀을
  덮어쓰던 동작이 바뀜**. 예시값 덮어쓰기가 목적이면 `direction="right"` 또는
  `"below"` 명시 필요. 양식 문서의 라벨 오염 방지를 우선.
- auto 모드에서 target 이 병합 non-anchor 면 skip (anchor redirect 로 엉뚱한
  스페이서 셀에 쓰이는 것 방지).

### Fixed
- fill_form 이 인접한 다른 라벨을 덮어쓰던 버그 (label_index 전체를 보호 대상으로
  확장).
- candidate context 수집 시 자기 자신 (col=0) 포함하던 버그.

## [0.7.0] — 2026-04-16

### Added
- **`fill_form(data, direction="auto", strict=False)` — 라벨 기반 일괄 채우기 API**
  (프로젝트의 "LLM 이 쓰기 편하게" 목표의 핵심 개선)
  - LLM 이 좌표 (table_index/row/col) 계산 없이 "접수번호", "성명" 같은 라벨
    key-value dict 로 양식 채움
  - auto 모드: 라벨 셀 오른쪽 → 아래 → 같은 셀 순으로 값 셀 탐색
  - 사용자 요청 라벨 간 cross-check 로 인접 라벨 오염 방지
    (예: `{"성명": "...", "주소": "..."}` 함께 넘기면 "성명" 옆의 "주소" 덮어쓰기 방지)
  - same 셀 fallback 시 append_to_cell 로 라벨 뒤에 값 덧붙임
  - 반환: `{filled: [...], not_found: [...], ambiguous: [...]}`
- HWPX 셀 크기 메타 (v0.6.0 에서 DOCX/PPTX 만 제공했던 것과 대칭)
  - `<hp:cellSz width height>` (HU → cm) 파싱
- MCP `tools.py` 에 `fill_form` 도구 등록 (JSON payload 로 그대로 노출)
- smoke test 5 건 추가 (DOCX/PPTX 라벨-값 분리, not_found, strict, label 정규화)

### Added (Docs)
- CHANGELOG.md 신규
- README: 셀 크기 메타 노트 추가

## [0.6.0] — 2026-04-16

### Added
- **셀 크기 메타 노출** (오버플로 방지 힌트). LLM 이 작은 셀에 긴 텍스트를 넣어
  레이아웃을 깨뜨리는 실패 패턴을 사전에 판단할 수 있도록 추가.
  - `TableSchema.column_widths_cm: list[float] | None`
  - `TableSchema.row_heights_cm: list[float] | None`
  - `CellContent.width_cm: float | None` (anchor + span 영역 합산)
  - `CellContent.height_cm: float | None`
  - `CellContent.char_count: int | None`
- 3 포맷 모두 구현
  - DOCX / PPTX: EMU → cm (1 cm = 360000 EMU) 1자리 반올림
  - HWPX: HU → cm (1 cm ≈ 2834.6457 HU) 1자리 반올림, `<hp:cellSz>` 속성 사용
- MCP 서버 payload 에도 자동 포함 (`to_dict()` 체인)

### Changed
- `to_dict()` 는 `None` 필드를 생략해 JSON payload 간결 유지 (non-breaking)

## [0.5.0] — 2026-04-16

### Added
- **PPTX `append_row` 자체 구현** — 3 포맷 기능 패리티 달성. python-pptx 에는
  공식 add_row API 가 없지만, HWPX 에서 쓰던 lxml deepcopy 패턴을 `<a:tr>` 레벨에
  이식해 마지막 행 복제 + 텍스트 비움으로 구현. gridSpan/tcPr 등 서식 상속.

### Notes
- 제약: 마지막 행이 위 행의 rowSpan 에 걸려있으면 (vMerge="1" 또는 rowSpan>1
  셀 존재) `NotImplementedForFormat` — HWPX 와 동일 정책.

## [0.4.0] — 2026-04-16

### Changed (License)
- **`python-hwpx` 런타임 의존성 제거** → dev extras 로 이동 (Non-Commercial
  License 라 상용 배포 시 블로커였음). 이제 MIT / BSD / Apache-2.0 / LGPL-2.1
  허용형 OSS 만 런타임에 의존.
- **자체 `document_adapter.hwpx_core` 패키지 도입** — `zipfile` + `lxml` 로 HWPX
  ZIP+XML 직접 관리. python-hwpx 없이 동일 기능 제공.
  - `constants`: HWPX XML 네임스페이스
  - `package.HwpxPackage`: ZIP 컨테이너 + XML 파트 dirty tracking 저장
  - `grid.iter_grid`, `table_shape`: cellAddr + cellSpan 기반 logical grid
  - `paragraph`: run-level 편집 헬퍼 (포맷 보존)
- `HwpxAdapter` 전면 교체 — 공개 API 불변, 내부는 `hwpx_core` 만 사용
- `lxml>=5.0` 을 명시적 런타임 의존성으로 추가

### Added
- `scripts/hwpx_regression.py` — 4 스테이지 round-trip 회귀 harness
  (bytes copy / lxml rt / adapter rt / adapter edit) + `--baseline` / `--compare`
- `NOTICE` — xgen-doc2chunk (Apache-2.0) grid 파싱 로직 차용 고지

### Notes
- 실전 HWPX 10 fixture 전 스테이지 그린 (gov_large 865KB / 188 tables / 281 merges 포함)
- 한컴 Office HWP Viewer 수동 호환성 확인

## [0.3.0] — 2026-04-16

### Added
- `get_cell` 도구 — preview 40자 컷 없이 셀 전체 내용 조회. paragraphs, is_anchor,
  anchor, span, nested_table_indices 반환.
- `append_to_cell` 도구 — 기존 셀 텍스트 유지한 채 값 뒤에 덧붙임 (라벨 보존용).
- HWPX `append_row` 지원.
- DOCX / PPTX 병합 셀 감지 통일.

## 이전 버전

오늘(2026-04-16) 세션에서 v0.3.0 부터 v0.6.0 까지 연속 릴리스. 이전 버전 히스토리는
git log 참조.

[Unreleased]: https://github.com/PlateerLab/document-adapter/compare/v0.6.0...HEAD
[0.6.0]: https://github.com/PlateerLab/document-adapter/compare/v0.5.0...v0.6.0
[0.5.0]: https://github.com/PlateerLab/document-adapter/compare/v0.4.0...v0.5.0
[0.4.0]: https://github.com/PlateerLab/document-adapter/compare/v0.3.0...v0.4.0
[0.3.0]: https://github.com/PlateerLab/document-adapter/releases/tag/v0.3.0
