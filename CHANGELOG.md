# Changelog

All notable changes to this project will be documented in this file.

Format: [Keep a Changelog](https://keepachangelog.com/en/1.1.0/). Versioning:
[SemVer](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

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
