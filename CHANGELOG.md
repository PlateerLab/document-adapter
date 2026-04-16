# Changelog

All notable changes to this project will be documented in this file.

Format: [Keep a Changelog](https://keepachangelog.com/en/1.1.0/). Versioning:
[SemVer](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

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
