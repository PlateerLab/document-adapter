"""공개 평가 시나리오 — 합성 fixture + golden 기대 최종상태.

scripts/ollama_scenarios.py 의 비공개-fixture 시나리오를 공개 합성본으로
재구성한 것. 각 시나리오는 "지시 → 끝난 문서의 기대 셀 값"을 명시하므로,
어떤 백엔드로 돌리든 결과를 결정적으로 채점할 수 있다.
"""
from __future__ import annotations

from document_adapter.eval import fixtures as fx
from document_adapter.eval.harness import FieldExpectation as FE
from document_adapter.eval.harness import Scenario

PUBLIC_SCENARIOS: list[Scenario] = [
    Scenario(
        name="A hwpx blank auto",
        fmt="hwpx",
        build_fixture=fx.hwpx_blank_form,
        instruction=(
            "이 양식을 다음 정보로 채워줘: "
            "접수번호=2026-0001, 성명=홍길동, 주소=서울시 강남구"
        ),
        expectations=[
            FE(0, 0, 1, "2026-0001"),
            FE(0, 1, 1, "홍길동"),
            FE(0, 2, 1, "서울시 강남구"),
        ],
        protected_cells=[(0, 0, 0), (0, 1, 0), (0, 2, 0)],  # 라벨 보존
    ),
    Scenario(
        name="B docx blank auto",
        fmt="docx",
        build_fixture=fx.docx_blank_form,
        instruction="이 직원 양식을 채워줘: 성명=김철수, 부서=개발팀, 입사일=2026-04-17",
        expectations=[
            FE(0, 0, 1, "김철수"),
            FE(0, 1, 1, "개발팀"),
            FE(0, 2, 1, "2026-04-17"),
        ],
        protected_cells=[(0, 0, 0), (0, 1, 0), (0, 2, 0)],
    ),
    Scenario(
        name="C hwpx dup-section dot-path",
        fmt="hwpx",
        build_fixture=fx.hwpx_dup_section_form,
        instruction=(
            "피해자 정보 섹션의 금액=1,000,000, "
            "지급정지요청계좌 섹션의 금액=2,000,000 으로 채워줘. "
            "두 섹션에 같은 '금액' 라벨이 있으니 구분이 필요해."
        ),
        expectations=[
            FE(0, 1, 1, "1,000,000"),
            FE(0, 3, 1, "2,000,000"),
        ],
        protected_cells=[(0, 0, 0), (0, 2, 0)],  # 섹션 헤더 라벨 보존
    ),
    Scenario(
        name="D pptx filled-sample direction-right",
        fmt="pptx",
        build_fixture=fx.pptx_filled_sample_form,
        instruction=(
            "이 보고서 양식의 예시값을 다음으로 교체해줘: "
            "보고일자=2026-04-17, 작성자=홍길동, 담당부서=개발팀. "
            "기존 값은 예시일 뿐이라 덮어써야 해."
        ),
        expectations=[
            FE(0, 0, 1, "2026-04-17"),
            FE(0, 1, 1, "홍길동"),
            FE(0, 2, 1, "개발팀"),
        ],
        protected_cells=[(0, 0, 0), (0, 1, 0), (0, 2, 0)],
    ),
]


# ---- HARD: 실제 한국 공공 행정서식의 까다로운 구조를 재현한 시나리오 -------
# (헤더행 라벨 / 와이드 그리드 / 병합 섹션헤더 / 3섹션 중복라벨)
HARD_SCENARIOS: list[Scenario] = [
    Scenario(
        name="H1 hwpx header-below",
        fmt="hwpx",
        build_fixture=fx.hwpx_header_below_form,
        instruction=(
            "이 명부 표를 채워줘. 각 항목(성명/부서/직급)의 값은 "
            "라벨 바로 아래 칸에 들어가야 해: 성명=홍길동, 부서=개발팀, 직급=책임"
        ),
        expectations=[
            FE(0, 1, 0, "홍길동"),
            FE(0, 1, 1, "개발팀"),
            FE(0, 1, 2, "책임"),
        ],
        protected_cells=[(0, 0, 0), (0, 0, 1), (0, 0, 2)],  # 헤더행 라벨 보존
    ),
    Scenario(
        name="H2 hwpx wide-grid",
        fmt="hwpx",
        build_fixture=fx.hwpx_wide_grid_form,
        instruction=(
            "인적사항 표를 채워줘: 성명=김철수, 생년월일=1990-01-01, "
            "주소=서울시 강남구, 연락처=010-1111-2222"
        ),
        expectations=[
            FE(0, 0, 1, "김철수"),
            FE(0, 0, 3, "1990-01-01"),
            FE(0, 1, 1, "서울시 강남구"),
            FE(0, 1, 3, "010-1111-2222"),
        ],
        protected_cells=[(0, 0, 0), (0, 0, 2), (0, 1, 0), (0, 1, 2)],
    ),
    Scenario(
        name="H3 hwpx merged-section-header",
        fmt="hwpx",
        build_fixture=fx.hwpx_merged_section_form,
        instruction=(
            "신청인 정보를 채워줘: 성명=홍길동, 연락처=010-1234-5678. "
            "맨 위 '신청인 정보'는 섹션 제목이라 건드리지 마."
        ),
        expectations=[
            FE(0, 1, 1, "홍길동"),
            FE(0, 2, 1, "010-1234-5678"),
        ],
        protected_cells=[(0, 0, 0), (0, 1, 0), (0, 2, 0)],
    ),
    Scenario(
        name="H4 hwpx three-section dot-path",
        fmt="hwpx",
        build_fixture=fx.hwpx_three_section_form,
        instruction=(
            "신청인/대리인/보증인 세 섹션의 성명을 각각 채워줘: "
            "신청인 성명=홍길동, 대리인 성명=김대리, 보증인 성명=이보증. "
            "세 섹션에 같은 '성명' 라벨이 있으니 섹션별로 구분해야 해."
        ),
        expectations=[
            FE(0, 1, 1, "홍길동"),   # 신청인 성명
            FE(0, 4, 1, "김대리"),   # 대리인 성명
            FE(0, 7, 1, "이보증"),   # 보증인 성명
        ],
        protected_cells=[(0, 0, 0), (0, 3, 0), (0, 6, 0)],  # 섹션 헤더 보존
    ),
]

ALL_SCENARIOS: list[Scenario] = PUBLIC_SCENARIOS + HARD_SCENARIOS
