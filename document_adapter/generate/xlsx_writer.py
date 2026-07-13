"""sheet spec(dict) → XLSX 결정적 렌더러 (LLM 없음).

edit2docs v0.9 `documents/xlsx_engine.xlsx_from_spec`(Apache-2.0)의 렌더
전략을 dict 입력으로 이식 — NOTICE 참조. YAML 의존을 추가하지 않기 위해
스펙은 파싱 완료된 dict 를 받는다 (JSON 파싱은 도구 레이어 몫).

sheet spec 스키마 (LLM 도구 설명과 1:1):
    sheets: [
      { "name": "매출 요약",             # 필수, 시트 탭 이름 (31자 절단)
        "headers": ["분기", "매출(억원)"],  # 필수
        "rows": [["1분기", 120], ...],     # 필수. 숫자는 숫자 타입으로
        "widths": [10, 14],                # 선택, 문자 단위 열 폭
        "number_formats": {"B": "#,##0"} } # 선택, 열문자 → Excel 서식
    ]

렌더 규칙: 헤더행 굵게+회색 채움+테두리, 데이터행 테두리, 헤더 틀고정,
`=` 로 시작하는 문자열은 수식으로 통과.

모든 검증 실패는 이중어 ValueError — 호출 레이어의 1회 재시도 계약이
이 메시지를 LLM 리마인더로 사용한다.
"""
from __future__ import annotations

import io
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

__all__ = ["xlsx_from_sheets"]

_HEADER_FILL = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
_THIN = Side(style="thin", color="BFBFBF")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)

_MIN_WIDTH = 8
_MAX_WIDTH = 60


def _err(en: str, ko: str) -> ValueError:
    return ValueError(f"{en} {ko}")


def xlsx_from_sheets(sheets: list[dict[str, Any]]) -> bytes:
    """sheet spec → 스타일 잡힌 .xlsx bytes. 구조 오류는 ValueError."""
    if not isinstance(sheets, list) or not sheets:
        raise _err(
            "`sheets` must be a non-empty list.",
            "`sheets`는 비어있지 않은 리스트여야 합니다.",
        )

    wb = Workbook()
    wb.remove(wb.active)
    seen_titles: set[str] = set()

    for idx, sheet in enumerate(sheets):
        if not isinstance(sheet, dict):
            raise _err(
                f"sheets[{idx}] must be a mapping.",
                f"sheets[{idx}] 항목은 객체(dict)여야 합니다.",
            )
        name = str(sheet.get("name") or "").strip()
        if not name:
            raise _err(
                f"sheets[{idx}]: every sheet needs a `name`.",
                f"sheets[{idx}]: 모든 시트에는 `name`이 필요합니다.",
            )
        title = name[:31]
        if title in seen_titles:
            raise _err(
                f"duplicate sheet name {title!r} (after 31-char truncation).",
                f"시트 이름 {title!r} 이 중복됩니다 (31자 절단 후).",
            )
        seen_titles.add(title)

        headers = sheet.get("headers")
        rows = sheet.get("rows")
        if not isinstance(headers, list) or not headers:
            raise _err(
                f"sheet {title!r}: `headers` must be a non-empty list.",
                f"시트 {title!r}: `headers`는 비어있지 않은 리스트여야 합니다.",
            )
        if not isinstance(rows, list):
            raise _err(
                f"sheet {title!r}: `rows` must be a list.",
                f"시트 {title!r}: `rows`는 리스트여야 합니다.",
            )

        ws = wb.create_sheet(title=title)
        n_cols = len(headers)

        # 헤더행
        for c, header in enumerate(headers, start=1):
            cell = ws.cell(row=1, column=c, value=str(header))
            cell.font = Font(bold=True)
            cell.fill = _HEADER_FILL
            cell.border = _BORDER

        # 데이터행
        for r_idx, row in enumerate(rows):
            if not isinstance(row, list):
                raise _err(
                    f"sheet {title!r}: rows[{r_idx}] must be a list.",
                    f"시트 {title!r}: rows[{r_idx}] 항목은 리스트여야 합니다.",
                )
            if len(row) > n_cols:
                raise _err(
                    f"sheet {title!r}: rows[{r_idx}] has {len(row)} cells "
                    f"but only {n_cols} headers.",
                    f"시트 {title!r}: rows[{r_idx}] 셀 수({len(row)})가 "
                    f"헤더 수({n_cols})를 초과합니다.",
                )
            for c in range(n_cols):
                value = row[c] if c < len(row) else ""
                cell = ws.cell(row=r_idx + 2, column=c + 1, value=value)
                cell.border = _BORDER

        # 열 폭: widths 지정 우선, 없으면 내용 기반 자동
        widths = sheet.get("widths")
        for c in range(1, n_cols + 1):
            letter = get_column_letter(c)
            if isinstance(widths, list) and c - 1 < len(widths):
                try:
                    ws.column_dimensions[letter].width = max(
                        _MIN_WIDTH, min(_MAX_WIDTH, int(widths[c - 1]))
                    )
                    continue
                except (TypeError, ValueError):
                    pass
            longest = max(
                [len(str(headers[c - 1]))]
                + [len(str(row[c - 1])) for row in rows if c - 1 < len(row)],
                default=_MIN_WIDTH,
            )
            # CJK 는 2칸 폭이지만 근사만 — 정확한 폭은 편집 도구 몫
            ws.column_dimensions[letter].width = max(
                _MIN_WIDTH, min(_MAX_WIDTH, longest + 2)
            )

        # 숫자 서식 (열 문자 기준, 데이터행에만)
        number_formats = sheet.get("number_formats")
        if isinstance(number_formats, dict):
            for letter, fmt in number_formats.items():
                col_letter = str(letter).strip().upper()
                for r in range(2, len(rows) + 2):
                    ws[f"{col_letter}{r}"].number_format = str(fmt)

        ws.freeze_panes = "A2"

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
