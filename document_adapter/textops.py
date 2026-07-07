"""Run 단위 텍스트 검색/치환 핵심 알고리즘 (포맷 독립, 순수 함수).

DOCX/HWPX 는 한 문단(paragraph)의 텍스트가 여러 run 으로 쪼개져 저장될 수 있다
(예: "홍길동" 이 "홍길" + "동" — 편집 이력·서식 경계·rsid 때문에 워드프로세서가
임의로 분할). 그래서 run 하나씩 보면 단어를 못 찾는다.

여기의 두 함수가 그 문제를 푼다:

- :func:`find_spans`   — 문단의 run 텍스트를 이어붙인(concat) 문자열에서
  쿼리의 등장 구간 [start, end) 목록을 구한다.
- :func:`splice_runs`  — concat 오프셋 기준 [start, end) 구간을 치환 텍스트로
  갈아끼우되, **매치에 걸친 run 만** 수정한다. 치환 텍스트는 매치가 시작되는
  run 에 들어가므로 그 run 의 서식(rPr/charPr)을 그대로 상속한다. 매치 밖
  run 은 손대지 않아 문단 내 혼합 서식이 보존된다.

두 함수 모두 XML 을 모른다 — 어댑터(DOCX/HWPX)가 run 텍스트 리스트를
읽어 넘기고, 반환된 리스트를 다시 run 에 써넣는 방식으로 재사용한다.
"""
from __future__ import annotations

import re

__all__ = ["find_spans", "splice_runs"]


def find_spans(
    text: str,
    query: str,
    *,
    whole_word: bool = False,
) -> list[tuple[int, int]]:
    """``text`` 에서 ``query`` 의 모든 (겹치지 않는) 등장 구간을 반환.

    Args:
        text: 검색 대상 (문단 run 텍스트의 concat).
        query: 찾을 리터럴 문자열 (regex 아님 — 이스케이프됨).
        whole_word: True 면 단어 경계 강제. 한글 조사/접미 방지용 —
            "홍길동님" 안의 "홍길동" 은 매치되지 않는다. (한글도 ``\\w`` 에
            포함되므로 lookaround 로 판정 가능.)

    Returns:
        [(start, end), ...] — 등장 순서대로. 없으면 빈 리스트.
    """
    if not query:
        return []
    pattern = re.escape(query)
    if whole_word:
        pattern = rf"(?<!\w){pattern}(?!\w)"
    return [(m.start(), m.end()) for m in re.finditer(pattern, text)]


def splice_runs(
    texts: list[str],
    start: int,
    end: int,
    repl: str,
) -> list[str]:
    """run 텍스트 리스트에서 concat 오프셋 [start, end) 를 ``repl`` 로 치환.

    반환 리스트는 입력과 **같은 길이** — i번째 원소가 i번째 run 의 새 텍스트다.
    (호출자는 값이 달라진 run 만 다시 써서 무관한 run 의 XML 재구성을 피할 것.)

    동작:
        - 매치가 시작되는 run: ``앞부분 + repl + (같은 run 에서 끝나면 뒷부분)``
        - 매치 중간 run: 전부 소비 → ``""``
        - 매치가 끝나는 run: 매치 이후 부분만 남김
        - 매치 밖 run: 불변

    빈 구간(start == end, 삽입)도 지원: 해당 오프셋을 포함하는 run 에 삽입하며,
    run 경계에 정확히 걸린 경우 **왼쪽 run 의 꼬리**에 붙는다 (왼쪽 서식 상속).

    Raises:
        ValueError: start/end 가 역전됐거나 전체 길이를 벗어난 경우,
            또는 텍스트를 붙일 run 이 하나도 없는 경우.
    """
    if start > end:
        raise ValueError(f"start({start}) > end({end})")
    total = sum(len(t) for t in texts)
    if start < 0 or end > total:
        raise ValueError(
            f"span [{start},{end}) out of range (total={total})"
        )

    out = list(texts)
    placed = False
    pos = 0
    for i, t in enumerate(texts):
        run_s, run_e = pos, pos + len(t)
        pos = run_e
        if run_e <= start or run_s >= end:
            continue  # 매치 밖 run — 불변
        local_s = max(start - run_s, 0)
        local_e = min(end - run_s, len(t))
        if not placed:
            # 첫 겹침 run — run 순서상 반드시 run_s <= start
            out[i] = t[:local_s] + repl + t[local_e:]
            placed = True
        else:
            out[i] = t[local_e:]

    if not placed:
        # start == end 가 run 경계(또는 문단 맨앞/맨뒤)에 정확히 걸린 삽입.
        # 오프셋을 포함하는 첫 run 의 해당 위치에 삽입 → 그 run 이 곧
        # "왼쪽에서 끝나는 run" 이므로 왼쪽 서식을 상속한다.
        pos = 0
        for i, t in enumerate(texts):
            run_s, run_e = pos, pos + len(t)
            pos = run_e
            if run_s <= start <= run_e:
                local = start - run_s
                out[i] = t[:local] + repl + t[local:]
                placed = True
                break
        if not placed:
            raise ValueError("no runs available to place text into")

    return out
