"""textops.find_spans / splice_runs 단위 테스트 (XML 무관 순수 로직)."""
from __future__ import annotations

import pytest

from document_adapter.textops import find_spans, splice_runs


# -------- find_spans --------

def test_find_spans_basic():
    assert find_spans("가나다라 가나", "가나") == [(0, 2), (5, 7)]


def test_find_spans_none():
    assert find_spans("가나다", "없음") == []


def test_find_spans_empty_query():
    assert find_spans("가나다", "") == []


def test_find_spans_regex_metachar_escaped():
    # 쿼리는 리터럴 — regex 메타문자가 있어도 그대로 매치
    assert find_spans("금액(원) 표기", "금액(원)") == [(0, 5)]


def test_find_spans_whole_word_korean_suffix():
    text = "홍길동님께서 홍길동 씨와 김홍길동을 만났다"
    # 부분 일치 허용: 3건
    assert len(find_spans(text, "홍길동")) == 3
    # whole_word: "홍길동님"의 홍길동, "김홍길동"의 홍길동 배제 → 1건
    spans = find_spans(text, "홍길동", whole_word=True)
    assert len(spans) == 1
    s, e = spans[0]
    assert text[s:e] == "홍길동"
    assert text[s - 1] == " " and text[e] == " "


def test_find_spans_whole_word_punctuation_boundary():
    # 문장부호/괄호는 단어 경계로 인정
    assert find_spans("(홍길동)", "홍길동", whole_word=True) == [(1, 4)]
    assert find_spans("성명:홍길동.", "홍길동", whole_word=True) == [(3, 6)]


# -------- splice_runs: 치환 --------

def test_splice_single_run_middle():
    out = splice_runs(["성명: 홍길동 님"], 4, 7, "유지수")
    assert out == ["성명: 유지수 님"]


def test_splice_whole_single_run():
    assert splice_runs(["홍길동"], 0, 3, "유지수") == ["유지수"]


def test_splice_span_two_runs():
    # "홍길" + "동님께" → 홍길동(0..3) 치환
    out = splice_runs(["홍길", "동님께"], 0, 3, "유지수")
    assert out == ["유지수", "님께"]
    assert "".join(out) == "유지수님께"


def test_splice_span_three_runs_middle_consumed():
    # "a홍" + "길" + "동b" → 홍길동(1..4)
    out = splice_runs(["a홍", "길", "동b"], 1, 4, "유지수")
    assert out == ["a유지수", "", "b"]
    assert "".join(out) == "a유지수b"


def test_splice_replacement_shorter_and_longer():
    assert "".join(splice_runs(["가나다라마"], 1, 4, "X")) == "가X마"
    assert "".join(splice_runs(["가나"], 0, 1, "아주긴치환문")) == "아주긴치환문나"


def test_splice_at_paragraph_start_and_end():
    assert splice_runs(["홍길동 귀하"], 0, 3, "유지수") == ["유지수 귀하"]
    assert splice_runs(["담당: 홍길동"], 4, 7, "유지수") == ["담당: 유지수"]


def test_splice_untouched_runs_preserved():
    # 매치 밖 run 은 객체/값 완전 불변
    texts = ["앞부분", "홍길동", "뒷부분"]
    out = splice_runs(texts, 3, 6, "유지수")
    assert out[0] == "앞부분" and out[2] == "뒷부분"
    assert out[1] == "유지수"


# -------- splice_runs: 삽입 (빈 구간) --------

def test_splice_insert_middle_of_run():
    out = splice_runs(["성명 확인"], 2, 2, ":")
    assert out == ["성명: 확인"]


def test_splice_insert_at_run_boundary_attaches_left():
    # 경계 삽입은 왼쪽 run 꼬리에 → 왼쪽(앵커) 서식 상속
    out = splice_runs(["성명", " 확인"], 2, 2, " 유지수")
    assert out == ["성명 유지수", " 확인"]


def test_splice_insert_at_offset_zero():
    out = splice_runs(["본문"], 0, 0, "머리")
    assert out == ["머리본문"]


def test_splice_insert_at_paragraph_end():
    # "성명: "(4자) + "홍길동"(3자) = 총 7자 — 맨 끝(7,7)에 삽입
    out = splice_runs(["성명: ", "홍길동"], 7, 7, " (인)")
    assert out == ["성명: ", "홍길동 (인)"]


# -------- splice_runs: 오류 --------

def test_splice_errors():
    with pytest.raises(ValueError):
        splice_runs(["가나"], 2, 1, "x")          # start > end
    with pytest.raises(ValueError):
        splice_runs(["가나"], 0, 5, "x")          # 범위 초과
    with pytest.raises(ValueError):
        splice_runs(["가나"], -1, 1, "x")         # 음수
    with pytest.raises(ValueError):
        splice_runs([], 0, 0, "x")               # run 없음


# -------- 조합: find + splice (역순 다중 치환) --------

def test_find_then_splice_multiple_reversed():
    texts = ["홍길동은 홍길동을 아낀다"]
    text = "".join(texts)
    spans = find_spans(text, "홍길동")
    assert len(spans) == 2
    for s, e in reversed(spans):
        texts = splice_runs(texts, s, e, "유지수")
    assert "".join(texts) == "유지수은 유지수을 아낀다"
