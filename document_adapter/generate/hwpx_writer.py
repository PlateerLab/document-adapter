"""블록 IR → HWPX(OWPML) 결정적 렌더러 (LLM 없음).

python-hwpx 같은 외부 의존 없이, 최소 유효 OWPML 패키지를 직접 조립한다.
구조는 한컴오피스 한글이 저장한 실제 .hwpx (KS X 6101) 를 기준으로 했다:

    mimetype                  "application/hwp+zip" (ZIP 첫 엔트리, STORED)
    version.xml               HCFVersion
    META-INF/container.xml    OCF rootfile → Contents/content.hpf
    META-INF/manifest.xml     빈 ODF manifest
    Contents/content.hpf      OPF package (metadata/manifest/spine)
    Contents/header.xml       hh:head — fontface/borderFill/charPr/paraPr/style
    Contents/section0.xml     hs:sec — secPr + 본문 문단/표
    settings.xml              캐럿 위치

렌더 원칙:
- charPr/paraPr 은 이 모듈이 정의한 고정 팔레트만 사용 (아래 _CHARPR_*).
- lineseg(레이아웃 캐시)는 쓰지 않는다 — 한글이 열 때 재계산한다.
- 생성 직후 HwpxAdapter.load() roundtrip 이 성립해야 한다
  (inspect_document → set_cell / replace_text 로 이어 편집).

단위: HWPUNIT (1/7200 inch). A4 세로 = 59528 x 84186, 본문폭(여백 제외) 42520.
"""
from __future__ import annotations

import io
import zipfile
from xml.sax.saxutils import escape

from .markdown_parser import Block, Span, parse_markdown

__all__ = ["hwpx_from_markdown"]

# ── 고정 팔레트: charPr id ──────────────────────────────────────────────
_CHARPR_NORMAL = 0
_CHARPR_BOLD = 1
_CHARPR_ITALIC = 2
_CHARPR_CODE = 3
_CHARPR_H1 = 4      # 16pt bold
_CHARPR_H2 = 5      # 14pt bold
_CHARPR_H3 = 6      # 12pt bold (레벨 3+ 공용)
_CHARPR_HEADER_CELL = 1  # 표 헤더행 = bold

# paraPr id: 0=본문(양쪽정렬) 1=헤딩(왼쪽) 2=목록(왼쪽 들여쓰기)
_PARAPR_BODY = 0
_PARAPR_HEADING = 1
_PARAPR_LIST = 2

_PAGE_W, _PAGE_H = 59528, 84186
_MARGIN_LR = 8504
_BODY_W = _PAGE_W - 2 * _MARGIN_LR  # 42520
_ROW_H = 1985

_HEAD_NS = (
    'xmlns:ha="http://www.hancom.co.kr/hwpml/2011/app" '
    'xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph" '
    'xmlns:hp10="http://www.hancom.co.kr/hwpml/2016/paragraph" '
    'xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" '
    'xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core" '
    'xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head" '
    'xmlns:hhs="http://www.hancom.co.kr/hwpml/2011/history" '
    'xmlns:hm="http://www.hancom.co.kr/hwpml/2011/master-page" '
    'xmlns:hpf="http://www.hancom.co.kr/schema/2011/hpf" '
    'xmlns:dc="http://purl.org/dc/elements/1.1/" '
    'xmlns:opf="http://www.idpf.org/2007/opf/" '
    'xmlns:ooxmlchart="http://www.hancom.co.kr/hwpml/2016/ooxmlchart" '
    'xmlns:epub="http://www.idpf.org/2007/ops" '
    'xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0"'
)

_XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>'

_FONT_LANGS = ("HANGUL", "LATIN", "HANJA", "JAPANESE", "OTHER", "SYMBOL", "USER")


def _header_xml(font: str) -> str:
    """Contents/header.xml — 고정 팔레트의 최소 refList."""
    fontfaces = "".join(
        f'<hh:fontface lang="{lang}" fontCnt="1">'
        f'<hh:font id="0" face="{escape(font)}" type="TTF" isEmbedded="0"/>'
        f"</hh:fontface>"
        for lang in _FONT_LANGS
    )

    def border_fill(fid: int, border_type: str) -> str:
        edge = f'<hh:{{side}}Border type="{border_type}" width="0.12 mm" color="#000000"/>'
        sides = "".join(
            edge.format(side=s) for s in ("left", "right", "top", "bottom")
        )
        return (
            f'<hh:borderFill id="{fid}" threeD="0" shadow="0" '
            f'centerLine="NONE" breakCellSeparateLine="0">'
            f'<hh:slash type="NONE" Crooked="0" isCounter="0"/>'
            f'<hh:backSlash type="NONE" Crooked="0" isCounter="0"/>'
            f"{sides}"
            f'<hh:diagonal type="SOLID" width="0.1 mm" color="#000000"/>'
            f"</hh:borderFill>"
        )

    border_fills = border_fill(1, "NONE") + border_fill(2, "SOLID")

    def char_pr(cid: int, height: int, *, bold: bool = False,
                italic: bool = False) -> str:
        refs = " ".join(f'{a}="0"' for a in
                        ("hangul", "latin", "hanja", "japanese", "other",
                         "symbol", "user"))
        vals100 = " ".join(
            f'{a}="100"' for a in
            ("hangul", "latin", "hanja", "japanese", "other", "symbol", "user")
        )
        vals0 = " ".join(
            f'{a}="0"' for a in
            ("hangul", "latin", "hanja", "japanese", "other", "symbol", "user")
        )
        marks = ("<hh:bold/>" if bold else "") + ("<hh:italic/>" if italic else "")
        return (
            f'<hh:charPr id="{cid}" height="{height}" textColor="#000000" '
            f'shadeColor="none" useFontSpace="0" useKerning="0" '
            f'symMark="NONE" borderFillIDRef="1">'
            f"<hh:fontRef {refs}/><hh:ratio {vals100}/><hh:spacing {vals0}/>"
            f"<hh:relSz {vals100}/><hh:offset {vals0}/>{marks}"
            f"</hh:charPr>"
        )

    char_prs = (
        char_pr(_CHARPR_NORMAL, 1000)
        + char_pr(_CHARPR_BOLD, 1000, bold=True)
        + char_pr(_CHARPR_ITALIC, 1000, italic=True)
        + char_pr(_CHARPR_CODE, 900)
        + char_pr(_CHARPR_H1, 1600, bold=True)
        + char_pr(_CHARPR_H2, 1400, bold=True)
        + char_pr(_CHARPR_H3, 1200, bold=True)
    )

    def para_pr(pid: int, align: str, indent: int = 0) -> str:
        return (
            f'<hh:paraPr id="{pid}" tabPrIDRef="0" condense="0" '
            f'fontLineHeight="0" snapToGrid="1" suppressLineNumbers="0" '
            f'checked="0">'
            f'<hh:align horizontal="{align}" vertical="BASELINE"/>'
            f'<hh:heading type="NONE" idRef="0" level="0"/>'
            f'<hh:breakSetting breakLatinWord="KEEP_WORD" '
            f'breakNonLatinWord="BREAK_WORD" widowOrphan="0" keepWithNext="0" '
            f'keepLines="0" pageBreakBefore="0" lineWrap="BREAK"/>'
            f"<hh:autoSpacing eAsianEng=\"0\" eAsianNum=\"0\"/>"
            f'<hh:margin><hc:intent value="0" unit="HWPUNIT"/>'
            f'<hc:left value="{indent}" unit="HWPUNIT"/>'
            f'<hc:right value="0" unit="HWPUNIT"/>'
            f'<hc:prev value="0" unit="HWPUNIT"/>'
            f'<hc:next value="0" unit="HWPUNIT"/></hh:margin>'
            f'<hh:lineSpacing type="PERCENT" value="160" unit="HWPUNIT"/>'
            f'<hh:border borderFillIDRef="1" offsetLeft="0" offsetRight="0" '
            f'offsetTop="0" offsetBottom="0" connect="0" '
            f'ignoreMargin="0"/>'
            f"</hh:paraPr>"
        )

    para_prs = (
        para_pr(_PARAPR_BODY, "JUSTIFY")
        + para_pr(_PARAPR_HEADING, "LEFT")
        + para_pr(_PARAPR_LIST, "LEFT", indent=1600)
    )

    return (
        f"{_XML_DECL}<hh:head {_HEAD_NS} version=\"1.2\" secCnt=\"1\">"
        f'<hh:beginNum page="1" footnote="1" endnote="1" pic="1" tbl="1" '
        f'equation="1"/>'
        f"<hh:refList>"
        f'<hh:fontfaces itemCnt="7">{fontfaces}</hh:fontfaces>'
        f'<hh:borderFills itemCnt="2">{border_fills}</hh:borderFills>'
        f'<hh:charProperties itemCnt="7">{char_prs}</hh:charProperties>'
        f'<hh:tabProperties itemCnt="1">'
        f'<hh:tabPr id="0" autoTabLeft="0" autoTabRight="0"/>'
        f"</hh:tabProperties>"
        f'<hh:paraProperties itemCnt="3">{para_prs}</hh:paraProperties>'
        f'<hh:styles itemCnt="1">'
        f'<hh:style id="0" type="PARA" name="바탕글" engName="Normal" '
        f'paraPrIDRef="0" charPrIDRef="0" nextStyleIDRef="0" langID="1042" '
        f'lockForm="0"/>'
        f"</hh:styles>"
        f"</hh:refList></hh:head>"
    )


def _char_pr_for(span: Span, default: int = _CHARPR_NORMAL) -> int:
    if span.code:
        return _CHARPR_CODE
    if span.bold:
        return _CHARPR_BOLD
    if span.italic:
        return _CHARPR_ITALIC
    return default


def _runs_xml(spans: tuple[Span, ...], default_char_pr: int) -> str:
    """Span 튜플 → <hp:run> 나열. 스팬이 없으면 빈 run 하나.

    default_char_pr 가 NORMAL 이면 스팬별 인라인 서식(bold/italic/code)을,
    그 외(헤딩/코드/표 헤더 등 문단 전체 서식)면 default 를 일괄 적용한다.
    """
    if not spans:
        return f'<hp:run charPrIDRef="{default_char_pr}"><hp:t/></hp:run>'
    parts = []
    for span in spans:
        cid = (
            _char_pr_for(span)
            if default_char_pr == _CHARPR_NORMAL
            else default_char_pr
        )
        parts.append(
            f'<hp:run charPrIDRef="{cid}">'
            f"<hp:t>{escape(span.text)}</hp:t></hp:run>"
        )
    return "".join(parts)


def _para_xml(spans: tuple[Span, ...], *, para_pr: int = _PARAPR_BODY,
              char_pr: int = _CHARPR_NORMAL, inner: str = "") -> str:
    """<hp:p> 한 개. inner 는 run 앞에 붙는 추가 XML (표 run 등)."""
    return (
        f'<hp:p paraPrIDRef="{para_pr}" styleIDRef="0" pageBreak="0" '
        f'columnBreak="0" merged="0">'
        f"{inner}{_runs_xml(spans, char_pr)}</hp:p>"
    )


def _cell_xml(spans: tuple[Span, ...], row: int, col: int, width: int,
              char_pr: int) -> str:
    return (
        f'<hp:tc name="" header="0" hasMargin="0" protect="0" editable="0" '
        f'dirty="0" borderFillIDRef="2">'
        f'<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" '
        f'vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0" '
        f'textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">'
        f"{_para_xml(spans, char_pr=char_pr)}"
        f"</hp:subList>"
        f'<hp:cellAddr colAddr="{col}" rowAddr="{row}"/>'
        f'<hp:cellSpan colSpan="1" rowSpan="1"/>'
        f'<hp:cellSz width="{width}" height="{_ROW_H}"/>'
        f'<hp:cellMargin left="510" right="510" top="141" bottom="141"/>'
        f"</hp:tc>"
    )


def _table_xml(block: Block, table_seq: int) -> str:
    """표 블록 → 문단 run 안에 인라인(treatAsChar) 배치되는 <hp:tbl>."""
    n_rows = len(block.rows)
    n_cols = len(block.rows[0]) if block.rows else 0
    if n_rows == 0 or n_cols == 0:
        return ""
    col_w = _BODY_W // n_cols
    rows_xml = []
    for r, row in enumerate(block.rows):
        cells = "".join(
            _cell_xml(
                row[c],
                r,
                c,
                col_w,
                _CHARPR_HEADER_CELL if r == 0 else _CHARPR_NORMAL,
            )
            for c in range(n_cols)
        )
        rows_xml.append(f"<hp:tr>{cells}</hp:tr>")
    tbl = (
        f'<hp:tbl id="{1859000000 + table_seq}" zOrder="{table_seq}" '
        f'numberingType="TABLE" textWrap="TOP_AND_BOTTOM" '
        f'textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" '
        f'pageBreak="CELL" repeatHeader="1" rowCnt="{n_rows}" '
        f'colCnt="{n_cols}" cellSpacing="0" borderFillIDRef="2" '
        f'noAdjust="0">'
        f'<hp:sz width="{_BODY_W}" widthRelTo="ABSOLUTE" '
        f'height="{_ROW_H * n_rows}" heightRelTo="ABSOLUTE" protect="0"/>'
        f'<hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1" '
        f'allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PARA" '
        f'horzRelTo="COLUMN" vertAlign="TOP" horzAlign="LEFT" '
        f'vertOffset="0" horzOffset="0"/>'
        f'<hp:outMargin left="283" right="283" top="283" bottom="283"/>'
        f'<hp:inMargin left="510" right="510" top="141" bottom="141"/>'
        f"{''.join(rows_xml)}</hp:tbl>"
    )
    # 표는 문단의 run 안에 위치한다 (treatAsChar=1 인라인 개체).
    return (
        f'<hp:p paraPrIDRef="{_PARAPR_BODY}" styleIDRef="0" pageBreak="0" '
        f'columnBreak="0" merged="0">'
        f'<hp:run charPrIDRef="{_CHARPR_NORMAL}">{tbl}<hp:t/></hp:run></hp:p>'
    )


_SEC_PR = (
    '<hp:secPr id="" textDirection="HORIZONTAL" spaceColumns="1134" '
    'tabStop="8000" outlineShapeIDRef="0" memoShapeIDRef="0" '
    'textVerticalWidthHead="0" masterPageCnt="0">'
    '<hp:grid lineGrid="0" charGrid="0" wonggojiFormat="0"/>'
    '<hp:startNum pageStartsOn="BOTH" page="0" pic="0" tbl="0" equation="0"/>'
    '<hp:visibility hideFirstHeader="0" hideFirstFooter="0" '
    'hideFirstMasterPage="0" border="SHOW_ALL" fill="SHOW_ALL" '
    'hideFirstPageNum="0" hideFirstEmptyLine="0" showLineNumber="0"/>'
    '<hp:lineNumberShape restartType="0" countBy="0" distance="0" '
    'startNumber="0"/>'
    f'<hp:pagePr landscape="WIDELY" width="{_PAGE_W}" height="{_PAGE_H}" '
    'gutterType="LEFT_ONLY">'
    f'<hp:margin header="4252" footer="4252" gutter="0" left="{_MARGIN_LR}" '
    f'right="{_MARGIN_LR}" top="5668" bottom="4252"/></hp:pagePr>'
    "</hp:secPr>"
)


def _section_xml(blocks: list[Block]) -> str:
    body: list[str] = []
    numbered_count = 0
    table_seq = 0
    for block in blocks:
        if block.kind == "heading":
            char_pr = {1: _CHARPR_H1, 2: _CHARPR_H2}.get(block.level, _CHARPR_H3)
            body.append(
                _para_xml(block.spans, para_pr=_PARAPR_HEADING, char_pr=char_pr)
            )
            numbered_count = 0
        elif block.kind == "bullet":
            spans = (Span("• "),) + block.spans
            body.append(_para_xml(spans, para_pr=_PARAPR_LIST))
            numbered_count = 0
        elif block.kind == "numbered":
            numbered_count += 1
            spans = (Span(f"{numbered_count}. "),) + block.spans
            body.append(_para_xml(spans, para_pr=_PARAPR_LIST))
        elif block.kind == "quote":
            spans = tuple(
                Span(s.text, bold=s.bold, italic=True, code=s.code)
                for s in block.spans
            )
            body.append(_para_xml(spans, para_pr=_PARAPR_LIST))
            numbered_count = 0
        elif block.kind == "hr":
            body.append(_para_xml((Span("─" * 40),)))
            numbered_count = 0
        elif block.kind == "code":
            for line in block.lines:
                body.append(
                    _para_xml((Span(line, code=True),), char_pr=_CHARPR_CODE)
                )
            numbered_count = 0
        elif block.kind == "table":
            table_seq += 1
            body.append(_table_xml(block, table_seq))
            numbered_count = 0
        else:  # paragraph (+ 안전망)
            body.append(_para_xml(block.spans))
            numbered_count = 0

    # 첫 문단의 첫 run 에 secPr — 한글이 요구하는 섹션 정의 위치.
    first = (
        f'<hp:p paraPrIDRef="{_PARAPR_BODY}" styleIDRef="0" pageBreak="0" '
        f'columnBreak="0" merged="0">'
        f'<hp:run charPrIDRef="{_CHARPR_NORMAL}">{_SEC_PR}<hp:t/></hp:run>'
        f"</hp:p>"
    )
    return (
        f"{_XML_DECL}<hs:sec {_HEAD_NS}>{first}{''.join(body)}</hs:sec>"
    )


_VERSION_XML = (
    f"{_XML_DECL}<hv:HCFVersion "
    'xmlns:hv="http://www.hancom.co.kr/hwpml/2011/version" '
    'tagetApplication="WORDPROCESSOR" major="5" minor="1" micro="0" '
    'buildNumber="1" os="1" xmlVersion="1.2" '
    'application="document-adapter" appVersion="0.16.0"/>'
)

_CONTAINER_XML = (
    f"{_XML_DECL}<ocf:container "
    'xmlns:ocf="urn:oasis:names:tc:opendocument:xmlns:container" '
    'xmlns:hpf="http://www.hancom.co.kr/schema/2011/hpf">'
    "<ocf:rootfiles>"
    '<ocf:rootfile full-path="Contents/content.hpf" '
    'media-type="application/hwpml-package+xml"/>'
    "</ocf:rootfiles></ocf:container>"
)

_MANIFEST_XML = (
    f"{_XML_DECL}<odf:manifest "
    'xmlns:odf="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"/>'
)

_SETTINGS_XML = (
    f"{_XML_DECL}<ha:HWPApplicationSetting "
    'xmlns:ha="http://www.hancom.co.kr/hwpml/2011/app" '
    'xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0">'
    '<ha:CaretPosition listIDRef="0" paraIDRef="0" pos="0"/>'
    "</ha:HWPApplicationSetting>"
)


def _content_hpf(title: str) -> str:
    return (
        f"{_XML_DECL}<opf:package {_HEAD_NS} version=\"\" "
        'unique-identifier="" id="">'
        f"<opf:metadata><opf:title>{escape(title)}</opf:title>"
        "<opf:language>ko</opf:language>"
        '<opf:meta name="creator" content="text">document-adapter</opf:meta>'
        "</opf:metadata>"
        "<opf:manifest>"
        '<opf:item id="header" href="Contents/header.xml" '
        'media-type="application/xml"/>'
        '<opf:item id="section0" href="Contents/section0.xml" '
        'media-type="application/xml"/>'
        '<opf:item id="settings" href="settings.xml" '
        'media-type="application/xml"/>'
        "</opf:manifest>"
        "<opf:spine><opf:itemref idref="
        '"header"/><opf:itemref idref="section0" linear="yes"/></opf:spine>'
        "</opf:package>"
    )


def hwpx_from_markdown(
    markdown: str, *, lang: str | None = "ko", base_font: str | None = None
) -> bytes:
    """markdown 서브셋 → .hwpx bytes.

    Raises:
        ValueError: 본문이 비어 있을 때 (이중어 메시지 — 호출 레이어의
            재시도 계약이 이 메시지를 LLM 리마인더로 사용한다).
    """
    if not markdown or not markdown.strip():
        raise ValueError(
            "empty document body — provide markdown content. "
            "문서 본문이 비어 있습니다 — markdown 내용을 작성하세요."
        )

    blocks = parse_markdown(markdown)
    if not blocks:
        raise ValueError(
            "markdown produced no renderable blocks. "
            "markdown 에서 렌더 가능한 블록을 찾지 못했습니다."
        )

    font = base_font or ("맑은 고딕" if (lang or "ko").startswith("ko") else "Arial")

    # 제목(첫 헤딩)을 문서 메타데이터 title 로.
    title = "document"
    for block in blocks:
        if block.kind == "heading" and block.spans:
            title = "".join(s.text for s in block.spans)
            break

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        # mimetype 은 반드시 첫 엔트리 + 무압축 (OCF 규약)
        info = zipfile.ZipInfo("mimetype")
        info.compress_type = zipfile.ZIP_STORED
        zf.writestr(info, "application/hwp+zip")
        zf.writestr("version.xml", _VERSION_XML)
        zf.writestr("META-INF/container.xml", _CONTAINER_XML)
        zf.writestr("META-INF/manifest.xml", _MANIFEST_XML)
        zf.writestr("Contents/content.hpf", _content_hpf(title))
        zf.writestr("Contents/header.xml", _header_xml(font))
        zf.writestr("Contents/section0.xml", _section_xml(blocks))
        zf.writestr("settings.xml", _SETTINGS_XML)
    return buf.getvalue()
