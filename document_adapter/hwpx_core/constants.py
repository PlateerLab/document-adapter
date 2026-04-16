"""HWPX XML 네임스페이스 상수."""
from __future__ import annotations

HP_NS = "http://www.hancom.co.kr/hwpml/2011/paragraph"
HC_NS = "http://www.hancom.co.kr/hwpml/2011/core"
HH_NS = "http://www.hancom.co.kr/hwpml/2011/head"
HS_NS = "http://www.hancom.co.kr/hwpml/2011/section"
OPF_NS = "http://www.idpf.org/2007/opf/"

HP_P = f"{{{HP_NS}}}p"
HP_RUN = f"{{{HP_NS}}}run"
HP_T = f"{{{HP_NS}}}t"
HP_TBL = f"{{{HP_NS}}}tbl"
HP_TR = f"{{{HP_NS}}}tr"
HP_TC = f"{{{HP_NS}}}tc"
HP_CELL_ADDR = f"{{{HP_NS}}}cellAddr"
HP_CELL_SPAN = f"{{{HP_NS}}}cellSpan"
HP_CELL_SZ = f"{{{HP_NS}}}cellSz"
HP_SUBLIST = f"{{{HP_NS}}}subList"
HS_SEC = f"{{{HS_NS}}}sec"

NAMESPACES = {
    "hp": HP_NS,
    "hc": HC_NS,
    "hh": HH_NS,
    "hs": HS_NS,
    "opf": OPF_NS,
}
