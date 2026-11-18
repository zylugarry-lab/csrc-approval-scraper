from __future__ import annotations

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from .formatter import apply_indent, apply_spacing
from .config import ParaSpec
from .xmlutil import (
    W_P,
    W_TBL,
    W_TC,
    W_TR,
    WP_ANCHOR,
    ancestor_is,
    paragraph_plain_text,
)

WP_POS_H = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}positionH"
WP_POS_V = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}positionV"
WP_OFF = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}posOffset"


def _direct_cells(tr):
    return [c for c in tr if c.tag == W_TC]


def _ensure_tcPr(tc):
    tcPr = tc.find(qn("w:tcPr"))
    if tcPr is None:
        tcPr = OxmlElement("w:tcPr")
        tc.insert(0, tcPr)
    return tcPr


def _set_vmerge(tc, val: str) -> None:
    tcPr = _ensure_tcPr(tc)
    vm = tcPr.find(qn("w:vMerge"))
    if vm is None:
        vm = OxmlElement("w:vMerge")
        tcPr.append(vm)
    if val == "continue":
        vm.set(qn("w:val"), "continue")
    else:
        vm.set(qn("w:val"), val)


def _set_valign(tc, val: str) -> None:
    tcPr = _ensure_tcPr(tc)
    va = tcPr.find(qn("w:vAlign"))
    if va is None:
        va = OxmlElement("w:vAlign")
        tcPr.append(va)
    va.set(qn("w:val"), val)


def _first_para(tc):
    for child in tc:
        if child.tag == W_P:
            return child
    return None


def _looks_like_formula_left(text: str) -> bool:
    t = "".join((text or "").split())
    if not t:
        return False
    return t.endswith("=") or t.endswith("＝") or "业绩报酬" in t


def is_piecewise_formula_table(tbl) -> bool:
    """嵌套小表 + 浮动大括号图 + 左侧「xxx=」：分段函数。"""
    if not ancestor_is(tbl, W_TC):
        return False
    if tbl.find(".//{0}".format(WP_ANCHOR)) is None:
        return False
    rows = tbl.findall(W_TR)
    if len(rows) < 2:
        return False
    cells0 = _direct_cells(rows[0])
    if len(cells0) < 2:
        return False
    left = _first_para(cells0[0])
    if left is None:
        return False
    return _looks_like_formula_left(paragraph_plain_text(left))


def _zero_anchor_offset(tbl) -> None:
    """大括号贴到右列左边，垂直从首行顶开始，避免漂在文字中间。"""
    for pos in tbl.iter(WP_POS_H):
        for off in pos:
            if off.tag == WP_OFF:
                off.text = "20000"
    for pos in tbl.iter(WP_POS_V):
        for off in pos:
            if off.tag == WP_OFF:
                off.text = "0"


def _zero_tbl_ind(tbl) -> None:
    tblPr = tbl.find(qn("w:tblPr"))
    if tblPr is None:
        return
    ind = tblPr.find(qn("w:tblInd"))
    if ind is not None:
        ind.set(qn("w:w"), "0")


def normalize_piecewise_formula_table(tbl) -> None:
    rows = tbl.findall(W_TR)
    cells0 = _direct_cells(rows[0])
    cells1 = _direct_cells(rows[1])
    left0, right0 = cells0[0], cells0[1]
    left1, right1 = cells1[0], cells1[1]

    _set_vmerge(left0, "restart")
    _set_vmerge(left1, "continue")
    _set_valign(left0, "center")
    _set_valign(left1, "center")
    _set_valign(right0, "top")
    _set_valign(right1, "top")

    spec = ParaSpec(
        alignment="left",
        first_line_chars=0.0,
        line_spacing_rule="auto",
        line_spacing=1.5,
        space_before_pt=0.0,
        space_after_pt=0.0,
    )
    size_pt = 9.0
    for p_el, left_chars in (
        (_first_para(left0), 0.0),
        (_first_para(right0), 2.0),
        (_first_para(right1), 2.0),
    ):
        if p_el is None:
            continue
        pPr = p_el.find(qn("w:pPr"))
        if pPr is None:
            pPr = OxmlElement("w:pPr")
            p_el.insert(0, pPr)
        apply_spacing(pPr, spec)
        apply_indent(pPr, size_pt, left_chars, 0.0, 0.0)

    _zero_anchor_offset(tbl)
    _zero_tbl_ind(tbl)


def normalize_piecewise_formulas(root) -> int:
    n = 0
    for tbl in list(root.iter(W_TBL)):
        if is_piecewise_formula_table(tbl):
            normalize_piecewise_formula_table(tbl)
            n += 1
    return n
