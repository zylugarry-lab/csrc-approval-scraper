from __future__ import annotations

from docx.oxml.ns import qn
from docx.oxml import OxmlElement

W_P = qn("w:p")
W_R = qn("w:r")
W_T = qn("w:t")
W_PPR = qn("w:pPr")
W_RPR = qn("w:rPr")
W_TBL = qn("w:tbl")
W_TR = qn("w:tr")
W_TC = qn("w:tc")
W_HYPERLINK = qn("w:hyperlink")
W_SDT = qn("w:sdt")
W_TXBX = qn("w:txbxContent")
W_DRAWING = qn("w:drawing")
W_PICT = qn("w:pict")
W_OBJECT = qn("w:object")
W_BR = qn("w:br")
W_TAB = qn("w:tab")
W_SECTPR = qn("w:sectPr")
WP_ANCHOR = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}anchor"
M_OMATH = "{http://schemas.openxmlformats.org/officeDocument/2006/math}oMath"
M_OMATH_PARA = "{http://schemas.openxmlformats.org/officeDocument/2006/math}oMathPara"


def get_or_create(parent, tag: str):
    el = parent.find(qn(tag))
    if el is None:
        el = OxmlElement(tag)
        parent.append(el)
    return el


def drop_children(parent, tag: str) -> int:
    count = 0
    for child in parent.findall(qn(tag)):
        parent.remove(child)
        count += 1
    return count


def drop_attrs(el, names) -> None:
    for name in names:
        key = qn(name)
        if key in el.attrib:
            del el.attrib[key]


def ensure_pPr(p_el):
    pPr = p_el.find(W_PPR)
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        p_el.insert(0, pPr)
    return pPr


def ensure_rPr(r_el):
    rPr = r_el.find(W_RPR)
    if rPr is None:
        rPr = OxmlElement("w:rPr")
        r_el.insert(0, rPr)
    return rPr


def ancestor_is(el, tag) -> bool:
    parent = el.getparent()
    while parent is not None:
        if parent.tag == tag:
            return True
        parent = parent.getparent()
    return False


def nearest_tbl(el):
    parent = el.getparent()
    while parent is not None:
        if parent.tag == W_TBL:
            return parent
        parent = parent.getparent()
    return None


def table_has_special_layout(tbl) -> bool:
    if tbl is None:
        return False
    return (
        tbl.find(".//{0}".format(W_DRAWING)) is not None
        or tbl.find(".//{0}".format(W_PICT)) is not None
        or tbl.find(".//{0}".format(M_OMATH)) is not None
        or tbl.find(".//{0}".format(M_OMATH_PARA)) is not None
    )


def should_protect_layout(p_el) -> bool:
    """分段函数大括号、公式、浮动图所在段落：不能改缩进和行距。"""
    if p_el.find(".//{0}".format(WP_ANCHOR)) is not None:
        return True
    if p_el.find(".//{0}".format(M_OMATH)) is not None:
        return True
    if p_el.find(".//{0}".format(M_OMATH_PARA)) is not None:
        return True
    tbl = nearest_tbl(p_el)
    if tbl is None:
        return False
    # 单元格里的嵌套小表（业绩报酬={...}）整表保护
    if ancestor_is(tbl, W_TC) and table_has_special_layout(tbl):
        return True
    return False


def in_table(p_el) -> bool:
    return ancestor_is(p_el, W_TBL)


def cell_position(p_el):
    """返回 (列序号从0起, 该行列数)。不在单元格内则 None。"""
    tc = p_el
    while tc is not None and tc.tag != W_TC:
        tc = tc.getparent()
    if tc is None:
        return None
    tr = tc.getparent()
    while tr is not None and tr.tag != W_TR:
        tr = tr.getparent()
    if tr is None:
        return None
    cells = [c for c in tr if c.tag == W_TC]
    try:
        return cells.index(tc), len(cells)
    except ValueError:
        return None


def is_table_field_label(p_el, text: str) -> bool:
    """名称/投资范围等字段名：非末列、短文本、不像正文。"""
    if should_protect_layout(p_el):
        return False
    pos = cell_position(p_el)
    if pos is None:
        return False
    idx, n = pos
    if n < 2 or idx >= n - 1:
        return False
    t = "".join((text or "").split())
    if not t or len(t) > 40:
        return False
    if "。" in t or "；" in t:
        return False
    if t.endswith("=") or t.endswith("＝"):
        return False
    return True


def in_textbox(p_el) -> bool:
    return ancestor_is(p_el, W_TXBX)


def has_drawing(p_el) -> bool:
    return (
        p_el.find(".//{0}".format(W_DRAWING)) is not None
        or p_el.find(".//{0}".format(W_PICT)) is not None
        or p_el.find(".//{0}".format(W_OBJECT)) is not None
    )


def is_page_number_para(p_el) -> bool:
    """页眉页脚里的 PAGE 域。"""
    for instr in p_el.iter(qn("w:instrText")):
        if instr.text and "PAGE" in instr.text.upper():
            return True
    for fld in p_el.iter(qn("w:fldSimple")):
        instr = fld.get(qn("w:instr")) or ""
        if "PAGE" in instr.upper():
            return True
    return False


def iter_runs(p_el):
    for r_el in p_el.iter(W_R):
        # iter() 会进嵌套，但 run 一般不会再套 run
        yield r_el


def paragraph_plain_text(p_el) -> str:
    return "".join(node.text or "" for node in p_el.iter(W_T))


def is_effectively_empty(p_el) -> bool:
    if has_drawing(p_el):
        return False
    text = paragraph_plain_text(p_el)
    return not (text or "").strip()


def pt_to_twips(pt: float) -> int:
    return int(round(float(pt) * 20))


def chars_to_twips(chars: float, font_size_pt: float) -> int:
    return int(round(float(chars) * float(font_size_pt) * 20))


def pt_to_half_points(pt: float) -> str:
    return str(int(round(float(pt) * 2)))


def line_value(rule: str, spacing: float) -> int:
    """auto: 240 = 单倍；exact: 磅 * 20 twips。"""
    if rule == "exact":
        return pt_to_twips(spacing)
    return int(round(float(spacing) * 240))
