from __future__ import annotations

from dataclasses import replace

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from .classifier import ParaKind, RE_LEAD, classify, is_cover_line, normalize_for_match, paragraph_text
from .config import AppConfig, FontSpec, ParaSpec
from .xmlutil import (
    W_PPR,
    W_T,
    chars_to_twips,
    drop_attrs,
    drop_children,
    ensure_pPr,
    ensure_rPr,
    get_or_create,
    is_page_number_para,
    is_table_field_label,
    iter_runs,
    line_value,
    pt_to_half_points,
    pt_to_twips,
    should_protect_layout,
)

ALIGN_MAP = {
    "left": "left",
    "center": "center",
    "right": "right",
    "both": "both",
    "justify": "both",
}


def current_jc(p_el) -> str:
    pPr = p_el.find(W_PPR)
    if pPr is None:
        return ""
    jc = pPr.find(qn("w:jc"))
    if jc is None:
        return ""
    return (jc.get(qn("w:val")) or "").lower()


def strip_leading_whitespace(p_el) -> bool:
    changed = False
    leading = True
    for r_el in iter_runs(p_el):
        if not leading:
            break
        for tab in list(r_el.findall(qn("w:tab"))):
            r_el.remove(tab)
            changed = True
        for t_el in r_el.findall(W_T):
            if not leading:
                break
            text = t_el.text or ""
            new = RE_LEAD.sub("", text)
            if new != text:
                t_el.text = new
                changed = True
            if new:
                leading = False
    return changed


def remove_numPr(p_el) -> bool:
    pPr = p_el.find(W_PPR)
    if pPr is None:
        return False
    return drop_children(pPr, "w:numPr") > 0


def apply_font_to_rPr(rPr, font: FontSpec, cfg: AppConfig, force_bold: bool = False) -> dict:
    stats = {"highlight": 0, "shading": 0}

    rFonts = get_or_create(rPr, "w:rFonts")
    rFonts.set(qn("w:ascii"), font.ascii)
    rFonts.set(qn("w:hAnsi"), font.ascii)
    rFonts.set(qn("w:eastAsia"), font.east_asia)
    rFonts.set(qn("w:cs"), font.ascii)
    drop_attrs(rFonts, ["w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"])

    sz = pt_to_half_points(font.size_pt)
    get_or_create(rPr, "w:sz").set(qn("w:val"), sz)
    get_or_create(rPr, "w:szCs").set(qn("w:val"), sz)

    color = get_or_create(rPr, "w:color")
    color.set(qn("w:val"), font.color)
    drop_attrs(color, ["w:themeColor", "w:themeTint", "w:themeShade"])

    if cfg.options.clear_highlight:
        stats["highlight"] = drop_children(rPr, "w:highlight")
    if cfg.options.clear_run_shading:
        stats["shading"] = drop_children(rPr, "w:shd")

    drop_children(rPr, "w:spacing")
    drop_children(rPr, "w:w")
    drop_children(rPr, "w:kern")
    drop_children(rPr, "w:fitText")

    if force_bold:
        b = get_or_create(rPr, "w:b")
        drop_attrs(b, ["w:val"])
        get_or_create(rPr, "w:bCs")
    elif not cfg.options.preserve_bold:
        drop_children(rPr, "w:b")
        drop_children(rPr, "w:bCs")
    if not cfg.options.preserve_italic:
        drop_children(rPr, "w:i")
        drop_children(rPr, "w:iCs")
    if not cfg.options.preserve_underline:
        drop_children(rPr, "w:u")
    if not cfg.options.preserve_super_sub:
        drop_children(rPr, "w:vertAlign")

    lang = get_or_create(rPr, "w:lang")
    lang.set(qn("w:eastAsia"), "zh-CN")
    return stats


def apply_run_font(r_el, font: FontSpec, cfg: AppConfig, force_bold: bool = False) -> dict:
    return apply_font_to_rPr(ensure_rPr(r_el), font, cfg, force_bold=force_bold)


def apply_spacing(pPr, para: ParaSpec) -> None:
    drop_children(pPr, "w:spacing")
    sp = OxmlElement("w:spacing")
    sp.set(qn("w:before"), str(pt_to_twips(para.space_before_pt)))
    sp.set(qn("w:after"), str(pt_to_twips(para.space_after_pt)))
    sp.set(qn("w:line"), str(line_value(para.line_spacing_rule, para.line_spacing)))
    sp.set(qn("w:lineRule"), "exact" if para.line_spacing_rule == "exact" else "auto")
    pPr.append(sp)


def apply_indent(
    pPr,
    font_size_pt: float,
    left_chars: float,
    first_line_chars: float,
    hanging_chars: float,
) -> None:
    drop_children(pPr, "w:ind")
    ind = OxmlElement("w:ind")

    left_tw = chars_to_twips(max(left_chars, 0.0), font_size_pt)
    left_c = str(int(round(max(left_chars, 0.0) * 100)))
    ind.set(qn("w:left"), str(left_tw))
    ind.set(qn("w:leftChars"), left_c)
    ind.set(qn("w:start"), str(left_tw))
    ind.set(qn("w:startChars"), left_c)

    if hanging_chars and hanging_chars > 0:
        hang_tw = chars_to_twips(hanging_chars, font_size_pt)
        ind.set(qn("w:hanging"), str(hang_tw))
        ind.set(qn("w:hangingChars"), str(int(round(hanging_chars * 100))))
    else:
        first_tw = chars_to_twips(max(first_line_chars, 0.0), font_size_pt)
        ind.set(qn("w:firstLine"), str(first_tw))
        ind.set(qn("w:firstLineChars"), str(int(round(max(first_line_chars, 0.0) * 100))))

    pPr.append(ind)


def apply_alignment(pPr, value: str) -> None:
    if not value or value == "keep":
        return
    mapped = ALIGN_MAP.get(value, value)
    drop_children(pPr, "w:jc")
    jc = OxmlElement("w:jc")
    jc.set(qn("w:val"), mapped)
    pPr.append(jc)


def read_size_pt(p_el):
    for r_el in iter_runs(p_el):
        rPr = r_el.find(qn("w:rPr"))
        if rPr is None:
            continue
        sz = rPr.find(qn("w:sz"))
        if sz is None:
            continue
        val = sz.get(qn("w:val"))
        if not val:
            continue
        try:
            return int(val) / 2.0
        except ValueError:
            continue
    return None


def indent_for_kind(kind: ParaKind, cfg: AppConfig, in_table: bool, skip_indent: bool):
    if cfg.options.flush_all or skip_indent or kind.role in ("empty", "label", "heading"):
        return 0.0, 0.0, 0.0
    para = cfg.table_para if in_table else cfg.body_para
    if kind.role == "list" and kind.list_key:
        spec = cfg.lists.get(kind.list_key)
        if spec is None:
            return 0.0, para.first_line_chars, 0.0
        extra = kind.nest_level * cfg.extra_left_chars
        hanging = spec.hanging
        left = spec.left + extra
        if in_table:
            left = hanging
        return left, 0.0, hanging
    return 0.0, para.first_line_chars, 0.0


def format_paragraph(p_el, cfg: AppConfig, font: FontSpec, para: ParaSpec, in_table: bool) -> dict:
    stats = {
        "paragraphs": 1,
        "highlight": 0,
        "shading": 0,
        "stripped": 0,
        "numpr_removed": 0,
    }
    protect = cfg.options.protect_formula_layout and should_protect_layout(p_el)
    if cfg.options.strip_leading_whitespace and not protect:
        if strip_leading_whitespace(p_el):
            stats["stripped"] = 1

    text = paragraph_text(p_el)
    kind = classify(text, in_table)
    jc = current_jc(p_el)
    skip_indent = protect
    if cfg.options.skip_centered and jc in ("center", "centre"):
        skip_indent = True
    if cfg.options.skip_right_aligned and jc == "right":
        skip_indent = True

    if (not protect) and (
        cfg.options.flush_all
        or (
            cfg.options.remove_auto_num_if_text_has_marker
            and kind.role in ("list", "heading")
        )
    ):
        if remove_numPr(p_el):
            stats["numpr_removed"] = 1

    pPr = ensure_pPr(p_el)
    if not protect:
        apply_spacing(pPr, para)
        snap = get_or_create(pPr, "w:snapToGrid")
        snap.set(qn("w:val"), "0")
        drop_children(pPr, "w:contextualSpacing")
        if not is_page_number_para(p_el):
            drop_children(pPr, "w:tabs")
        left, first, hanging = indent_for_kind(kind, cfg, in_table, skip_indent)
        apply_indent(pPr, font.size_pt, left, first, hanging)
        skip_align = cfg.options.skip_right_aligned and jc == "right"
        if not skip_align:
            apply_alignment(pPr, para.alignment)
        if cfg.options.center_cover and is_cover_line(text):
            apply_alignment(pPr, "center")
        if cfg.options.center_page_numbers and is_page_number_para(p_el):
            apply_alignment(pPr, "center")

    font_to_use = font
    if protect or (skip_indent and not in_table and cfg.options.preserve_centered_size):
        orig = read_size_pt(p_el)
        if orig:
            font_to_use = replace(font, size_pt=orig)

    force_bold = bool(cfg.options.bold_table_labels and is_table_field_label(p_el, text))

    for r_el in iter_runs(p_el):
        run_stats = apply_run_font(r_el, font_to_use, cfg, force_bold=force_bold)
        stats["highlight"] += run_stats["highlight"]
        stats["shading"] += run_stats["shading"]

    p_rPr = get_or_create(pPr, "w:rPr")
    apply_font_to_rPr(p_rPr, font_to_use, cfg, force_bold=force_bold)
    return stats
