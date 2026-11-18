from __future__ import annotations

import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, Optional

import yaml

SIZE_NAMES = {
    "初号": 42.0,
    "小初": 36.0,
    "一号": 26.0,
    "小一": 24.0,
    "二号": 22.0,
    "小二": 18.0,
    "三号": 16.0,
    "小三": 15.0,
    "四号": 14.0,
    "小四": 12.0,
    "五号": 10.5,
    "小五": 9.0,
    "六号": 7.5,
    "小六": 6.5,
    "七号": 5.5,
    "八号": 5.0,
}


def parse_size(value) -> float:
    if value is None:
        return 12.0
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if text in SIZE_NAMES:
        return SIZE_NAMES[text]
    if text.endswith("磅"):
        text = text[:-1]
    return float(text)


def parse_color(value) -> str:
    text = str(value or "000000").strip().lstrip("#")
    if len(text) == 3:
        text = "".join(ch * 2 for ch in text)
    return text.upper().zfill(6)[:6]


@dataclass
class FontSpec:
    east_asia: str = "宋体"
    ascii: str = "Times New Roman"
    size_pt: float = 12.0
    color: str = "000000"


@dataclass
class ParaSpec:
    alignment: str = "left"  # left / center / right / both / keep
    first_line_chars: float = 0.0
    line_spacing_rule: str = "auto"  # auto | exact
    line_spacing: float = 1.5
    space_before_pt: float = 0.0
    space_after_pt: float = 0.0


@dataclass
class ListSpec:
    left: float = 2.0
    hanging: float = 2.0


@dataclass
class Options:
    process_body: bool = True
    process_tables: bool = True
    process_headers_footers: bool = True
    process_textboxes: bool = True
    clear_highlight: bool = True
    clear_run_shading: bool = True
    preserve_bold: bool = True
    preserve_italic: bool = True
    preserve_underline: bool = True
    preserve_super_sub: bool = True
    strip_leading_whitespace: bool = True
    collapse_empty_in_tables: bool = True
    disable_document_grid: bool = True
    skip_centered: bool = False
    skip_right_aligned: bool = True
    preserve_centered_size: bool = True
    remove_auto_num_if_text_has_marker: bool = True
    flush_all: bool = True
    protect_formula_layout: bool = True
    center_cover: bool = True
    center_page_numbers: bool = True
    bold_table_labels: bool = True
    fit_cells: bool = True
    normalize_formulas: bool = True


@dataclass
class AppConfig:
    body_font: FontSpec = field(default_factory=FontSpec)
    body_para: ParaSpec = field(default_factory=ParaSpec)
    table_font: FontSpec = field(default_factory=lambda: FontSpec(size_pt=10.5))
    table_para: ParaSpec = field(
        default_factory=lambda: ParaSpec(
            alignment="left", first_line_chars=0.0, line_spacing=1.0
        )
    )
    header_font: FontSpec = field(default_factory=lambda: FontSpec(size_pt=9.0))
    header_para: ParaSpec = field(
        default_factory=lambda: ParaSpec(
            alignment="left", first_line_chars=0.0, line_spacing=1.0
        )
    )
    lists: Dict[str, ListSpec] = field(default_factory=dict)
    extra_left_chars: float = 0.0
    options: Options = field(default_factory=Options)
    raw: Dict[str, Any] = field(default_factory=dict)

    def summary_lines(self) -> list:
        o = self.options
        return [
            "中文字体 {0}　西文 {1}".format(self.body_font.east_asia, self.body_font.ascii),
            "正文 {0:g} 磅　表格 {1:g} 磅　颜色 #{2}".format(
                self.body_font.size_pt, self.table_font.size_pt, self.body_font.color
            ),
            "正文行距 {0}　表格行距 {1}".format(
                _spacing_label(self.body_para), _spacing_label(self.table_para)
            ),
            "正文首行缩进 {0:g} 字　表格内正文缩进 {1:g} 字　全部顶格：{2}".format(
                self.body_para.first_line_chars,
                self.table_para.first_line_chars,
                _yn(o.flush_all),
            ),
            "清除高亮：{0}　处理表格：{1}　处理页眉页脚：{2}".format(
                _yn(o.clear_highlight), _yn(o.process_tables), _yn(o.process_headers_footers)
            ),
        ]


def _yn(flag: bool) -> str:
    return "是" if flag else "否"


def _spacing_label(para: ParaSpec) -> str:
    if para.line_spacing_rule == "exact":
        return "固定值 {0:g} 磅".format(para.line_spacing)
    return "{0:g} 倍".format(para.line_spacing)


def _section(data: dict, key: str) -> dict:
    value = data.get(key) or {}
    return value if isinstance(value, dict) else {}


def _font_from(section: dict, default: FontSpec) -> FontSpec:
    return FontSpec(
        east_asia=str(section.get("font_east_asia") or default.east_asia),
        ascii=str(section.get("font_ascii") or default.ascii),
        size_pt=parse_size(section.get("size", default.size_pt)),
        color=parse_color(section.get("color", default.color)),
    )


def _para_from(section: dict, default: ParaSpec) -> ParaSpec:
    return ParaSpec(
        alignment=str(section.get("alignment") or default.alignment),
        first_line_chars=float(section.get("first_line_chars", default.first_line_chars)),
        line_spacing_rule=str(section.get("line_spacing_rule") or default.line_spacing_rule),
        line_spacing=float(section.get("line_spacing", default.line_spacing)),
        space_before_pt=float(section.get("space_before_pt", default.space_before_pt)),
        space_after_pt=float(section.get("space_after_pt", default.space_after_pt)),
    )


def _options_from(section: dict) -> Options:
    opts = Options()
    for name in opts.__dataclass_fields__:
        if name in section:
            setattr(opts, name, bool(section[name]))
    return opts


def load_config(path: Optional[Path] = None) -> AppConfig:
    cfg = AppConfig()
    cfg.lists = {
        "cn1": ListSpec(0, 0),
        "cn2": ListSpec(0, 0),
        "num_dot": ListSpec(0, 0),
        "num_paren": ListSpec(0, 0),
        "circled": ListSpec(0, 0),
    }
    cfg.extra_left_chars = 0.0
    if path is None:
        if getattr(sys, "frozen", False):
            path = Path(sys.executable).resolve().parent / "config.yaml"
        else:
            path = Path(__file__).resolve().parent.parent / "config.yaml"
    if not path.exists():
        return cfg

    with path.open("r", encoding="utf-8") as fh:
        data = yaml.safe_load(fh) or {}
    cfg.raw = data

    body = _section(data, "body")
    table = _section(data, "table")
    header = _section(data, "header_footer")
    lists = _section(data, "lists")
    options = _section(data, "options")

    cfg.body_font = _font_from(body, cfg.body_font)
    cfg.body_para = _para_from(body, cfg.body_para)
    cfg.table_font = _font_from(table, cfg.table_font)
    cfg.table_para = _para_from(table, cfg.table_para)
    cfg.header_font = _font_from(header, cfg.header_font)
    cfg.header_para = _para_from(header, cfg.header_para)
    cfg.options = _options_from(options)

    for key in ("cn1", "cn2", "num_dot", "num_paren", "circled"):
        item = lists.get(key)
        if isinstance(item, dict):
            cfg.lists[key] = ListSpec(
                left=float(item.get("left", cfg.lists[key].left)),
                hanging=float(item.get("hanging", cfg.lists[key].hanging)),
            )
    if "extra_left_chars" in lists:
        cfg.extra_left_chars = float(lists["extra_left_chars"])
    return cfg
