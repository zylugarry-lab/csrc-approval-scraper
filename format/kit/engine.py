from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, Iterable, List, Optional

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from .config import AppConfig, load_config
from .formatter import format_paragraph
from .formula import normalize_piecewise_formulas
from .xmlutil import (
    W_P,
    W_SECTPR,
    W_TC,
    W_TR,
    W_TBL,
    ancestor_is,
    drop_children,
    in_table,
    in_textbox,
    is_effectively_empty,
    nearest_tbl,
    table_has_special_layout,
)


@dataclass
class RefreshReport:
    input_path: str
    output_path: str
    body_paragraphs: int = 0
    table_paragraphs: int = 0
    header_paragraphs: int = 0
    textbox_paragraphs: int = 0
    highlights_cleared: int = 0
    shading_cleared: int = 0
    leading_stripped: int = 0
    numpr_removed: int = 0
    empty_collapsed: int = 0
    files: int = 1
    errors: List[str] = field(default_factory=list)

    def add_stats(self, stats: Dict[str, int], zone: str) -> None:
        self.highlights_cleared += stats.get("highlight", 0)
        self.shading_cleared += stats.get("shading", 0)
        self.leading_stripped += stats.get("stripped", 0)
        self.numpr_removed += stats.get("numpr_removed", 0)
        n = stats.get("paragraphs", 0)
        if zone == "table":
            self.table_paragraphs += n
        elif zone == "header":
            self.header_paragraphs += n
        elif zone == "textbox":
            self.textbox_paragraphs += n
        else:
            self.body_paragraphs += n

    def merge(self, other: "RefreshReport") -> None:
        self.files += other.files
        self.body_paragraphs += other.body_paragraphs
        self.table_paragraphs += other.table_paragraphs
        self.header_paragraphs += other.header_paragraphs
        self.textbox_paragraphs += other.textbox_paragraphs
        self.highlights_cleared += other.highlights_cleared
        self.shading_cleared += other.shading_cleared
        self.leading_stripped += other.leading_stripped
        self.numpr_removed += other.numpr_removed
        self.empty_collapsed += other.empty_collapsed
        self.errors.extend(other.errors)

    @property
    def total_paragraphs(self) -> int:
        return (
            self.body_paragraphs
            + self.table_paragraphs
            + self.header_paragraphs
            + self.textbox_paragraphs
        )

    def as_text(self) -> str:
        lines = [
            "处理完成：{0}".format(self.output_path),
            "段落合计 {0}（正文 {1} / 表格 {2} / 页眉页脚 {3} / 文本框 {4}）".format(
                self.total_paragraphs,
                self.body_paragraphs,
                self.table_paragraphs,
                self.header_paragraphs,
                self.textbox_paragraphs,
            ),
            "清除高亮 {0} 处，清除字符底纹 {1} 处".format(
                self.highlights_cleared, self.shading_cleared
            ),
            "去掉段首空格 {0} 段，去掉自动编号 {1} 段，折叠空段 {2} 处".format(
                self.leading_stripped, self.numpr_removed, self.empty_collapsed
            ),
        ]
        if self.errors:
            lines.append("失败 {0} 个：".format(len(self.errors)))
            lines.extend("  - " + e for e in self.errors)
        return "\n".join(lines)


def _zone_for(p_el, in_header: bool) -> str:
    if in_header:
        return "header"
    if in_table(p_el):
        return "table"
    if in_textbox(p_el):
        return "textbox"
    return "body"


def _should_process(zone: str, cfg: AppConfig) -> bool:
    o = cfg.options
    if zone == "body":
        return o.process_body
    if zone == "table":
        return o.process_tables
    if zone == "header":
        return o.process_headers_footers
    if zone == "textbox":
        return o.process_textboxes
    return True


def _specs_for(zone: str, cfg: AppConfig):
    if zone == "table":
        return cfg.table_font, cfg.table_para
    if zone == "header":
        return cfg.header_font, cfg.header_para
    if zone == "textbox":
        return cfg.body_font, cfg.body_para
    return cfg.body_font, cfg.body_para


def _iter_root_elements(doc) -> Iterable:
    yield doc.element.body, False
    seen = set()
    if not hasattr(doc, "sections"):
        return
    for section in doc.sections:
        for attr in (
            "header",
            "footer",
            "first_page_header",
            "first_page_footer",
            "even_page_header",
            "even_page_footer",
        ):
            part = getattr(section, attr, None)
            if part is None:
                continue
            el = getattr(part, "_element", None)
            if el is None:
                continue
            ident = id(el)
            if ident in seen:
                continue
            seen.add(ident)
            yield el, True


def disable_document_grid(doc) -> None:
    for sect in doc.element.iter(W_SECTPR):
        grid = sect.find(qn("w:docGrid"))
        if grid is not None:
            grid.set(qn("w:type"), "none")


def collapse_empty_in_tables(root) -> int:
    """同一单元格内连续空段只留一个；若单元格还有正文，则空段全部去掉。"""
    removed = 0
    for tc in root.iter(W_TC):
        tbl = nearest_tbl(tc)
        if tbl is not None and ancestor_is(tbl, W_TC) and table_has_special_layout(tbl):
            continue
        children = list(tc)
        paras = [ch for ch in children if ch.tag == W_P]
        if len(paras) <= 1:
            continue
        empties = [p for p in paras if is_effectively_empty(p)]
        nonempty = [p for p in paras if p not in empties]
        if nonempty:
            targets = empties
        else:
            targets = empties[1:]
        for p in targets:
            parent = p.getparent()
            if parent is not None:
                parent.remove(p)
                removed += 1
    return removed


def fit_table_layout(root) -> int:
    """单元格内容自动贴合：垂直顶端对齐，去掉固定/最小行高。公式小表除外。"""
    changed = 0
    for tbl in root.iter(W_TBL):
        if ancestor_is(tbl, W_TC) and table_has_special_layout(tbl):
            continue
        for tr in tbl.findall(W_TR):
            trPr = tr.find(qn("w:trPr"))
            if trPr is not None and drop_children(trPr, "w:trHeight"):
                changed += 1
            for tc in [c for c in tr if c.tag == W_TC]:
                tcPr = tc.find(qn("w:tcPr"))
                if tcPr is None:
                    tcPr = OxmlElement("w:tcPr")
                    tc.insert(0, tcPr)
                va = tcPr.find(qn("w:vAlign"))
                if va is None:
                    va = OxmlElement("w:vAlign")
                    tcPr.append(va)
                    va.set(qn("w:val"), "top")
                    changed += 1
                elif va.get(qn("w:val")) != "top":
                    va.set(qn("w:val"), "top")
                    changed += 1
    return changed


def _process_root(root, cfg: AppConfig, report: RefreshReport, in_header: bool) -> None:
    if cfg.options.collapse_empty_in_tables:
        report.empty_collapsed += collapse_empty_in_tables(root)

    for p_el in list(root.iter(W_P)):
        zone = _zone_for(p_el, in_header)
        if not _should_process(zone, cfg):
            continue
        # 文本框若已算 textbox，就不要再按 body 处理；表格优先
        font, para = _specs_for(zone, cfg)
        stats = format_paragraph(p_el, cfg, font, para, in_table=(zone == "table"))
        report.add_stats(stats, zone)

    if cfg.options.fit_cells:
        fit_table_layout(root)
    if cfg.options.normalize_formulas:
        normalize_piecewise_formulas(root)


def refresh_document(
    input_path,
    output_path=None,
    config: Optional[AppConfig] = None,
    config_path: Optional[Path] = None,
) -> RefreshReport:
    src = Path(input_path)
    if output_path is None:
        output_path = src.with_name(src.stem + "_已刷新" + src.suffix)
    dst = Path(output_path)
    dst.parent.mkdir(parents=True, exist_ok=True)

    if config is None:
        config = load_config(config_path)

    report = RefreshReport(input_path=str(src), output_path=str(dst))
    doc = Document(str(src))

    if config.options.disable_document_grid:
        disable_document_grid(doc)

    for root, in_header in _iter_root_elements(doc):
        _process_root(root, config, report, in_header)

    # 正文里的文本框（drawing 内）可能不在 body.iter 的常规结构外，root.iter(W_P) 已覆盖。
    # 再扫一遍 txbx，防止被 options.process_body=False 漏掉。
    if config.options.process_textboxes:
        pass

    doc.save(str(dst))
    return report


def refresh_many(paths: List[Path], output_dir: Optional[Path], config: AppConfig) -> RefreshReport:
    if not paths:
        raise ValueError("没有可处理的 .docx 文件")
    first = refresh_document(paths[0], _out_path(paths[0], output_dir), config)
    summary = first
    summary.files = 1
    for path in paths[1:]:
        try:
            one = refresh_document(path, _out_path(path, output_dir), config)
            summary.merge(one)
        except Exception as exc:
            summary.errors.append("{0}: {1}".format(path.name, exc))
            summary.files += 1
    return summary


def _out_path(src: Path, output_dir: Optional[Path]) -> Path:
    name = src.stem + "_已刷新" + src.suffix
    if output_dir is None:
        return src.with_name(name)
    return Path(output_dir) / name


def collect_docx(path: Path) -> List[Path]:
    path = Path(path)
    if path.is_file():
        if path.suffix.lower() != ".docx":
            raise ValueError("仅支持 .docx（请先在 Word / WPS 另存为 docx）")
        if path.name.startswith("~$"):
            raise ValueError("这是 Word 临时锁文件，请关闭文档后选原文件")
        return [path]
    files = [
        p
        for p in sorted(path.glob("*.docx"))
        if not p.name.startswith("~$") and "_已刷新" not in p.stem
    ]
    if not files:
        raise ValueError("文件夹里没有 .docx")
    return files
