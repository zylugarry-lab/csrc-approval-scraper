#!/usr/bin/env python3
"""根据 YAML 数据，填充说明书模板中「黄色字段名」右侧的内容，生成新说明书。"""

from __future__ import annotations

import argparse
import copy
import sys
from pathlib import Path

import yaml
from docx import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

ROOT = Path(__file__).resolve().parent
DEFAULT_TEMPLATE = ROOT / "template" / "说明书模板.docx"
DEFAULT_OUTPUT_DIR = ROOT / "output"


def is_yellow(run) -> bool:
    rPr = run._element.find(qn("w:rPr"))
    if rPr is None:
        return False
    highlight = rPr.find(qn("w:highlight"))
    if highlight is None:
        return False
    val = highlight.get(qn("w:val"))
    return bool(val and val != "none")


def yellow_texts(para: Paragraph) -> list[str]:
    parts: list[str] = []
    buf = ""
    cur = False
    for run in para.runs:
        if is_yellow(run):
            buf += run.text or ""
            cur = True
        else:
            if cur and buf.strip():
                parts.append(buf.strip())
            buf, cur = "", False
    if cur and buf.strip():
        parts.append(buf.strip())
    return parts


def clear_highlight(run) -> None:
    rPr = run._element.find(qn("w:rPr"))
    if rPr is None:
        return
    for highlight in rPr.findall(qn("w:highlight")):
        rPr.remove(highlight)


def unique_cells(row):
    seen = set()
    cells = []
    for ci, cell in enumerate(row.cells):
        cid = id(cell._tc)
        if cid in seen:
            continue
        seen.add(cid)
        cells.append((ci, cell))
    return cells


def set_paragraph_text_keep_style(paragraph: Paragraph, text: str) -> None:
    """用一段文本替换段落内容，尽量保留首个 run 的样式。"""
    if paragraph.runs:
        first = paragraph.runs[0]
        first.text = text
        for run in paragraph.runs[1:]:
            run.text = ""
        return
    paragraph.add_run(text)


def _cell_block_elements(cell):
    """按文档顺序返回单元格内段落与嵌套表。"""
    paras, tables = [], []
    for child in cell._tc:
        tag = child.tag.split("}")[-1]
        if tag == "p":
            paras.append(child)
        elif tag == "tbl":
            tables.append(child)
    return paras, tables


def _append_para_to_cell(cell, proto_p, line: str) -> None:
    if proto_p is not None:
        new_p = copy.deepcopy(proto_p)
        for child in list(new_p):
            if child.tag == qn("w:r"):
                new_p.remove(child)
        cell._tc.append(new_p)
        Paragraph(new_p, cell).add_run(line)
    else:
        cell.add_paragraph(line)


def set_cell_text(cell, text: str) -> None:
    """按换行拆成多段写入单元格；单元格内嵌套小表（如持有时间/退出费率分档）直接去掉。"""
    lines = text.split("\n") if text is not None else [""]
    if not lines:
        lines = [""]

    p_els, tbl_els = _cell_block_elements(cell)

    # 去掉嵌套表后按段落重写
    if tbl_els:
        proto = copy.deepcopy(p_els[0]) if p_els else None
        for el in p_els + tbl_els:
            cell._tc.remove(el)
        for line in lines:
            _append_para_to_cell(cell, proto, line)
        return

    paragraphs = cell.paragraphs
    for i, line in enumerate(lines):
        if i < len(paragraphs):
            set_paragraph_text_keep_style(paragraphs[i], line)
        else:
            new_p = copy.deepcopy(paragraphs[0]._p)
            for child in list(new_p):
                if child.tag == qn("w:r"):
                    new_p.remove(child)
            cell._tc.append(new_p)
            Paragraph(new_p, cell).add_run(line)
    for para in paragraphs[len(lines) :]:
        para._p.getparent().remove(para._p)


def replace_yellow_run_text(paragraph: Paragraph, new_text: str) -> bool:
    """将段落中连续黄色高亮 runs 的文本替换为 new_text（合并到第一个黄 run）。"""
    yellow_runs = [r for r in paragraph.runs if is_yellow(r)]
    if not yellow_runs:
        return False
    # 若存在多段黄色（如封面产品名与年份在不同段），调用方按段处理
    # 这里把「连续黄色块」逐块替换为同一 new_text 不适用；改为：整段内所有黄色拼成一块后整体替换
    # 更稳妥：找到第一个黄 run，写入全文，其余黄 run 清空
    yellow_runs[0].text = new_text
    for r in yellow_runs[1:]:
        r.text = ""
    return True


def replace_in_runs(paragraph: Paragraph, old: str, new: str) -> int:
    if not old or old == new:
        return 0
    # 先尝试单 run 内替换
    count = 0
    for run in paragraph.runs:
        if old in (run.text or ""):
            run.text = (run.text or "").replace(old, new)
            count += 1
    if count:
        return count
    # 跨 run：拼全文再重写（保留首 run 样式）
    full = "".join(r.text or "" for r in paragraph.runs)
    if old not in full:
        return 0
    set_paragraph_text_keep_style(paragraph, full.replace(old, new))
    return 1


def resolve_field_value(fields: dict, label: str, section: str | None) -> str | None:
    """按优先级查找字段值。"""
    candidates = []
    if section:
        candidates.append(f"{section}·{label}")
    candidates.append(label)
    # 常见别名
    if label == "办理时间" and section:
        candidates.append(f"{section}·办理时间")

    for key in candidates:
        if key in fields and fields[key] is not None:
            val = fields[key]
            if isinstance(val, list):
                return "\n".join(str(x) for x in val)
            return str(val)
    return None


def fill_document(doc: Document, data: dict, *, clear_marks: bool = True) -> list[str]:
    product = str(data.get("产品简称") or "").strip()
    year = str(data.get("年份") or "").strip()
    fields = data.get("字段") or {}
    if not isinstance(fields, dict):
        raise ValueError("YAML 中「字段」必须是字典")

    # 从模板封面推断旧产品简称 / 旧年份，用于全文替换
    table = doc.tables[0]
    old_product = None
    old_year = None
    cover_paras = table.rows[0].cells[0].paragraphs
    if cover_paras:
        ys = yellow_texts(cover_paras[0])
        if ys:
            old_product = ys[0]
        if len(cover_paras) > 1:
            ys2 = yellow_texts(cover_paras[1])
            if ys2:
                old_year = ys2[0]

    report: list[str] = []
    filled_labels: list[str] = []
    missing: list[str] = []

    for ri, row in enumerate(table.rows):
        cells = unique_cells(row)

        # —— 封面 ——
        if ri == 0:
            if product and cells:
                # 标题段：黄色 = 产品简称
                paras = cells[0][1].paragraphs
                if paras and yellow_texts(paras[0]) is not None:
                    if yellow_texts(paras[0]):
                        replace_yellow_run_text(paras[0], product)
                        report.append(f"封面标题产品简称 → {product}")
                if year and len(paras) > 1 and yellow_texts(paras[1]):
                    replace_yellow_run_text(paras[1], year)
                    report.append(f"封面年份 → {year}")
            continue

        # —— 特别提示：黄色处为产品简称 ——
        if ri == 1:
            if product and cells:
                for p in cells[0][1].paragraphs:
                    if yellow_texts(p):
                        replace_yellow_run_text(p, product)
                        report.append(f"特别提示中产品简称 → {product}")
            continue

        section = None
        label = None
        label_ci = None
        for ci, cell in cells:
            ys: list[str] = []
            for p in cell.paragraphs:
                ys.extend(yellow_texts(p))
            text = "\n".join(p.text for p in cell.paragraphs).strip()
            if len(cells) == 3 and ci == 0 and not ys:
                section = text or None
            if ys:
                label = ys[-1]
                label_ci = ci

        if label is None:
            continue

        value_cell = None
        for ci, cell in reversed(cells):
            if ci != label_ci:
                value_cell = cell
                break
        if value_cell is None:
            continue

        display_key = f"{section}·{label}" if section else label
        value = resolve_field_value(fields, label, section)
        if value is None:
            missing.append(display_key)
            continue

        set_cell_text(value_cell, value)
        filled_labels.append(display_key)

    # 全文替换旧产品简称 / 旧年份（覆盖未单独配置但仍含旧名的段落）
    if product and old_product and product != old_product:
        n = 0
        for row in table.rows:
            for _, cell in unique_cells(row):
                for p in cell.paragraphs:
                    n += replace_in_runs(p, old_product, product)
        report.append(f"全文替换产品简称 {old_product!r} → {product!r}（{n} 处）")

    if year and old_year and year != old_year:
        n = 0
        for row in table.rows:
            for _, cell in unique_cells(row):
                for p in cell.paragraphs:
                    n += replace_in_runs(p, old_year, year)
        report.append(f"全文替换年份 {old_year!r} → {year!r}（{n} 处）")

    if clear_marks:
        for row in table.rows:
            for _, cell in unique_cells(row):
                for p in cell.paragraphs:
                    for run in p.runs:
                        clear_highlight(run)

    report.append(f"已填充字段 {len(filled_labels)} 个")
    if missing:
        report.append("未在 YAML 中找到、已保留原文的字段：")
        for m in missing:
            report.append(f"  - {m}")

    matched_keys: set[str] = set()
    for k in fields:
        for filled in filled_labels:
            if k == filled or k.split("·")[-1] == filled.split("·")[-1]:
                matched_keys.add(k)
                break
    unused = [k for k in fields if k not in matched_keys]
    if unused:
        report.append("YAML 中有但本次未匹配到的键（请检查命名）：")
        for u in unused:
            report.append(f"  - {u}")

    return report


def extract_to_yaml(template_path: Path, out_path: Path) -> None:
    """从带高亮的模板反抽一份 YAML，便于改数。"""
    doc = Document(str(template_path))
    table = doc.tables[0]

    product = ""
    year = ""
    cover = table.rows[0].cells[0].paragraphs
    if cover:
        ys = yellow_texts(cover[0])
        if ys:
            product = ys[0]
        if len(cover) > 1:
            ys2 = yellow_texts(cover[1])
            if ys2:
                year = ys2[0]

    fields: dict[str, str] = {}
    for ri, row in enumerate(table.rows):
        if ri < 2:
            continue
        cells = unique_cells(row)
        section = None
        label = None
        label_ci = None
        for ci, cell in cells:
            ys: list[str] = []
            for p in cell.paragraphs:
                ys.extend(yellow_texts(p))
            text = "\n".join(p.text for p in cell.paragraphs).strip()
            if len(cells) == 3 and ci == 0 and not ys:
                section = text or None
            if ys:
                label = ys[-1]
                label_ci = ci
        if label is None:
            continue
        value_cell = None
        for ci, cell in reversed(cells):
            if ci != label_ci:
                value_cell = cell
                break
        paras = [p.text for p in value_cell.paragraphs] if value_cell else []
        while paras and paras[-1] == "":
            paras.pop()
        key = label
        if key in fields:
            key = f"{section}·{label}" if section else f"{label}_{ri}"
        # 参与/退出办理时间统一带章节前缀，避免歧义
        if label == "办理时间" and section:
            key = f"{section}·办理时间"
            # 若之前已用裸名写入，迁走
            if "办理时间" in fields and key != "办理时间":
                # 仅当当前是第二个办理时间时；第一个应也用带前缀
                pass
        fields[key] = "\n".join(paras)

    # 把第一个裸「办理时间」改成带章节名（若能从模板推断）
    # 重新扫一遍更干净
    fields = {}
    for ri, row in enumerate(table.rows):
        if ri < 2:
            continue
        cells = unique_cells(row)
        section = None
        label = None
        label_ci = None
        for ci, cell in cells:
            ys = []
            for p in cell.paragraphs:
                ys.extend(yellow_texts(p))
            text = "\n".join(p.text for p in cell.paragraphs).strip()
            if len(cells) == 3 and ci == 0 and not ys:
                section = text or None
            if ys:
                label = ys[-1]
                label_ci = ci
        if label is None:
            continue
        value_cell = None
        for ci, cell in reversed(cells):
            if ci != label_ci:
                value_cell = cell
                break
        paras = [p.text for p in value_cell.paragraphs] if value_cell else []
        while paras and paras[-1] == "":
            paras.pop()
        if label == "办理时间" and section:
            key = f"{section}·办理时间"
        else:
            key = label
            if key in fields:
                key = f"{section}·{label}" if section else f"{label}_{ri}"
        fields[key] = "\n".join(paras)

    data = {"产品简称": product, "年份": year, "字段": fields}

    def str_presenter(dumper, data):
        if isinstance(data, str) and ("\n" in data or len(data) > 60):
            return dumper.represent_scalar("tag:yaml.org,2002:str", data, style="|")
        return dumper.represent_scalar("tag:yaml.org,2002:str", data)

    yaml.add_representer(str, str_presenter)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as f:
        f.write("# 由模板反抽的说明书填报数据，按需修改后用于 generate\n\n")
        yaml.dump(data, f, allow_unicode=True, sort_keys=False, width=100)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="根据 YAML 自动生成集合计划说明书")
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_gen = sub.add_parser("generate", help="用 YAML 填充模板生成说明书")
    p_gen.add_argument("-d", "--data", required=True, type=Path, help="填报 YAML 路径")
    p_gen.add_argument("-t", "--template", type=Path, default=DEFAULT_TEMPLATE, help="模板 docx")
    p_gen.add_argument("-o", "--output", type=Path, default=None, help="输出 docx 路径")
    p_gen.add_argument("--keep-highlight", action="store_true", help="保留黄色高亮标记")

    p_ext = sub.add_parser("extract", help="从带高亮的模板反抽 YAML")
    p_ext.add_argument("-t", "--template", type=Path, default=DEFAULT_TEMPLATE, help="模板 docx")
    p_ext.add_argument("-o", "--output", type=Path, required=True, help="输出 YAML 路径")

    args = parser.parse_args(argv)

    if args.cmd == "extract":
        extract_to_yaml(args.template, args.output)
        print(f"已导出: {args.output}")
        return 0

    if not args.template.exists():
        print(f"找不到模板: {args.template}", file=sys.stderr)
        return 1
    if not args.data.exists():
        print(f"找不到数据文件: {args.data}", file=sys.stderr)
        return 1

    with open(args.data, encoding="utf-8") as f:
        data = yaml.safe_load(f) or {}

    doc = Document(str(args.template))
    if not doc.tables:
        print("模板中没有表格", file=sys.stderr)
        return 1

    report = fill_document(doc, data, clear_marks=not args.keep_highlight)

    product = str(data.get("产品简称") or "说明书").strip()
    out = args.output
    if out is None:
        DEFAULT_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        out = DEFAULT_OUTPUT_DIR / f"{product}集合资产管理计划说明书-自动生成.docx"
    else:
        out.parent.mkdir(parents=True, exist_ok=True)

    doc.save(str(out))
    print(f"已生成: {out}")
    for line in report:
        print(line)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
