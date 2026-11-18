#!/usr/bin/env python3
"""根据 YAML 填充风险揭示书模板中黄亮对应内容，生成新风险揭示书。

规则（与说明书一致）：
- 黄色高亮多为字段名/可变片段
- 对应段落（如「特殊风险揭示」「一般风险揭示」标题下正文）由 YAML 填入
- 封面产品简称、风险等级、适合投资者、合同条款引用一并替换
"""

from __future__ import annotations

import argparse
import copy
import re
import sys
from pathlib import Path

import yaml
from docx import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

ROOT = Path(__file__).resolve().parent
DEFAULT_TEMPLATE = ROOT / "template" / "风险揭示书模板.docx"
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


def set_paragraph_text_keep_style(paragraph: Paragraph, text: str) -> None:
    if paragraph.runs:
        paragraph.runs[0].text = text
        for run in paragraph.runs[1:]:
            run.text = ""
        return
    paragraph.add_run(text)


def replace_in_runs(paragraph: Paragraph, old: str, new: str) -> int:
    if not old or old == new:
        return 0
    count = 0
    for run in paragraph.runs:
        if old in (run.text or ""):
            run.text = (run.text or "").replace(old, new)
            count += 1
    if count:
        return count
    full = "".join(r.text or "" for r in paragraph.runs)
    if old not in full:
        return 0
    set_paragraph_text_keep_style(paragraph, full.replace(old, new))
    return 1


def replace_yellow_chunks(paragraph: Paragraph, mapping: dict[str, str]) -> int:
    """按黄亮文本块替换；若旧值不在 mapping，则按顺序替换连续黄亮块。"""
    n = 0
    # 先尝试精确匹配旧黄亮文本
    for run in paragraph.runs:
        if not is_yellow(run):
            continue
        text = run.text or ""
        for old, new in mapping.items():
            if old and old in text:
                run.text = text.replace(old, new)
                n += 1
                text = run.text
    return n


def find_para_index(paras: list, pred) -> int | None:
    for i, p in enumerate(paras):
        if pred(p):
            return i
    return None


def delete_paragraph(paragraph: Paragraph) -> None:
    p = paragraph._element
    parent = p.getparent()
    if parent is not None:
        parent.remove(p)


def is_actively_numbered(paragraph: Paragraph) -> bool:
    """Word 自动编号有效（排除 numId=0 / ilvl=-1 的占位）。"""
    pPr = paragraph._element.find(qn("w:pPr"))
    if pPr is None:
        return False
    numPr = pPr.find(qn("w:numPr"))
    if numPr is None:
        return False
    ilvl = numPr.find(qn("w:ilvl"))
    numId = numPr.find(qn("w:numId"))
    ilvl_v = ilvl.get(qn("w:val")) if ilvl is not None else None
    num_v = numId.get(qn("w:val")) if numId is not None else None
    return num_v not in (None, "0") and ilvl_v not in (None, "-1")


def _num_id(paragraph: Paragraph) -> str | None:
    pPr = paragraph._element.find(qn("w:pPr"))
    if pPr is None:
        return None
    numPr = pPr.find(qn("w:numPr"))
    if numPr is None:
        return None
    numId = numPr.find(qn("w:numId"))
    return numId.get(qn("w:val")) if numId is not None else None


def _collect_numbering_prototypes(old_paras: list[Paragraph]) -> dict[str, Paragraph]:
    """从模板原段收集 main/sub2/sub3/body 四种编号样式原型。"""
    protos: dict[str, Paragraph] = {}
    for p in old_paras:
        if not is_actively_numbered(p):
            if "body" not in protos and (p.text or "").strip():
                protos["body"] = p
            continue
        nid = _num_id(p)
        t = (p.text or "").strip()
        if nid == "1" and "main" not in protos:
            protos["main"] = p
        elif nid == "4" and "main" not in protos:
            # 一般风险主序号常用 numId=4
            protos["main"] = p
        elif nid == "2" and "sub2" not in protos:
            protos["sub2"] = p
        elif nid == "3" and "sub3" not in protos:
            protos["sub3"] = p
        # 一般风险标题也记一份
        if nid == "4" and "general_main" not in protos:
            protos["general_main"] = p
        if len(t) <= 40 and "风险" in t and nid in ("1", "4") and "main" not in protos:
            protos["main"] = p
    if "body" not in protos and old_paras:
        protos["body"] = old_paras[0]
    if "main" not in protos:
        for p in old_paras:
            if is_actively_numbered(p):
                protos["main"] = p
                break
    if "main" not in protos and old_paras:
        protos["main"] = old_paras[0]
    # fallbacks
    protos.setdefault("sub2", protos["main"])
    protos.setdefault("sub3", protos.get("sub2", protos["main"]))
    protos.setdefault("general_main", protos["main"])
    return protos


def classify_risk_line(
    text: str,
    *,
    in_other_special: bool,
    after_burst_intro: bool,
    section: str,
) -> str:
    """判定段落应使用的编号样式：main / sub2 / sub3 / body。"""
    t = (text or "").strip()
    if not t:
        return "body"
    # 正文已自带（1）（7）① 等序号，不再套 Word 自动编号
    if re.match(r"^[（(]\d+[）)]", t):
        return "body"
    if t.startswith(("①", "②", "③", "④", "⑤", "⑥", "⑦", "⑧", "⑨", "⑩")):
        return "body"
    if t.startswith(("其中", "此外", "托管人关联方", "管理人的责任承担方式")):
        return "body"
    if t in ("其他特殊风险", "1.其他特殊风险") or t.replace(" ", "") == "其他特殊风险":
        return "main"
    # 突发偶然事件子项
    if after_burst_intro and len(t) <= 55 and t.endswith(("；", "。", ";")):
        return "sub3"
    # 其他特殊风险下的长条目
    if in_other_special:
        if t.startswith("突发偶然事件"):
            return "sub2"
        if len(t) > 35:
            return "sub2"
        if t.startswith(("计划管理人", "管理人操作", "因人为", "在现行", "金融市场")):
            return "sub2"
    # 短风险小标题 → 主序号
    if len(t) <= 40 and "风险" in t:
        # 排除以「本计划/本合同…」开头的叙述句，但保留「本金损失风险」这类标题
        if t.startswith(("本计划", "本合同", "本集合", "本资产管理", "在", "因", "若", "当", "对于", "金融", "投资于", "提示")):
            return "body"
        if t.startswith("本") and len(t) > 12:
            return "body"
        return "main"
    if len(t) <= 25 and t.endswith(("：", ":")) and "风险" in t:
        return "main"
    return "body"


def pick_prototype_by_content(
    lines: list[str],
    old_paras: list[Paragraph],
    *,
    section: str,
) -> list[Paragraph]:
    """按正文语义为每一行选择编号原型（不依赖与模板行数对齐）。"""
    protos = _collect_numbering_prototypes(old_paras)
    main_key = "general_main" if section == "general" else "main"
    result: list[Paragraph] = []
    in_other_special = False
    after_burst_intro = False
    for line in lines:
        t = (line or "").strip()
        if t in ("其他特殊风险", "1.其他特殊风险"):
            in_other_special = True
            after_burst_intro = False
        kind = classify_risk_line(
            line,
            in_other_special=in_other_special,
            after_burst_intro=after_burst_intro,
            section=section,
        )
        if t.startswith("其中") and "突发偶然" in t:
            after_burst_intro = True
        elif after_burst_intro and kind != "sub3":
            after_burst_intro = False
        if kind == "main":
            result.append(protos.get(main_key) or protos["main"])
        elif kind == "sub2":
            result.append(protos["sub2"])
        elif kind == "sub3":
            result.append(protos["sub3"])
        else:
            result.append(protos["body"])
    return result


def insert_paragraph_after(
    paragraph: Paragraph, text: str, *, prototype: Paragraph | None = None
) -> Paragraph:
    src = prototype._p if prototype is not None else paragraph._p
    new_p = copy.deepcopy(src)
    for child in list(new_p):
        if child.tag == qn("w:r"):
            new_p.remove(child)
    paragraph._p.addnext(new_p)
    para = Paragraph(new_p, paragraph._parent)
    para.add_run(text)
    return para


def replace_section_body(
    doc: Document, title_pred, end_pred, lines: list[str], *, section: str = "special"
) -> str:
    """替换某黄亮标题下、下一终止标题前的全部段落正文，并按内容挂回自动编号。"""
    paras = list(doc.paragraphs)
    start = find_para_index(paras, title_pred)
    if start is None:
        return "未找到章节标题"
    end = None
    for i in range(start + 1, len(paras)):
        if end_pred(paras[i]):
            end = i
            break
    if end is None:
        end = len(paras)

    to_delete = list(doc.paragraphs)[start + 1 : end]
    old_protos = [Paragraph(copy.deepcopy(p._p), p._parent) for p in to_delete]
    numbered_n = sum(1 for p in to_delete if is_actively_numbered(p))

    for p in to_delete:
        delete_paragraph(p)

    title_para = None
    for p in doc.paragraphs:
        if title_pred(p):
            title_para = p
            break
    if title_para is None:
        return "删除后丢失标题"

    cursor = title_para
    clean_lines = lines if lines else [""]
    if old_protos:
        chosen = pick_prototype_by_content(clean_lines, old_protos, section=section)
    else:
        chosen = [title_para] * len(clean_lines)
    for line, proto in zip(clean_lines, chosen):
        cursor = insert_paragraph_after(cursor, line, prototype=proto)
    return f"已替换正文 {len(clean_lines)} 段（按内容挂编号，原编号段 {numbered_n}）"


def split_lines(text) -> list[str]:
    if text is None:
        return []
    if isinstance(text, list):
        return [str(x) for x in text]
    return str(text).split("\n")


def fill_document(doc: Document, data: dict, *, clear_marks: bool = True) -> list[str]:
    report: list[str] = []
    product = str(data.get("产品简称") or "").strip()
    risk = str(data.get("风险等级") or "").strip()
    suit = str(data.get("适合投资者") or "").strip()
    refs = data.get("合同引用") or {}

    # 从模板推断旧值
    old_product = ""
    old_risk = ""
    old_suit = ""
    for p in doc.paragraphs:
        ys = yellow_texts(p)
        if not ys:
            continue
        if "集合资产管理计划" in p.text and not old_product:
            old_product = ys[0]
        if any("风险" in y for y in ys) and not old_risk:
            for y in ys:
                if "风险" in y:
                    old_risk = y
                    break
        if any("型" in y and ("C" in y or "普通投资者" in p.text) for y in ys) and not old_suit:
            for y in ys:
                if "型" in y or "C" in y:
                    old_suit = y
                    break

    # 封面产品简称（黄亮）
    if product:
        for p in doc.paragraphs[:5]:
            ys = yellow_texts(p)
            if ys and "集合资产管理计划" in p.text:
                # 替换黄亮 runs
                mapping = {ys[0]: product}
                if old_product:
                    mapping[old_product] = product
                n = replace_yellow_chunks(p, mapping)
                # 若黄亮已换但前后缀仍含旧名
                if old_product and old_product != product:
                    replace_in_runs(p, old_product, product)
                report.append(f"封面产品简称 → {product}（黄亮替换 {n}）")
                break

    # 风险等级 / 适合投资者（黄亮段落）
    for p in doc.paragraphs:
        ys = yellow_texts(p)
        if not ys:
            continue
        if "投资品种" in p.text or "适合合格投资者" in p.text:
            mapping = {}
            if risk and old_risk:
                mapping[old_risk] = risk
            elif risk and len(ys) >= 1 and "风险" in ys[0]:
                mapping[ys[0]] = risk
            if suit and old_suit:
                mapping[old_suit] = suit
            elif suit:
                for y in ys:
                    if "型" in y or "C" in y:
                        mapping[y] = suit
                        break
            n = replace_yellow_chunks(p, mapping)
            if old_risk and risk and old_risk != risk:
                replace_in_runs(p, old_risk, risk)
            if old_suit and suit and old_suit != suit:
                replace_in_runs(p, old_suit, suit)
            report.append(f"风险评级段落更新（黄亮 {n}）")

    # 全文替换旧产品名 / 风险等级（含一般风险里重复出现的评级句）
    if product and old_product and product != old_product:
        n = sum(replace_in_runs(p, old_product, product) for p in doc.paragraphs)
        report.append(f"全文替换产品简称（{n} 处）")
    if risk and old_risk and risk != old_risk:
        n = sum(replace_in_runs(p, old_risk, risk) for p in doc.paragraphs)
        report.append(f"全文替换风险等级（{n} 处）")
    if suit and old_suit and suit != old_suit:
        n = sum(replace_in_runs(p, old_suit, suit) for p in doc.paragraphs)
        report.append(f"全文替换适合投资者（{n} 处）")

    # 合同引用
    ref_map = {
        "投资者权利义务": None,
        "集合计划投资": None,
        "费用业绩报酬": None,
        "争议处理": None,
    }
    # 从模板黄亮收集旧引用并替换
    for p in doc.paragraphs:
        ys = yellow_texts(p)
        if not ys:
            continue
        text = p.text
        if "投资者的权利和义务" in text and refs.get("投资者权利义务"):
            replace_yellow_chunks(p, {ys[0]: str(refs["投资者权利义务"])})
            report.append(f"合同引用·投资者权利义务 → {refs['投资者权利义务']}")
        elif "集合计划的投资" in text and refs.get("集合计划投资"):
            replace_yellow_chunks(p, {ys[0]: str(refs["集合计划投资"])})
            report.append(f"合同引用·集合计划投资 → {refs['集合计划投资']}")
        elif "费用、业绩报酬" in text and refs.get("费用业绩报酬"):
            replace_yellow_chunks(p, {ys[0]: str(refs["费用业绩报酬"])})
            report.append(f"合同引用·费用业绩报酬 → {refs['费用业绩报酬']}")
        elif "争议处理" in text and refs.get("争议处理"):
            replace_yellow_chunks(p, {ys[0]: str(refs["争议处理"])})
            report.append(f"合同引用·争议处理 → {refs['争议处理']}")

    # 特殊风险 / 一般风险正文
    if "特殊风险揭示" in data:
        msg = replace_section_body(
            doc,
            title_pred=lambda p: "特殊风险揭示" in "".join(yellow_texts(p)) or (
                "特殊风险揭示" in p.text and p.text.strip().startswith("（一）")
            ),
            end_pred=lambda p: "一般风险揭示" in p.text and "（二）" in p.text,
            lines=split_lines(data["特殊风险揭示"]),
            section="special",
        )
        report.append(f"特殊风险揭示：{msg}")

    if "一般风险揭示" in data:
        msg = replace_section_body(
            doc,
            title_pred=lambda p: "一般风险揭示" in "".join(yellow_texts(p)) or (
                "一般风险揭示" in p.text and p.text.strip().startswith("（二）")
            ),
            end_pred=lambda p: p.text.strip().startswith("三、投资者声明"),
            lines=split_lines(data["一般风险揭示"]),
            section="general",
        )
        report.append(f"一般风险揭示：{msg}")

    if clear_marks:
        for p in doc.paragraphs:
            for run in p.runs:
                clear_highlight(run)
        report.append("已清除黄色高亮")

    return report


def extract_to_yaml(template_path: Path, out_path: Path) -> None:
    doc = Document(str(template_path))
    paras = list(doc.paragraphs)

    product = ""
    risk = ""
    suit = ""
    for p in paras[:10]:
        ys = yellow_texts(p)
        if not ys:
            continue
        if "集合资产管理计划" in p.text and not product:
            product = ys[0]
        if "投资品种" in p.text or "适合合格投资者" in p.text:
            for y in ys:
                if "风险" in y and not risk:
                    risk = y
                if ("型" in y or "C" in y) and not suit:
                    suit = y

    # section bodies
    special_start = general_start = declare_start = None
    for i, p in enumerate(paras):
        t = p.text.strip()
        ys = yellow_texts(p)
        if "特殊风险揭示" in t or "特殊风险揭示" in "".join(ys):
            special_start = i
        if "一般风险揭示" in t or "一般风险揭示" in "".join(ys):
            general_start = i
        if t.startswith("三、投资者声明"):
            declare_start = i

    special_lines = []
    general_lines = []
    if special_start is not None and general_start is not None:
        special_lines = [paras[i].text for i in range(special_start + 1, general_start)]
    if general_start is not None and declare_start is not None:
        general_lines = [paras[i].text for i in range(general_start + 1, declare_start)]
    while special_lines and not special_lines[-1].strip():
        special_lines.pop()
    while general_lines and not general_lines[-1].strip():
        general_lines.pop()

    refs = {}
    for p in paras:
        ys = yellow_texts(p)
        if not ys:
            continue
        if "投资者的权利和义务" in p.text:
            refs["投资者权利义务"] = ys[0]
        elif "集合计划的投资" in p.text:
            refs["集合计划投资"] = ys[0]
        elif "费用、业绩报酬" in p.text:
            refs["费用业绩报酬"] = ys[0]
        elif "争议处理" in p.text:
            refs["争议处理"] = ys[0]

    data = {
        "产品简称": product,
        "风险等级": risk,
        "适合投资者": suit,
        "合同引用": refs,
        "特殊风险揭示": "\n".join(special_lines),
        "一般风险揭示": "\n".join(general_lines),
    }

    def str_presenter(dumper, data):
        if isinstance(data, str) and ("\n" in data or len(data) > 60):
            return dumper.represent_scalar("tag:yaml.org,2002:str", data, style="|")
        return dumper.represent_scalar("tag:yaml.org,2002:str", data)

    yaml.add_representer(str, str_presenter)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as f:
        f.write("# 风险揭示书填报数据（黄亮字段/对应段落）\n\n")
        yaml.dump(data, f, allow_unicode=True, sort_keys=False, width=100)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="根据 YAML 自动生成集合计划风险揭示书")
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_gen = sub.add_parser("generate", help="用 YAML 填充模板生成风险揭示书")
    p_gen.add_argument("-d", "--data", required=True, type=Path)
    p_gen.add_argument("-t", "--template", type=Path, default=DEFAULT_TEMPLATE)
    p_gen.add_argument("-o", "--output", type=Path, default=None)
    p_gen.add_argument("--keep-highlight", action="store_true")

    p_ext = sub.add_parser("extract", help="从带高亮的模板反抽 YAML")
    p_ext.add_argument("-t", "--template", type=Path, default=DEFAULT_TEMPLATE)
    p_ext.add_argument("-o", "--output", type=Path, required=True)

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
    report = fill_document(doc, data, clear_marks=not args.keep_highlight)

    product = str(data.get("产品简称") or "风险揭示书").strip()
    out = args.output
    if out is None:
        DEFAULT_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        out = DEFAULT_OUTPUT_DIR / f"{product}集合资产管理计划风险揭示书-自动生成.docx"
    else:
        out.parent.mkdir(parents=True, exist_ok=True)

    doc.save(str(out))
    print(f"已生成: {out}")
    for line in report:
        print(line)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
