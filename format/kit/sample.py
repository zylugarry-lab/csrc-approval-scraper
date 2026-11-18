from __future__ import annotations

"""生成一份「多方编辑后」的混乱样例，用来验证一键刷新。"""

from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor, Twips


def _font(run, east, ascii_name, size, color=None, highlight=None, bold=False):
    run.bold = bold
    run.font.size = Pt(size)
    run.font.name = ascii_name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), east)
    if color:
        run.font.color.rgb = RGBColor(*color)
    if highlight:
        run.font.highlight_color = highlight


def build_messy_sample(path: Path) -> Path:
    doc = Document()
    t = doc.add_paragraph()
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = t.add_run("国泰海通某某集合资产管理计划说明书")
    _font(r, "微软雅黑", "Calibri", 18, bold=True)

    p = doc.add_paragraph()
    p.paragraph_format.first_line_indent = Cm(0)
    r = p.add_run("信息披露的内容、频率和方式")
    _font(r, "黑体", "Arial", 14, (0, 0, 180), bold=True)

    p = doc.add_paragraph()
    r = p.add_run("一、定期报告")
    _font(r, "宋体", "Times New Roman", 12, bold=True)

    p = doc.add_paragraph()  # 缺首行缩进
    p.paragraph_format.first_line_indent = Cm(0)
    r = p.add_run("定期报告包括集合计划单位净值报告、集合计划的资产管理季度报告、年度报告等。")
    _font(r, "微软雅黑", "Calibri", 11, (30, 30, 30))
    r.font.highlight_color = WD_COLOR_INDEX.YELLOW

    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(1.2)
    r = p.add_run("1. 集合计划单位净值报告。")
    _font(r, "宋体", "Calibri", 10.5)

    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.4)
    r = p.add_run("2、集合计划的资产管理季度报告、年度报告。")
    _font(r, "楷体", "Times New Roman", 12, (200, 0, 0))

    for i, extra in enumerate([0.2, 0.85, 0.15, 1.1, 0.4, 0.6], start=2):
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(extra)
        p.paragraph_format.space_after = Pt(12 if i == 4 else 0)
        r = p.add_run("（{0}）本计划投资范围包括债券、基金、资产管理产品及其他金融工具。".format(i))
        _font(r, "宋体" if i % 2 else "微软雅黑", "Arial", 11 if i % 2 else 14)

    table = doc.add_table(rows=4, cols=2)
    table.style = "Table Grid"
    labels = ["投资范围与投资比例", "收益分配和风险承担", "费用种类", "其他事项（利益冲突情况及其他重要事项）"]
    contents = [
        [
            "投资范围：",
            "　　(1) 固定收益类资产：国债、地方政府债、央行票据、金融债、企业债、公司债、中期票据、短期融资券等。",
            "  (2) 现金管理类资产：银行存款、同业存单、货币市场基金等。",
        ],
        [
            "本计划收益分配原则如下。",
            "",
            "",
            "在符合相关规定的前提下，由管理人拟定分配方案。",
        ],
        [
            "1、管理费",
            "2、托管费",
            "3、业绩报酬",
        ],
        [
            "一、利益冲突",
            "（一）关联交易",
            "1. 本计划可能与管理人管理的其他产品发生交易，管理人将遵循公平原则处理。本句故意写得很长，用来检查换行后悬挂缩进是否对齐。",
            "2. 管理人应建立健全内部隔离制度。",
            "(1) 信息披露",
            "(2) 内部审批",
        ],
    ]
    for i, label in enumerate(labels):
        cell0 = table.cell(i, 0)
        cell0.text = ""
        p = cell0.paragraphs[0]
        r = p.add_run(label)
        _font(r, "黑体", "Calibri", 10, bold=True)
        cell1 = table.cell(i, 1)
        cell1.text = ""
        first = True
        for line in contents[i]:
            if first:
                para = cell1.paragraphs[0]
                first = False
            else:
                para = cell1.add_paragraph()
            para.paragraph_format.space_after = Pt(18 if line == "" else 0)
            para.paragraph_format.left_indent = Twips(200 * (i + 1))
            r = para.add_run(line)
            _font(
                r,
                "微软雅黑" if i else "宋体",
                "Calibri",
                9 + i,
                (0, 80, 0) if i == 1 else None,
                WD_COLOR_INDEX.TURQUOISE if i == 2 else None,
            )

    path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(path))
    return path


if __name__ == "__main__":
    out = Path(__file__).resolve().parent.parent / "samples" / "混乱样例.docx"
    build_messy_sample(out)
    print("已生成", out)
