"""从集合计划管理合同抽出说明书和风险揭示书要填的字段。"""

from __future__ import annotations

import re
import tempfile
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

from contract_sections import extract_说明书_sections, load_paras
from risk_sections import extract_风险揭示_sections

PLACEHOLDER = "【未从合同抽到，请对照合同填写】"

SMS_KEYS = [
    "名称",
    "类型",
    "目标规模",
    "管理期限",
    "募集期",
    "封闭期",
    "开放期",
    "份额面值",
    "最低金额",
    "相关费率",
    "投资范围",
    "投资管理的方法和标准",
    "投资限制",
    "风险收益特征",
    "适合推广对象",
    "管理人",
    "托管人",
    "销售机构",
    "初始募集",
    "认购的原则",
    "集合计划的参与·办理时间",
    "参与原则",
    "参与费",
    "集合计划的退出·办理时间",
    "退出原则",
    "退出费",
    "管理人自有资金参与情况",
    "集合计划的分级",
    "集合计划成立的条件",
    "集合计划设立失败",
    "集合计划份额转让",
    "费用种类",
    "不列入集合计划费用的项目",
    "业绩报酬",
    "收益分配",
    "集合计划展期",
    "终止",
    "信息披露的内容、频率和方式",
    "投资者权利与义务",
    "风险揭示",
    "利益冲突情况以及可能影响投资者合法权益的其他重要事项",
]

SINGLE_KEYS = [
    "名称",
    "产品类型",
    "管理期限",
    "投资范围",
    "投资策略",
    "投资限制和投资禁止",
    "投资风险揭示",
    "管理人概况",
    "托管人概况",
    "销售机构",
    "聘用投资顾问情况",
    "收益分配和风险承担",
    "费用种类",
    "业绩报酬",
    "认购费、参与费、退出费",
    "投资者的重要权利和义务",
    "募集期间",
    "信息披露的内容、方式和频率",
    "其他事项（利益冲突情况及其他重要事项）",
]


def _year_cn(year: int) -> str:
    digits = "〇一二三四五六七八九"
    return "".join(digits[int(c)] for c in str(year))


def accept_revisions(doc: Document) -> None:
    root = doc.element
    for node in list(root.iter(qn("w:del"))):
        parent = node.getparent()
        if parent is not None:
            parent.remove(node)
    for node in list(root.iter(qn("w:ins"))):
        parent = node.getparent()
        if parent is None:
            continue
        index = list(parent).index(node)
        for child in list(node):
            parent.insert(index, child)
            index += 1
        parent.remove(node)


def _node_text(element) -> str:
    parts: list[str] = []
    for node in element.iter():
        if node.tag.endswith("}tab"):
            parts.append("\t")
        elif node.tag.endswith("}t") and node.text:
            parts.append(node.text)
    return "".join(parts).strip()


def _is_toc(text: str) -> bool:
    if any(x in text for x in ("HYPERLINK", "PAGEREF", "_Toc")):
        return True
    if "\t" in text and re.search(r"\d+\s*$", text.strip()):
        return True
    return False


def load_lines(path: Path | str) -> tuple[list[str], list[tuple[int, str]]]:
    doc = Document(str(path))
    accept_revisions(doc)
    lines: list[str] = []
    body = doc.element.body
    for child in body:
        if child.tag == qn("w:p"):
            text = _node_text(child)
            if text:
                lines.append(text)
        elif child.tag == qn("w:tbl"):
            table = Table(child, doc)
            for row in table.rows:
                cells = []
                seen = set()
                for cell in row.cells:
                    if id(cell._tc) in seen:
                        continue
                    seen.add(id(cell._tc))
                    cell_text = " ".join(
                        _node_text(p._p) for p in cell.paragraphs if _node_text(p._p)
                    ).strip()
                    cells.append(cell_text)
                row_text = " | ".join(c for c in cells if c)
                if row_text:
                    lines.append(row_text)
    parts: list[tuple[int, str]] = []
    for text in lines:
        if not _is_toc(text):
            continue
        flat = re.sub(r"\s+", " ", text.replace("\t", " ")).strip()
        match = re.match(r"第\s*(\d+)\s*部分\s+(.+)$", flat)
        if not match:
            continue
        title = re.sub(r"\s+\d+$", "", match.group(2)).strip()
        if title and not title.isdigit():
            parts.append((int(match.group(1)), title))
    cleaned = []
    for ln in lines:
        if _is_toc(ln):
            continue
        cleaned.append(ln.replace("(", "（").replace(")", "）"))
    return cleaned, parts


def _find(lines: list[str], pred, start: int = 0, which: str = "first") -> int | None:
    found = None
    for i in range(start, len(lines)):
        if pred(lines[i].strip()):
            found = i
            if which == "first":
                return found
    return found


def _join(lines: list[str], start: int, stop, skip_head: bool = True) -> str:
    begin = start + 1 if skip_head else start
    out: list[str] = []
    i = begin
    while i < len(lines):
        text = lines[i].strip()
        if text and stop(text):
            break
        out.append(lines[i].rstrip())
        i += 1
    while out and not out[0].strip():
        out.pop(0)
    while out and not out[-1].strip():
        out.pop()
    return "\n".join(out).strip()


def _split_cn(text: str) -> list[tuple[str, str]]:
    items: list[tuple[str, str]] = []
    title = None
    buf: list[str] = []
    for line in text.split("\n"):
        stripped = line.strip()
        match = re.match(r"^([一二三四五六七八九十]+)、\s*(.*)$", stripped)
        if match and len(stripped) <= 80 and "。" not in match.group(1):
            if title is not None:
                items.append((title, "\n".join(buf).strip()))
            raw = match.group(2).strip()
            if "：" in raw:
                name, rest = raw.split("：", 1)
                title = name.strip()
                buf = [rest.strip()] if rest.strip() else []
            else:
                title = raw
                buf = []
        else:
            if title is not None:
                buf.append(line.rstrip())
    if title is not None:
        items.append((title, "\n".join(buf).strip()))
    return items


def _item(items: list[tuple[str, str]], *needles: str) -> str:
    for title, body in items:
        if any(n in title for n in needles):
            if body:
                return body
            return title
    return ""


def _subsplit(text: str) -> list[tuple[str, str]]:
    items: list[tuple[str, str]] = []
    title = None
    buf: list[str] = []
    for line in text.split("\n"):
        stripped = line.strip()
        match = re.match(r"^(\d+)、\s*(.*)$", stripped)
        if match and len(stripped) <= 40:
            if title is not None:
                items.append((title, "\n".join(buf).strip()))
            raw = match.group(2).strip()
            if "：" in raw:
                name, rest = raw.split("：", 1)
                title = name.strip()
                buf = [rest.strip()] if rest.strip() else []
            else:
                title = raw or match.group(1)
                buf = []
        else:
            if title is not None:
                buf.append(line.rstrip())
    if title is not None:
        items.append((title, "\n".join(buf).strip()))
    return items


def _part_ref(parts: list[tuple[int, str]], *needles: str) -> str:
    for num, title in parts:
        if any(n in title for n in needles):
            return f"第{num}部分“{title}”"
    return ""


def _short_name(full: str) -> str:
    name = full.strip().rstrip("。")
    name = re.sub(r"集合资产管理计划$", "", name).strip()
    return name


def _risk_level(text: str) -> str:
    match = re.search(
        r"((?:低|中低|中高|中|高)风险\s*[（(]\s*R[1-5]\s*[）)])",
        text,
    )
    if match:
        return re.sub(r"\s+", "", match.group(1)).replace("(", "（").replace(")", "）")
    match = re.search(r"(R[1-5]\s*[（(]\s*(?:低|中低|中高|中|高)风险\s*[）)])", text)
    if match:
        return re.sub(r"\s+", "", match.group(1)).replace("(", "（").replace(")", "）")
    return ""


def _suitable(text: str) -> str:
    match = re.search(
        r"((?:保守型|谨慎型|稳健型|平衡型|积极型|进取型|激进型)\s*[（(]\s*C[1-5]\s*[）)]"
        r"(?:\s*[、，,]\s*(?:保守型|谨慎型|稳健型|平衡型|积极型|进取型|激进型)\s*[（(]\s*C[1-5]\s*[）)])+)",
        text,
    )
    if match:
        return re.sub(r"\s+", "", match.group(1)).replace("(", "（").replace(")", "）")
    match = re.search(
        r"((?:C[1-5]\s*[（(]\s*(?:保守型|谨慎型|稳健型|平衡型|积极型|进取型|激进型)\s*[）)]"
        r"(?:\s*[、，,]\s*C[1-5]\s*[（(]\s*(?:保守型|谨慎型|稳健型|平衡型|积极型|进取型|激进型)\s*[）)])+))",
        text,
    )
    if match:
        return re.sub(r"\s+", "", match.group(1)).replace("(", "（").replace(")", "）")
    return ""


def _first_sentence_with(lines: list[str], needle: str) -> str:
    for line in lines:
        if needle in line and not _is_toc(line):
            return line.strip()
    return ""


def build_payload(contract_path: Path | str, year: int = 2026) -> dict:
    path = Path(contract_path)
    source = Document(str(path))
    accept_revisions(source)
    tmp = Path(tempfile.mkstemp(suffix=".docx")[1])
    source.save(tmp)
    path = tmp
    lines, parts = load_lines(path)
    notes: list[str] = []
    fields: dict[str, str] = {}

    name_i = _find(lines, lambda t: t.startswith("一、名称"))
    recruit_i = _find(
        lines,
        lambda t: t in ("资产管理计划的募集", "集合计划的募集") or t.startswith("一、募集对象"),
        start=(name_i or 0),
    )
    basic = ""
    if name_i is not None:
        end = recruit_i if recruit_i is not None else min(len(lines), name_i + 80)
        basic = "\n".join(lines[name_i:end])
    items = _split_cn(basic)

    full_name = ""
    name_item = _item(items, "名称")
    if name_item:
        full_name = name_item.strip().split("\n")[0].rstrip("。")
        if "：" in full_name and full_name.startswith("国泰") is False and len(full_name) < 8:
            full_name = name_item.strip().split("\n")[0]
        full_name = re.sub(r"^名称[：:]", "", full_name).strip().rstrip("。")
    if not full_name:
        for line in lines:
            match = re.search(r"(国泰海通.+?集合资产管理计划)", line)
            if match and "\t" not in line:
                full_name = match.group(1)
                break
    if full_name and not full_name.endswith("集合资产管理计划"):
        if "集合资产管理计划" not in full_name:
            full_name = full_name + "集合资产管理计划"
    short = _short_name(full_name) if full_name else path.stem
    short = re.sub(r"管理合同.*$", "", short).strip()
    short = re.sub(r"集合资产管理计划.*$", "", short).strip()

    type_text = _item(items, "类型")
    operate = _item(items, "运作方式")
    if type_text and operate and operate not in type_text:
        type_text = type_text.rstrip("。") + "。" + operate
    fields["名称"] = full_name
    fields["类型"] = type_text

    lock_open = _item(items, "封闭期")
    subs = _subsplit(lock_open)
    closed = ""
    opened = ""
    for title, body in subs:
        if "封闭" in title and "开放" not in title:
            closed = body or title
        elif "开放" in title:
            opened = body or title
    if not closed and lock_open:
        parts_lock = re.split(r"\n(?=\d+、\s*开放期)", lock_open, maxsplit=1)
        closed = parts_lock[0]
        if len(parts_lock) > 1:
            opened = re.sub(r"^\d+、\s*开放期[：:]?", "", parts_lock[1]).strip()
    # 开放期标题后的续段（临时开放期）并入
    if opened and lock_open:
        extra = []
        seen = False
        for line in lock_open.split("\n"):
            if "开放期" in line and re.match(r"^\d+、", line.strip()):
                seen = True
                continue
            if seen and line.strip() and not re.match(r"^[一二三四五六七八九十]+、", line.strip()):
                if line.strip() not in opened:
                    extra.append(line.strip())
        if extra:
            opened = opened.rstrip() + "\n" + "\n".join(extra)
    fields["封闭期"] = closed.strip()
    fields["开放期"] = opened.strip()
    fields["目标规模"] = _item(items, "目标规模")
    fields["管理期限"] = _item(items, "存续期限", "管理期限")

    invest = _item(items, "投资目标")
    invest_subs = _subsplit(invest)
    goal = ""
    scope = ""
    ratio = ""
    grade = ""
    for title, body in invest_subs:
        if "投资目标" in title:
            goal = body
        elif "投资方向" in title or "投资范围" in title:
            scope = body
        elif "比例" in title:
            ratio = body
        elif "风险" in title:
            grade = body
    if scope and ratio:
        scope = scope.rstrip() + "\n投资比例：\n" + ratio
    fields["投资管理的方法和标准"] = goal
    fields["投资范围"] = scope
    fields["风险收益特征"] = grade

    face = _item(items, "面值")
    fields["份额面值"] = face
    fields["募集期"] = ""

    fee_blob = _item(items, "费用")
    if not fee_blob:
        fee_i = _find(lines, lambda t: t in ("资产管理计划的费用", "集合计划的费用") or t.startswith("十、资产管理计划的费用") or t.startswith("十一、资产管理计划的费用") or t.startswith("十、集合计划的费用"))
        if fee_i is not None:
            fee_blob = _join(
                lines,
                fee_i,
                lambda t: t.startswith("十一、份额")
                or t.startswith("十二、")
                or "份额登记" in t
                or t in ("资产管理计划的募集", "集合计划的募集")
                or t.startswith("一、募集对象"),
            )
    fields["相关费率"] = fee_blob

    # 最低金额
    lowest = ""
    for src in (fields["目标规模"], "\n".join(lines)):
        for line in src.split("\n"):
            if any(k in line for k in ("首次参与", "首次认购", "最低金额")) and ("万元" in line or "元" in line):
                lowest = line.strip()
                break
        if lowest:
            break
    fields["最低金额"] = lowest

    # 当事人
    fields["管理人"] = ""
    fields["托管人"] = ""
    for line in lines:
        if line.startswith("管理人名称"):
            fields["管理人"] = line.split("：", 1)[-1].strip()
            break
    for line in lines:
        if line.startswith("托管人名称"):
            fields["托管人"] = line.split("：", 1)[-1].strip()
            break

    # 募集
    init_i = _find(lines, lambda t: t.startswith("三、初始募集期限") or t.startswith("初始募集期限") or "初始募集期" == t)
    if init_i is not None:
        fields["初始募集"] = _join(
            lines,
            init_i,
            lambda t: t.startswith("四、") or t.startswith("五、") or "认购的原则" in t or t.startswith("一、集合计划成立"),
        )
        first = fields["初始募集"].split("\n")[0].strip() if fields["初始募集"] else ""
        fields["募集期"] = first
    sub_i = _find(lines, lambda t: "认购的原则" in t or t.startswith("投资者认购原则") or t.startswith("1、投资者认购原则"))
    if sub_i is not None:
        fields["认购的原则"] = _join(
            lines,
            sub_i,
            lambda t: "认购的程序" in t or "认购程序" in t or t.startswith("五、") or t.startswith("2、认购程序"),
        )
    sale_i = _find(lines, lambda t: t.startswith("二、销售机构") or t.startswith("销售机构与销售方式"))
    if sale_i is not None:
        block = _join(lines, sale_i, lambda t: t.startswith("三、") or t.startswith("2、销售方式") or t.startswith("三、初始"))
        for line in block.split("\n"):
            if "销售机构" in line and ("直销" in line or "公司" in line):
                fields["销售机构"] = re.sub(r"^\d+、\s*销售机构[：:]?", "", line).strip()
                break
        if not fields.get("销售机构"):
            fields["销售机构"] = block
    if not fields.get("销售机构"):
        for line in lines:
            if "直销" in line and ("销售机构" in line or "推广机构" in line) and len(line) < 160:
                fields["销售机构"] = re.sub(r"^[\d、.\s]*销售机构[：:]?", "", line).strip()
                break
    if not fields.get("募集期"):
        for line in lines:
            if "初始募集期" in line and any(k in line for k in ("不超过", "自然日", "60")) and len(line) < 120:
                fields["募集期"] = line.strip()
                if not fields.get("初始募集"):
                    fields["初始募集"] = line.strip()
                break

    # 适合推广对象：常写在募集对象
    suit_para = ""
    for line in lines:
        if "适合向合格投资者" in line or ("风险承受能力" in line and "投资者" in line and len(line) > 20):
            suit_para = line.strip()
            break
    fields["适合推广对象"] = suit_para
    if not fields["风险收益特征"]:
        for line in lines:
            if "风险评级" in line and "R" in line:
                fields["风险收益特征"] = line.strip()
                break

    # 参与 / 退出
    def section_after(pred, stops) -> str:
        idx = _find(lines, pred)
        if idx is None:
            return ""
        return _join(lines, idx, stops)

    fields["集合计划的参与·办理时间"] = section_after(
        lambda t: "参与的办理时间" in t and len(t) < 30,
        lambda t: t.startswith("（二）") or t.startswith("二、"),
    )
    fields["参与原则"] = section_after(
        lambda t: "参与的原则" in t and len(t) < 30,
        lambda t: t.startswith("（三）") or t.startswith("（四）") or t.startswith("三、"),
    )
    fields["参与费"] = section_after(
        lambda t: "参与费" in t and "计算" in t and len(t) < 30,
        lambda t: t.startswith("二、集合计划的退出") or t.startswith("（一）退出") or t.startswith("二、"),
    )
    if not fields["参与费"]:
        for line in lines:
            if "参与费率" in line and len(line) < 40:
                fields["参与费"] = line.strip()
                break
    fields["集合计划的退出·办理时间"] = section_after(
        lambda t: t.startswith("（一）退出的办理时间") or t == "退出的办理时间",
        lambda t: t.startswith("（二）"),
    )
    fields["退出原则"] = section_after(
        lambda t: "退出的原则" in t and len(t) < 30,
        lambda t: t.startswith("（三）") or t.startswith("（四）"),
    )
    fields["退出费"] = section_after(
        lambda t: t.startswith("（四）退出费") or (t.startswith("1、退出费") and len(t) < 40),
        lambda t: t.startswith("（五）") or t.startswith("（六）") or t.startswith("2、退出金额"),
    )
    if not fields["退出费"]:
        for line in lines:
            if line.startswith("退出费率") or "退出费率：" in line or line.startswith("2、退出费率") or line.startswith("1、退出费率"):
                fields["退出费"] = line.strip()
                break

    fields["管理人自有资金参与情况"] = section_after(
        lambda t: "管理人自有资金参与" in t and len(t) < 30,
        lambda t: t.startswith("六、") or "份额持有人大会" in t or t in ("资产管理计划份额的登记", "集合计划的投资"),
    )
    fields["集合计划成立的条件"] = section_after(
        lambda t: "成立的条件" in t and len(t) < 30,
        lambda t: "设立失败" in t and len(t) < 30,
    )
    fields["集合计划设立失败"] = section_after(
        lambda t: "设立失败" in t and len(t) < 24,
        lambda t: t.startswith("三、") or "开始运作" in t,
    )
    fields["集合计划份额转让"] = section_after(
        lambda t: "份额的转让" in t and len(t) < 30,
        lambda t: t.startswith("四、") or "非交易过户" in t,
    )

    grade_i = _find(lines, lambda t: "分级" in t and len(t) < 24 and "不分级" not in t)
    if grade_i is not None:
        fields["集合计划的分级"] = _join(
            lines,
            grade_i,
            lambda t: t.startswith("第") and "部分" in t or ("利益冲突" in t and len(t) < 20) or t.startswith("十四、"),
        )
    if not fields.get("集合计划的分级"):
        for line in lines:
            if "不分级" in line and len(line) < 40:
                fields["集合计划的分级"] = line.strip()
                break

    fields["集合计划展期"] = section_after(
        lambda t: "集合计划的展期" in t and len(t) < 24,
        lambda t: ("终止" in t and "展期" not in t and len(t) < 24) or t.startswith("四、集合计划的终止"),
    )
    fields["终止"] = section_after(
        lambda t: (t.startswith("四、集合计划的终止") or t in ("集合计划的终止", "终止") or "应当终止" in t and len(t) < 40),
        lambda t: "清算" in t and len(t) < 24 and "终止" not in t,
    )

    limit_i = _find(lines, lambda t: "投资限制" in t and len(t) <= 20 and "禁止" not in t)
    if limit_i is not None:
        fields["投资限制"] = _join(
            lines,
            limit_i,
            lambda t: t.startswith("（") and "限制" not in t and len(t) < 24
            or t in ("投资策略", "投资禁止行为")
            or (t.startswith("五、") and len(t) < 24),
        )

    # 投资策略若比投资目标更完整，用作投资管理的方法和标准
    strat = section_after(
        lambda t: t in ("投资策略", "一、投资策略", "（一）投资策略"),
        lambda t: t.startswith("（二）") or t.startswith("二、") or t in ("投资限制", "投资禁止行为"),
    )
    if len(strat) > len(fields.get("投资管理的方法和标准") or ""):
        fields["投资管理的方法和标准"] = strat

    fee_part = _find(
        lines,
        lambda t: "费用" in t and "业绩报酬" in t and len(t) <= 24 and "。" not in t and "详见" not in t,
        which="last",
    )
    if fee_part is not None:
        fee_body = _join(
            lines,
            fee_part,
            lambda t: t in ("投资收益与分配", "收益分配") or t.startswith("二十二、") or "收益的构成" == t,
        )
        kind = ""
        for line in fee_body.split("\n"):
            if "费用种类" in line or line.startswith("一、费用的种类") or line.startswith("（一）费用"):
                kind = line
                break
        if "1、集合计划的托管费" in fee_body or "1、托管费" in fee_body:
            chunk = []
            started = False
            for line in fee_body.split("\n"):
                if re.match(r"^1、", line) and ("费" in line):
                    started = True
                if started and (line.startswith("二、") or line.startswith("（二）") or "计提" in line and len(line) < 20):
                    break
                if started:
                    chunk.append(line)
            if chunk:
                fields["费用种类"] = "\n".join(chunk).strip()
        elif kind:
            fields["费用种类"] = kind
        for i, line in enumerate(fee_body.split("\n")):
            if "不列入" in line and len(line) < 40:
                rest = fee_body.split("\n")[i + 1 :]
                taken = []
                for row in rest:
                    if re.match(r"^[一二三四五六七八九十]+、", row) or row.startswith("（"):
                        break
                    taken.append(row)
                fields["不列入集合计划费用的项目"] = "\n".join(taken).strip() or line
                break
        perf = ""
        perf_start = None
        fee_lines = fee_body.split("\n")
        for i, line in enumerate(fee_lines):
            is_heading = len(line) <= 30 and "。" not in line
            if is_heading and (
                ("提取业绩报酬" in line)
                or line.startswith("业绩报酬的计提")
                or ("计提方法" in line and "业绩报酬" in line)
                or (re.match(r"^[一二三四五六七八九十\d]+、", line) and line.endswith("业绩报酬"))
            ):
                perf_start = i
                break
        if perf_start is not None:
            perf = "\n".join(fee_lines[perf_start:]).strip()
        if len(perf) > 80:
            fields["业绩报酬"] = perf
    if len(fields.get("业绩报酬") or "") < 80 and fields.get("相关费率"):
        # 基本情况里的业绩报酬摘要
        for title, body in _subsplit(fields["相关费率"]):
            if "业绩报酬" in title and len(body) > len(fields.get("业绩报酬") or ""):
                fields["业绩报酬"] = body

    try:
        pulled = extract_说明书_sections(path)
    except Exception as exc:
        pulled = {}
        notes.append(f"长段落抽取失败：{exc}")
    for key, value in pulled.items():
        if value and len(value) >= len(fields.get(key) or ""):
            fields[key] = value
    if len(fields.get("收益分配") or "") < 80:
        income_i = _find(lines, lambda t: t in ("投资收益与分配", "收益分配") or t.endswith("投资收益与分配"))
        if income_i is not None:
            fields["收益分配"] = _join(
                lines,
                income_i,
                lambda t: ("信息披露" in t and len(t) < 24) or t.endswith("信息披露与报告"),
            )

    rights = section_after(
        lambda t: t.startswith("三、投资者的权利和义务") or t == "投资者的权利和义务",
        lambda t: t.startswith("四、管理人的权利"),
    )
    if rights:
        fields["投资者权利与义务"] = rights

    intro = section_after(
        lambda t: t in ("风险揭示", "二十四、风险揭示") or t.endswith("风险揭示") and len(t) < 16,
        lambda t: "特殊风险" in t or "特定风险" in t,
    )
    if intro and len(intro) > 30:
        fields["风险揭示"] = intro
    elif short:
        fields["风险揭示"] = f"详见《{short}集合资产管理计划风险揭示书》。"

    # 参与费句子整理
    if fields.get("参与费") and "参与" not in fields["参与费"][:8]:
        fields["参与费"] = fields["参与费"]

    risk_level = _risk_level("\n".join(lines))
    suitable = _suitable("\n".join(lines))
    if not risk_level and fields.get("风险收益特征"):
        risk_level = _risk_level(fields["风险收益特征"])
    if not suitable and fields.get("适合推广对象"):
        suitable = _suitable(fields["适合推广对象"])

    investor_part = ""
    for num, title in parts:
        if "当事人" in title or "权利义务" in title:
            investor_part = f"第{num}部分第三条“投资者的权利和义务”"
            break
    invest_ref = _part_ref(parts, "集合计划的投资")
    fee_ref = _part_ref(parts, "费用")
    dispute_ref = _part_ref(parts, "争议")

    try:
        risk_body = extract_风险揭示_sections(path)
    except Exception as exc:
        risk_body = {"特殊风险揭示": "", "一般风险揭示": ""}
        notes.append(f"风险揭示正文未按标题抽出：{exc}")

    missing = []
    for key in SMS_KEYS:
        if not str(fields.get(key) or "").strip():
            fields[key] = PLACEHOLDER
            missing.append(key)

    risk = {
        "产品简称": short,
        "风险等级": risk_level or PLACEHOLDER,
        "适合投资者": suitable or PLACEHOLDER,
        "合同引用": {
            "投资者权利义务": investor_part or PLACEHOLDER,
            "集合计划投资": invest_ref or PLACEHOLDER,
            "费用业绩报酬": fee_ref or PLACEHOLDER,
            "争议处理": dispute_ref or PLACEHOLDER,
        },
        "特殊风险揭示": risk_body.get("特殊风险揭示") or PLACEHOLDER,
        "一般风险揭示": risk_body.get("一般风险揭示") or PLACEHOLDER,
    }
    if risk["风险等级"] == PLACEHOLDER:
        missing.append("风险等级")
    if risk["适合投资者"] == PLACEHOLDER:
        missing.append("适合投资者")
    for label, value in risk["合同引用"].items():
        if value == PLACEHOLDER:
            missing.append("合同引用·" + label)
    if risk["特殊风险揭示"] == PLACEHOLDER:
        missing.append("特殊风险揭示")
    if risk["一般风险揭示"] == PLACEHOLDER:
        missing.append("一般风险揭示")

    data = {"产品简称": short, "年份": _year_cn(year), "字段": fields}
    return {"sms": data, "risk": risk, "missing": missing, "notes": notes}


def guess_kind(contract_path) -> str:
    head = "\n".join(load_paras(contract_path)[:60])
    if "单一资产管理计划" in head and "集合资产管理计划" not in head[:2000]:
        return "单一"
    return "集合"


def build_payload_single(contract_path, year: int = 2026) -> dict:
    from extract_单一 import extract_单一

    data = extract_单一(contract_path)
    fields = dict(data.get("字段") or {})
    missing = []
    for key in SINGLE_KEYS:
        if not str(fields.get(key) or "").strip():
            fields[key] = PLACEHOLDER
            missing.append(key)
    name = str(data.get("产品简称") or fields.get("名称") or "").strip()
    sms = {"产品简称": name, "年份": data.get("年份") or _year_cn(year), "字段": fields}
    risk_level = str(data.get("风险等级") or "").strip()
    suit = str(data.get("适合投资者") or "").strip()
    risk = {
        "产品简称": name,
        "风险等级": risk_level or PLACEHOLDER,
        "适合投资者": suit or PLACEHOLDER,
        "合同引用": {},
        "特殊风险揭示": data.get("特殊风险揭示") or "",
        "一般风险揭示": data.get("一般风险揭示") or "",
    }
    if risk["风险等级"] == PLACEHOLDER:
        missing.append("风险等级")
    if risk["适合投资者"] == PLACEHOLDER:
        missing.append("适合投资者")
    return {"sms": sms, "risk": risk, "missing": missing, "notes": ["已按单一计划 22 行表填写。"]}


def write_pair(contract_path, sms_template, risk_template, out_sms, out_risk, year: int = 2026, kind: str = "auto") -> str:
    from generate_说明书 import fill_document as fill_sms
    from generate_风险揭示书 import fill_document as fill_risk

    resolved = (kind or "auto").strip()
    if resolved == "auto":
        resolved = guess_kind(contract_path)
    if resolved == "单一":
        payload = build_payload_single(contract_path, year=year)
    else:
        payload = build_payload(contract_path, year=year)
    sms_doc = Document(str(sms_template))
    fill_sms(sms_doc, payload["sms"])
    sms_doc.save(str(out_sms))
    risk_doc = Document(str(risk_template))
    fill_risk(risk_doc, payload["risk"])
    risk_doc.save(str(out_risk))
    return f"类型：{resolved}\n" + summarize(payload)


def summarize(payload: dict) -> str:
    missing = payload.get("missing") or []
    lines = [
        f"产品：{payload['sms'].get('产品简称') or '未识别'}",
        f"未抽到 {len(missing)} 项" if missing else "合同字段已全部写入",
    ]
    for item in missing:
        lines.append(f"- {item}")
    for note in payload.get("notes") or []:
        lines.append(note)
    lines.append("生成的是初稿。业绩报酬公式、费率分档表请对照合同。")
    return "\n".join(lines)
