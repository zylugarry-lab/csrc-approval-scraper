#!/usr/bin/env python3
"""从单一资产管理计划合同抽取说明书字段。"""

from __future__ import annotations

import re
from pathlib import Path

from docx import Document

from contract_sections import is_toc_line, load_paras

MANAGER = "上海国泰海通证券资产管理有限公司"

CN_NUM = "一二三四五六七八九十百零〇O○0-9"


def chapter_bounds(ps: list[str], start_key: str, end_keys: list[str]) -> tuple[int, int]:
    start = None
    for i, t in enumerate(ps):
        tt = t.strip()
        if is_toc_line(tt):
            continue
        if start is None and tt.replace(" ", "").startswith(start_key.replace(" ", "")):
            start = i
            continue
        if start is not None:
            compact = tt.replace(" ", "")
            if any(compact.startswith(k.replace(" ", "")) for k in end_keys):
                return start, i
    if start is None:
        return 0, 0
    return start, len(ps)


def join_slice(ps: list[str], start: int, end: int, skip_title: bool = True) -> str:
    lines = [ps[i].rstrip() for i in range(start, end)]
    if skip_title and lines:
        lines = lines[1:]
    while lines and not lines[0].strip():
        lines.pop(0)
    while lines and not lines[-1].strip():
        lines.pop()
    return "\n".join(lines).strip()


def numbered_blocks(text: str) -> dict[str, str]:
    parts = re.split(r"\n(?=[一二三四五六七八九十]{1,3}、)", "\n" + text)
    out = {}
    for part in parts:
        part = part.strip()
        m = re.match(r"^([一二三四五六七八九十]{1,3})、\s*(.*)$", part, re.S)
        if not m:
            continue
        title, body = m.group(2).split("\n", 1) if "\n" in m.group(2) else (m.group(2), "")
        out[title.strip()] = body.strip()
        out[m.group(1)] = part
    return out


def first_line(text: str) -> str:
    for line in text.splitlines():
        t = line.strip().strip("。.;；")
        if t:
            return t
    return ""


def field_custodian(ps: list[str]) -> str:
    for t in ps:
        tt = t.strip()
        if tt.startswith("托管人：") or tt.startswith("托管人:"):
            name = tt.split("：", 1)[-1].split(":", 1)[-1].strip()
            if name and "承诺" not in name:
                return name
        if "托管人（公章" in tt or "托管人(公章" in tt:
            name = tt.split("：", 1)[-1].split(":", 1)[-1].strip()
            if name:
                return name
    return ""


def field_manager(ps: list[str]) -> str:
    for t in ps[:80]:
        tt = t.strip()
        if tt.startswith("管理人：") or tt.startswith("管理人:"):
            name = tt.split("：", 1)[-1].split(":", 1)[-1].strip()
            if name:
                return name.replace("上海国泰君安证券资产管理有限公司", MANAGER)
    return MANAGER


def plan_name(ps: list[str], basic: str) -> str:
    m = re.search(r"(国泰[^\n]{0,40}单一资产管理计划)", basic)
    if m:
        return m.group(1).replace("。", "").strip()
    for t in ps[:30]:
        if "单一资产管理计划" in t and "合同" not in t and len(t) < 60:
            return t.strip().replace("。", "")
    return ""


def extract_basic(ps: list[str]) -> str:
    start, end = chapter_bounds(ps, "第五章", ["第六章", "第七章"])
    text = join_slice(ps, start, end)
    cut = re.search(r"\n一、因本资管计划|\n一、资产管理计划成立|\n一、参与和退出", text)
    if cut and cut.start() > 200:
        text = text[: cut.start()].strip()
    return text


def pick_block(blocks: dict[str, str], *needles: str) -> str:
    for key, val in blocks.items():
        for needle in needles:
            if needle in key:
                return val
    return ""


def extract_单一(contract_path: Path | str) -> dict:
    ps = load_paras(contract_path)
    basic = extract_basic(ps)
    blocks = numbered_blocks(basic)

    name = first_line(pick_block(blocks, "资产管理计划的名称", "名称")) or plan_name(ps, basic)
    raw_type = pick_block(blocks, "类别", "产品类型")
    type_m = re.search(r"(固定收益类|混合类|权益类|期货和衍生品类)", raw_type or basic)
    product_type = (type_m.group(1) if type_m else "").strip()
    type_src = raw_type or ""
    if re.search(r"基金中基金|FOF", type_src) and "非基金中基金" not in type_src and "非FOF" not in type_src:
        if product_type and "FOF" not in product_type:
            product_type = f"{product_type}、FOF类"

    duration = pick_block(blocks, "存续期限") or pick_block(blocks, "管理期限")
    duration = duration.split("\n")[0].strip() if duration else ""

    invest_start, invest_end = chapter_bounds(ps, "第十章", ["第十一章"])
    invest = join_slice(ps, invest_start, invest_end)
    invest_blocks = numbered_blocks(invest)

    scope = pick_block(invest_blocks, "投资范围") or pick_block(blocks, "主要投资方向", "投资方向")
    # drop 投资目标/比例 headings if mixed
    if "（一）投资范围" in scope:
        scope = scope.split("（一）投资范围", 1)[-1].strip()
    if "（二）" in scope and "投资比例" in scope:
        scope = scope.split("（二）")[0].strip()

    strategy = pick_block(invest_blocks, "投资策略")
    if not strategy:
        strategy = pick_block(invest_blocks, "投资管理的方法和标准")
    # 说明书常用决策依据开头
    if "（一）决策依据" in strategy:
        pass
    limits = pick_block(invest_blocks, "投资限制")
    prohibit = pick_block(invest_blocks, "投资禁止")
    if prohibit:
        limits = (limits + "\n\n投资禁止\n" + prohibit).strip() if limits else prohibit

    risk_line = ""
    m = re.search(r"风险评级评价为([^。\n]{2,40})", basic)
    if m:
        risk_line = m.group(0).strip()
    risk_code = ""
    risk_word = ""
    pair = re.search(r"(中低风险|中高风险|中风险|低风险|高风险)\s*[（(]\s*(R[1-5])\s*[）)]", basic)
    if pair:
        risk_word, risk_code = pair.group(1), pair.group(2)
    suit = ""
    sm = re.search(r"(C[1-5](?:[、,，]\s*C[1-5])*)", basic)
    if sm:
        suit = sm.group(1).replace(" ", "")

    sales = pick_block(blocks, "销售机构", "代销机构", "销售方式")
    sales_name = ""
    sm2 = re.search(r"销售机构[：:]\s*([^\n。]+)", sales or basic)
    if sm2:
        sales_name = sm2.group(1).strip()
    if not sales_name:
        sales_name = first_line(sales)

    advisor = "本单一资产管理计划未聘用投资顾问。"
    try:
        a0, a1 = chapter_bounds(ps, "第十一章", ["第十二章"])
        chap = join_slice(ps, a0, a1)
        if re.search(r"未聘用投资顾问|不聘请投资顾问|无投资顾问", chap):
            advisor = "本单一资产管理计划未聘用投资顾问。"
        elif "投资顾问" in chap and len(chap) > 40:
            advisor = chap[:800]
    except ValueError:
        pass

    try:
        r0, r1 = chapter_bounds(ps, "第四章", ["第五章"])
        rights = join_slice(ps, r0, r1)
        # keep 投资者权利义务
        mright = re.search(r"投资者享有如下权利[\s\S]{20,2500}?(?=\n二、管理人|\n二、托管人|$)", rights)
        rights_txt = mright.group(0).strip() if mright else rights[:2000]
    except ValueError:
        rights_txt = ""

    f0, f1 = chapter_bounds(ps, "第十九章", ["第二十章"])
    fees = join_slice(ps, f0, f1)
    if not fees:
        f0, f1 = chapter_bounds(ps, "第十六章", ["第十七章"])
        fees = join_slice(ps, f0, f1)
    if not fees:
        fees = pick_block(blocks, "费用")
    fee_kinds = fees
    # trim after 业绩报酬 details if too long for 费用种类
    if "二、" in fees and "管理费" in fees:
        fee_kinds = fees.split("二、")[0].strip()

    perf = pick_block(blocks, "费用")
    perf_txt = ""
    if re.search(r"不收取业绩报酬|不提取业绩报酬", basic + fees):
        perf_txt = "本计划不收取业绩报酬"
    else:
        pm = re.search(r"业绩报酬[\s\S]{0,800}", fees or perf or basic)
        perf_txt = pm.group(0).strip() if pm else (perf or "")

    subscribe = ""
    fee_basic = pick_block(blocks, "费用")
    for src in (fee_basic, fees):
        if not src:
            continue
        if re.search(r"认购费|参与费|退出费", src):
            chunk = src
            m3 = re.search(r"(认购费[\s\S]{0,600}?)(?:（二）|（三）|管理费|$)", chunk)
            subscribe = (m3.group(1) if m3 else chunk).strip()
            break
    if not subscribe:
        subscribe = "本计划不收取认购费、参与费和退出费。"

    try:
        d0, d1 = chapter_bounds(ps, "第二十章", ["第二十一章"])
        dist = join_slice(ps, d0, d1)
    except ValueError:
        dist = ""
    dist_short = (
        "投资者根据《管理合同》的规定参与或退出资管计划，取得受托资产的收益和清算后的剩余资产，"
        "包括但不限于投资收益（如有）等；投资者自行承担本资管计划的投资风险。"
    )

    try:
        i0, i1 = chapter_bounds(ps, "第二十一章", ["第二十二章"])
        disclose = join_slice(ps, i0, i1)
    except ValueError:
        disclose = ""

    try:
        c0, c1 = chapter_bounds(ps, "第十二章", ["第十三章"])
        conflict = join_slice(ps, c0, c1)
    except ValueError:
        conflict = ""

    try:
        k0, k1 = chapter_bounds(ps, "第六章", ["第七章"])
        raise_txt = join_slice(ps, k0, k1)
    except ValueError:
        raise_txt = ""
    if re.search(r"不设置募集期|不设初始募集期|不设募集期间", basic + raise_txt):
        raise_period = "本资管计划不设置募集期。"
    else:
        raise_period = first_line(raise_txt) or "本资管计划不设置募集期。"

    try:
        z0, z1 = chapter_bounds(ps, "第二十二章", ["第二十三章"])
        risk_body = join_slice(ps, z0, z1)
    except ValueError:
        risk_body = ""

    risk_reveal = (
        "管理人承诺以诚实信用，谨慎勤勉的原则管理和运用资管计划资产，管理人制定并执行相应的内部控制制度和风险管理制度，"
        "以降低风险发生的概率。但这些制度和方法不能完全防止风险出现的可能，管理人不保证本计划一定盈利，也不保证最低收益。"
        "本计划可能面临的风险包括但不限于本金损失风险、市场风险、流动性风险、管理风险、信用风险、投资标的风险、关联交易风险、"
        "操作风险、特定风险及其他风险。具体风险揭示内容详见《风险揭示书》。"
    )

    return {
        "产品简称": name,
        "年份": "二〇二六",
        "风险等级": f"{risk_word}（{risk_code}）" if risk_word and risk_code else (risk_word or ""),
        "适合投资者": suit,
        "特殊风险揭示": "",
        "一般风险揭示": "",
        "风险揭示全文": risk_body,
        "字段": {
            "名称": name,
            "产品类型": product_type,
            "管理期限": duration,
            "投资范围": scope,
            "投资策略": strategy,
            "投资限制和投资禁止": limits,
            "投资风险揭示": risk_reveal,
            "管理人概况": field_manager(ps),
            "托管人概况": field_custodian(ps),
            "销售机构": sales_name,
            "代销机构": sales_name,
            "聘用投资顾问情况": advisor,
            "收益分配和风险承担": dist_short if not dist else dist_short,
            "费用种类": fee_kinds,
            "业绩报酬": perf_txt,
            "认购费、参与费、退出费": subscribe,
            "参与费、退出费": subscribe,
            "投资者的重要权利和义务": rights_txt,
            "募集期间": raise_period,
            "信息披露的内容、方式和频率": disclose,
            "其他事项（利益冲突情况及其他重要事项）": conflict,
        },
        "合同引用": {},
    }


if __name__ == "__main__":
    import json
    import sys

    data = extract_单一(sys.argv[1])
    print(json.dumps({k: (v if k != "字段" else {kk: (vv[:120] + "…") if len(vv) > 120 else vv for kk, vv in v.items()}) for k, v in data.items() if k != "风险揭示全文"}, ensure_ascii=False, indent=2))
