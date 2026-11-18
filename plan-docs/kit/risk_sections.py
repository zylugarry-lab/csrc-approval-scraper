#!/usr/bin/env python3
"""从管理合同「风险揭示」章节抽取风揭用的特殊/一般风险正文。"""

from __future__ import annotations

from pathlib import Path

from docx import Document


def _is_toc(t: str) -> bool:
    return ("\t" in t and t.strip()[-1:].isdigit()) or "HYPERLINK" in t or "PAGEREF" in t or "_Toc" in t


def load_paras(path: Path | str) -> list[str]:
    return [p.text for p in Document(str(path)).paragraphs]


def _find_exact(ps: list[str], name: str, start: int = 0) -> int | None:
    for i in range(start, len(ps)):
        tt = ps[i].strip()
        if _is_toc(tt):
            continue
        if tt == name:
            return i
    return None


def extract_风险揭示_sections(contract_path: Path | str) -> dict[str, str]:
    """返回 特殊风险揭示 / 一般风险揭示 正文（按合同，并把「其他特殊风险」并入特殊章节）。"""
    ps = load_paras(contract_path)

    # 优先：独立风揭附件体例 （一）特殊 / （二）一般
    i_special = _find_exact(ps, "（一）特殊风险揭示")
    i_general = _find_exact(ps, "（二）一般风险揭示")
    # 合同正文体例
    if i_special is None:
        i_special = _find_exact(ps, "特殊风险揭示") or _find_exact(ps, "一、特殊风险揭示")
    if i_general is None:
        i_general = _find_exact(ps, "一般风险揭示") or _find_exact(ps, "二、一般风险揭示")
    # 瑞策等：一、特定风险… / 二、…一般风险…
    if i_special is None or i_general is None:
        i_sp2 = i_gen2 = None
        for i, t in enumerate(ps):
            tt = t.strip()
            if _is_toc(tt):
                continue
            if i_sp2 is None and (
                tt.startswith("一、资产管理计划面临的特定风险")
                or tt.startswith("（一）特定投资方法")
            ):
                # 若是（一），往前找到「一、」
                if tt.startswith("（一）"):
                    for j in range(i, max(-1, i - 5), -1):
                        if ps[j].strip().startswith("一、资产管理计划面临的特定风险"):
                            i_sp2 = j
                            break
                    if i_sp2 is None:
                        i_sp2 = i
                else:
                    i_sp2 = i
            if i_gen2 is None and (
                tt.startswith("二、资产管理计划面临的一般风险")
                or tt == "二、一般风险揭示"
            ):
                i_gen2 = i
        if i_special is None:
            i_special = i_sp2
        if i_general is None:
            i_general = i_gen2

    if i_special is None or i_general is None:
        raise ValueError(f"未找到特殊/一般风险揭示标题: {contract_path}")
    if i_general <= i_special:
        # 附件风揭可能出现两次；取最后一对
        pairs = []
        for i, t in enumerate(ps):
            tt = t.strip()
            if tt in ("（一）特殊风险揭示", "特殊风险揭示", "一、特殊风险揭示"):
                for j in range(i + 1, len(ps)):
                    tj = ps[j].strip()
                    if tj in ("（二）一般风险揭示", "一般风险揭示", "二、一般风险揭示"):
                        pairs.append((i, j))
                        break
        if pairs:
            i_special, i_general = pairs[-1]

    # 结束：合同变更 / 违约 / 下一部分 / 投资者声明
    i_end = None
    for name in (
        "合同的变更、展期、终止与财产清算",
        "资产管理合同的变更、终止与财产清算",
        "三、投资者声明",
        "违约责任",
        "争议的处理",
    ):
        i_end = _find_exact(ps, name, start=i_general)
        if i_end is not None:
            break
    if i_end is None:
        for i in range(i_general + 1, len(ps)):
            tt = ps[i].strip()
            if _is_toc(tt):
                continue
            if tt.startswith("第") and "部分" in tt and "风险" not in tt:
                i_end = i
                break
            if tt.startswith("三、") and "风险" not in tt and len(tt) < 40:
                # 瑞策「三、其他风险」仍属风险揭示，不要截断；「合同的变更」已处理
                if "其他风险" in tt:
                    continue
                i_end = i
                break
    if i_end is None:
        i_end = len(ps)

    i_other_special = None
    for i in range(i_general, i_end):
        if ps[i].strip() in ("其他特殊风险", "1.其他特殊风险"):
            i_other_special = i
            break
    i_other_general = None
    for i in range(i_general, i_end):
        if ps[i].strip() in ("2.其他一般风险", "其他一般风险"):
            i_other_general = i
            break

    special_lines = [ps[i].rstrip() for i in range(i_special + 1, i_general)]
    if i_other_special is not None:
        end_os = i_other_general if i_other_general is not None else i_end
        special_lines += [""] + [ps[i].rstrip() for i in range(i_other_special, end_os)]

    general_lines: list[str] = []
    i = i_general + 1
    while i < i_end:
        if i_other_special is not None and i == i_other_special:
            i = i_other_general if i_other_general is not None else i_end
            continue
        general_lines.append(ps[i].rstrip())
        i += 1

    while special_lines and not special_lines[-1].strip():
        special_lines.pop()
    while general_lines and not general_lines[-1].strip():
        general_lines.pop()

    return {
        "特殊风险揭示": "\n".join(special_lines).strip(),
        "一般风险揭示": "\n".join(general_lines).strip(),
    }
