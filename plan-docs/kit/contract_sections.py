#!/usr/bin/env python3
"""从管理合同中抽取说明书用的整段条款（收益分配 / 利益冲突 / 信息披露）。"""

from __future__ import annotations

import re
from pathlib import Path

from docx import Document


def is_toc_line(t: str) -> bool:
    if any(x in t for x in ("HYPERLINK", "PAGEREF", "_Toc")):
        return True
    if "\t" in t and re.search(r"\t\s*\d+\s*$", t.strip()):
        return True
    return False


def load_paras(path: Path | str) -> list[str]:
    return [p.text for p in Document(str(path)).paragraphs]


def _find_exact(ps: list[str], name: str) -> int | None:
    for i, t in enumerate(ps):
        tt = t.strip()
        if is_toc_line(tt):
            continue
        if tt == name:
            return i
    return None


def _find_startswith(ps: list[str], prefixes: list[str]) -> int | None:
    for i, t in enumerate(ps):
        tt = t.strip()
        if is_toc_line(tt):
            continue
        for p in prefixes:
            if tt.startswith(p):
                return i
    return None


def _find_contains_heading(ps: list[str], needle: str) -> tuple[int | None, str | None]:
    for i, t in enumerate(ps):
        tt = t.strip()
        if is_toc_line(tt):
            continue
        if tt == needle:
            return i, None
        if needle in tt and len(tt) < 120:
            idx = tt.find(needle)
            prefix = tt[:idx].rstrip("。.;； ")
            return i, prefix if prefix else None
    return None, None


def _join_slice(ps: list[str], start: int, end: int, *, skip_title: str | None = None) -> str:
    lines = [ps[i].rstrip() for i in range(start, end)]
    if skip_title and lines and skip_title in lines[0] and len(lines[0]) < 40:
        lines = lines[1:]
        while lines and not lines[0].strip():
            lines.pop(0)
    while lines and not lines[-1].strip():
        lines.pop()
    return "\n".join(lines).strip()


def _is_conflict_end(tt: str) -> bool:
    if tt in ("投资经理的指定与变更", "资产管理计划的财产"):
        return True
    if tt.startswith("十五、投资经理") or tt.startswith("第十五"):
        return True
    if re.match(r"^第\s*[一二三四五六七八九十百零〇\d]+\s*部分", tt) and "利益冲突" not in tt:
        return True
    return False


def extract_说明书_sections(contract_path: Path | str) -> dict[str, str]:
    """返回说明书字段名 → 合同整段正文。"""
    ps = load_paras(contract_path)
    out: dict[str, str] = {}

    i_c = _find_exact(ps, "利益冲突及关联交易")
    if i_c is None:
        i_c = _find_startswith(ps, ["十四、利益冲突及关联交易"])
    if i_c is not None:
        end = len(ps)
        for i in range(i_c + 1, len(ps)):
            tt = ps[i].strip()
            if is_toc_line(tt):
                continue
            if _is_conflict_end(tt):
                end = i
                break
        text = _join_slice(ps, i_c, end, skip_title="利益冲突及关联交易")
        if text.startswith("十四、"):
            lines = text.split("\n")[1:]
            while lines and not lines[0].strip():
                lines.pop(0)
            text = "\n".join(lines).strip()
        out["利益冲突情况以及可能影响投资者合法权益的其他重要事项"] = text

    i_inc = _find_exact(ps, "收益的构成")
    if i_inc is None:
        i_inc = _find_startswith(
            ps,
            ["二十二、资产管理计划的收益分配", "（一）可供分配利润的构成", "（一）收益的构成"],
        )
    i_disc, disc_prefix = _find_contains_heading(ps, "信息披露与报告")
    if i_disc is None:
        i_disc = _find_startswith(ps, ["二十三、信息披露与报告"])
        disc_prefix = None

    if i_inc is not None:
        end = i_disc if i_disc is not None and i_disc > i_inc else len(ps)
        text = _join_slice(ps, i_inc, end)
        if "信息披露与报告" in text:
            text = text.split("信息披露与报告")[0].rstrip("。.;； \n")
        lines = text.split("\n")
        if lines and lines[0].startswith("二十二、"):
            lines = lines[1:]
            while lines and not lines[0].strip():
                lines.pop(0)
            text = "\n".join(lines).strip()
        out["收益分配"] = text

    if i_disc is not None:
        end = len(ps)
        for i in range(i_disc + 1, len(ps)):
            tt = ps[i].strip()
            if is_toc_line(tt):
                continue
            if tt in ("风险揭示",) or tt.startswith("二十四、风险揭示") or tt.startswith("风险揭示书"):
                end = i
                break
            if re.match(r"^第\s*[一二三四五六七八九十百零〇\d]+\s*部分", tt) and "信息披露" not in tt:
                end = i
                break
            if re.match(r"^二十[四五]、", tt):
                end = i
                break
        if disc_prefix is not None:
            text = _join_slice(ps, i_disc + 1, end)
        else:
            text = _join_slice(ps, i_disc, end, skip_title="信息披露与报告")
            if text.startswith("二十三、"):
                lines = text.split("\n")[1:]
                while lines and not lines[0].strip():
                    lines.pop(0)
                text = "\n".join(lines).strip()
        out["信息披露的内容、频率和方式"] = text

    return out
