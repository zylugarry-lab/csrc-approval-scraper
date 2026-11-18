from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Optional

CN_CHARS = "一二三四五六七八九十零〇百千两"

# 全角数字/括号/句点归一；不要把半角 "." 换成顿号，否则 1.5亿元 会被当成列表
_FW = str.maketrans(
    "０１２３４５６７８９（）［］．",
    "0123456789()[].",
)

RE_CHAPTER = re.compile(r"^第[{0}\d]+[章节编部分条]".format(CN_CHARS))
RE_CN1 = re.compile(r"^[{0}]+、".format(CN_CHARS))
RE_CN2 = re.compile(r"^[（(][{0}]+[)）]".format(CN_CHARS))
RE_NUM_PAREN = re.compile(r"^[（(](\d{1,2})[)）]")
RE_NUM_DOT = re.compile(r"^\d+、|^\d+\.(?!\d)")
RE_CIRCLED = re.compile(r"^[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳]")
RE_LEAD = re.compile(r"^[\s\u3000\u00a0\u200b\t]+")
RE_COVER_YEAR = re.compile(
    r"^[一二三四五六七八九零〇O0-9]{2,12}年$"
)


@dataclass
class ParaKind:
    role: str
    list_key: Optional[str] = None
    nest_level: int = 0  # 0 = 本级顶格编号；1+ = 再缩进一级
    marker_len: int = 0


def normalize_for_match(text: str) -> str:
    text = text.replace("\u3000", " ").replace("\xa0", " ")
    text = RE_LEAD.sub("", text)
    return text.translate(_FW).strip()


def paragraph_text(p_el) -> str:
    from .xmlutil import W_T

    parts = []
    for node in p_el.iter(W_T):
        if node.text:
            parts.append(node.text)
        if node.tail:
            parts.append(node.tail)
    return "".join(parts)


def classify(text: str, in_table: bool) -> ParaKind:
    raw = text or ""
    stripped = normalize_for_match(raw)
    if not stripped:
        return ParaKind("empty")

    # 表格里的短标签（当事人、费用种类等）：顶格，不做列表悬挂
    if in_table and len(stripped) <= 20 and not _looks_like_list(stripped):
        if "。" not in stripped and "；" not in stripped:
            return ParaKind("label")

    if RE_CHAPTER.match(stripped):
        return ParaKind("heading", marker_len=_marker_len(RE_CHAPTER, stripped))
    if RE_CN1.match(stripped):
        return ParaKind("list", "cn1", 0, _marker_len(RE_CN1, stripped))
    if RE_CN2.match(stripped):
        return ParaKind("list", "cn2", 0, _marker_len(RE_CN2, stripped))
    if RE_NUM_PAREN.match(stripped):
        return ParaKind("list", "num_paren", 0, _marker_len(RE_NUM_PAREN, stripped))
    if RE_NUM_DOT.match(stripped):
        return ParaKind("list", "num_dot", 0, _marker_len(RE_NUM_DOT, stripped))
    if RE_CIRCLED.match(stripped):
        return ParaKind("list", "circled", 0, _marker_len(RE_CIRCLED, stripped))
    return ParaKind("body")


def _looks_like_list(text: str) -> bool:
    return bool(
        RE_CN1.match(text)
        or RE_CN2.match(text)
        or RE_NUM_PAREN.match(text)
        or RE_NUM_DOT.match(text)
        or RE_CIRCLED.match(text)
        or RE_CHAPTER.match(text)
    )


def is_cover_line(text: str) -> bool:
    """封面标题、年份：需要居中。"""
    t = re.sub(r"[\s\u3000]+", "", text or "")
    if not t:
        return False
    if RE_COVER_YEAR.match(t):
        return True
    if len(t) <= 80 and "。" not in t:
        if t.endswith("说明书") or t.endswith("风险揭示书"):
            return True
    return False


def _marker_len(regex: re.Pattern, text: str) -> int:
    m = regex.match(text)
    return len(m.group(0)) if m else 0
