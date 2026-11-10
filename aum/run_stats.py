#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
公募基金规模统计 — 自动化版（规则修订）

基于「公募基金规模统计_自动化.py」，仅调整分类规则以对齐
《公募基金非货分类规则》清单：
  - 灵活配置型全部归主动权益（不再按名称分流到固收+）
  - 指数关键词→量化仅限普通股票型/偏股混合型；平衡混合型除外
  - 名称含定开/定期开放的不因创业板等指数词进量化（清单标黄项）

用法:
  python 公募基金规模统计_自动化_规则修订.py
  python 公募基金规模统计_自动化_规则修订.py 2026Q2季报.xlsx 2026Q1季报.xlsx
  python 公募基金规模统计_自动化_规则修订.py --link-holdings ETF联接基金持仓占比.xlsx

每季准备:
  1) 本期+上期两份季报
  2) 建议放入「ETF联接基金持仓占比.xlsx」（含基金市值占净值比）；目录内会自动查找
     不放也能跑，回退约7%类型默认残差

输出: YYYYMMDD公募基金规模统计_自动化_规则修订.xlsx
依赖: pandas, openpyxl
"""

from __future__ import annotations

import argparse
import logging
import os
import re
import sys
from datetime import date
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
logger = logging.getLogger(__name__)

# ========== 常量 ==========
MY_COMPANY_FULL = "上海国泰海通证券资产管理有限公司"

CAT_PASSIVE_STOCK = "被动股票"
CAT_ACTIVE_EQUITY = "主动权益"
CAT_QUANT = "量化"
CAT_CREDIT = "固收"
CAT_CREDIT_PLUS = "固收+"
CAT_PASSIVE_CREDIT = "被动固收"
CAT_FOF = "FOF"
CAT_REITS = "REITs"
CAT_QDII = "QDII"
CAT_COMMODITY = "商品"
CAT_MONEY = "货币"
CAT_LINK = "ETF联接"
CAT_OTHER = "其他"
CAT_NON_MONETARY = "非货"

ALL_CATS = [CAT_PASSIVE_STOCK, CAT_ACTIVE_EQUITY, CAT_QUANT, CAT_CREDIT,
            CAT_CREDIT_PLUS, CAT_PASSIVE_CREDIT, CAT_FOF, CAT_REITS,
            CAT_QDII, CAT_COMMODITY]
RANK_CATS = [CAT_NON_MONETARY, CAT_ACTIVE_EQUITY, CAT_QUANT, CAT_PASSIVE_STOCK,
             CAT_CREDIT, CAT_CREDIT_PLUS, CAT_PASSIVE_CREDIT, CAT_FOF, CAT_REITS, CAT_QDII]

# 无持仓占比表时的回退：按底层类型穿透残差；0=整只剔除
ETF_LINK_RESIDUAL_RATIO = 0.07
LINK_RESIDUAL_BY_UNDERLYING = {
    "被动股票": 0.07, "量化": 0.07, "QDII": 0.06, "商品": 0.09,
    "固收": 0.07, "被动固收": 0.07, "固收+": 0.07, "FOF": 0.07,
    "REITs": 0.07, "主动权益": 0.07,
}


def _residual_from_pct(pct) -> float | None:
    """占比 → 残差比例。支持 0–100 百分数或 0–1 小数。"""
    if pd.isna(pct):
        return None
    p = float(pct)
    if p > 1.5:  # 百分数，如 94.11
        p = p / 100.0
    return max(0.0, min(1.0, 1.0 - p))


def extract_link_residuals_from_df(df: pd.DataFrame, scale_col: str | None = None) -> dict | None:
    """从季报内嵌穿透列提取联接残差 {6位代码: 残差比例}。

    优先「持有基金市值占比」→ 残差=1−占比；
    否则用「规模 − ETF联接持有的ETF规模」/规模。
    """
    cols = [str(c) for c in df.columns]
    pct_cols = [
        c for c in cols
        if "持有基金市值占比" in c or ("基金市值占" in c and "净值" in c)
    ]
    etf_cols = [c for c in cols if "联接持有的ETF" in c or "ETF联接持有" in c]
    if not pct_cols and not etf_cols:
        return None

    out: dict[str, float] = {}
    for _, r in df.iterrows():
        code = _fund_code6(r)
        if not code or code == "000000":
            continue
        residual = None
        if pct_cols:
            residual = _residual_from_pct(pd.to_numeric(r[pct_cols[0]], errors="coerce"))
        if residual is None and etf_cols and scale_col and scale_col in df.columns:
            etf = pd.to_numeric(r[etf_cols[0]], errors="coerce")
            sc = pd.to_numeric(r[scale_col], errors="coerce")
            if pd.notna(etf) and pd.notna(sc) and float(sc) > 0:
                residual = max(0.0, min(1.0, (float(sc) - float(etf)) / float(sc)))
        if residual is not None:
            out.setdefault(code, residual)

    if not out:
        return None
    src = pct_cols[0] if pct_cols else etf_cols[0]
    logger.info(f"季报穿透列「{src}」提取联接残差 {len(out)} 只")
    return out


def ensure_xlsx(path: str) -> None:
    """拒绝 Excel 97-2003 的 .xls。网页选择框虽可见，openpyxl 读不了这种 Windows 旧格式。"""
    with open(path, "rb") as f:
        magic = f.read(8)
    if magic.startswith(b"\xd0\xcf\x11\xe0"):
        raise ValueError("这是 Excel 97-2003 的 .xls，请在 Excel 中另存为“Excel 工作簿 (*.xlsx)”后再上传")
    if magic.startswith(b"PK"):
        return
    head = magic.lstrip().lower()
    if head.startswith(b"<") or head.startswith(b"<?xml"):
        raise ValueError("这不是 xlsx（内容像网页或旧版 XML）。请另存为“Excel 工作簿 (*.xlsx)”后再上传")


def _norm_col(name) -> str:
    """去掉 Windows Excel 单元格换行（\\r\\n）和不可见空格。"""
    s = str(name).replace("\r", "").replace("\n", "").replace("\xa0", "").replace("\u3000", "")
    return s.strip()


def _norm_code6(value) -> str:
    """证券代码 → 6 位。Excel 把 000216 存成数字 216.0 时，不能把小数点当成一位。"""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    if isinstance(value, bool):
        return ""
    if isinstance(value, (int, float)):
        if isinstance(value, float) and not value.is_integer():
            return ""
        text = str(int(value))
    else:
        text = str(value).strip().replace(".OF", "").replace(".SH", "").replace(".SZ", "")
        if re.fullmatch(r"\d+\.0+", text):
            text = text.split(".", 1)[0]
    digits = re.sub(r"\D", "", text)
    return digits.zfill(6)[-6:] if digits else ""


def _clean_header(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [_norm_col(c) for c in df.columns]
    return df


def _pct_columns(columns) -> list:
    return [c for c in columns if "基金市值占" in c or "占基金资产净值" in c or "持有基金市值占比" in c]


def _residual_columns(columns) -> list:
    return [c for c in columns if "残差比例" in str(c)]


def load_link_holdings_ratios(path: str) -> dict:
    """读取联接持仓占比或已算好的残差比例 → {6位代码: 残差比例}。

    残差比例 = 1 - 基金市值占净值比（未投向目标ETF/基金的部分）。
    重算工作簿首页常为「指标/值」汇总，明细在后续 sheet，这里会自动选有数据的那一页。
    占比缺失或>100%时做裁剪，避免负残差。
    """
    ensure_xlsx(path)
    xl = pd.ExcelFile(path)
    best = None
    best_n = -1
    first_cols = None
    for sheet in xl.sheet_names:
        df = _clean_header(pd.read_excel(xl, sheet_name=sheet))
        if first_cols is None:
            first_cols = list(df.columns)
        if "证券代码" not in df.columns:
            continue
        pct_cols = _pct_columns(df.columns)
        res_cols = _residual_columns(df.columns)
        if not pct_cols and not res_cols:
            continue
        use_res = bool(res_cols)
        col = res_cols[0] if use_res else pct_cols[0]
        n = int(pd.to_numeric(df[col], errors="coerce").notna().sum())
        if n > best_n:
            best = (df, col, use_res, sheet)
            best_n = n
    if best is None or best_n <= 0:
        raise ValueError(
            "持仓占比表未找到「基金市值占净值比」或「残差比例」列"
            f"（已查看 {xl.sheet_names}，首页列: {first_cols}）"
        )

    df, col, use_res, sheet = best
    out = {}
    for _, r in df.iterrows():
        code = _norm_code6(r["证券代码"])
        if not code or code == "000000":
            continue
        raw = pd.to_numeric(r[col], errors="coerce")
        if use_res:
            if pd.isna(raw):
                continue
            residual = float(raw)
            if residual > 1.5:
                residual = residual / 100.0
            residual = max(0.0, min(1.0, residual))
        else:
            residual = _residual_from_pct(raw)
        if residual is None:
            continue
        out.setdefault(code, residual)
    logger.info(f"已加载联接持仓占比: {path} sheet={sheet} 列={col}（{len(out)} 只）")
    return out


def find_link_holdings_file(directory: str) -> str | None:
    """在目录中自动查找联接持仓占比表。"""
    keywords = ("联接基金持仓占比", "持仓占比", "基金市值占")
    exclude = ("规模统计", "季报", "重算", "清单", "差额", "归因", "vs_", "VS")
    cands = []
    for f in os.listdir(directory):
        if not f.endswith(".xlsx") or f.startswith("~$"):
            continue
        if any(k in f for k in exclude):
            continue
        if any(k in f for k in keywords) or ("联接" in f and "占比" in f):
            cands.append(f)
    if not cands:
        return None
    cands.sort(key=lambda x: os.path.getmtime(os.path.join(directory, x)), reverse=True)
    path = os.path.join(directory, cands[0])
    logger.info(f"自动发现联接持仓占比表: {cands[0]}")
    return path


def _fund_code6(row) -> str:
    return _norm_code6(row.get("证券代码", ""))


# ========== 文件发现 ==========

def extract_date_key(filename: str) -> str:
    """从文件名中提取排序用的日期键。"""
    # 尝试匹配 YYYYQX 格式
    m = re.search(r"(\d{4})Q(\d)", filename)
    if m:
        return f"{m.group(1)}Q{m.group(2)}"
    # 尝试匹配 YYYYMMDD 格式
    m = re.search(r"(\d{4})(\d{2})(\d{2})", filename)
    if m:
        return f"{m.group(1)}Q{(int(m.group(2)) - 1) // 3 + 1}"
    return filename


def find_quarterly_files(directory: str) -> tuple[str, str]:
    """自动查找目录下最新的两个季报 Excel 文件。"""
    files = []
    for f in os.listdir(directory):
        if not f.endswith(".xlsx"):
            continue
        if "~$" in f:
            continue
        # 必须是季报文件（含"季报"或符合YYYYQX格式）
        if "季报" not in f and not re.search(r"\d{4}Q\d", f):
            continue
        # 排除已生成的统计文件
        if "公募基金规模统计" in f:
            continue
        files.append(f)

    if len(files) < 2:
        raise FileNotFoundError(
            f"在 {directory} 中只找到 {len(files)} 个季报文件，需要至少 2 个。\n"
            f"找到的文件: {files}\n"
            f"请确保文件名包含'季报'或'YYYYQX'格式。"
        )

    # 按日期键降序排序
    files.sort(key=lambda f: extract_date_key(f), reverse=True)
    logger.info("发现季报文件（按时间排序）:")
    for i, f in enumerate(files[:5], 1):
        marker = " <-- 将使用" if i <= 2 else ""
        logger.info(f"  {i}. {f}{marker}")

    q1_file = os.path.join(directory, files[0])
    q0_file = os.path.join(directory, files[1])
    return q1_file, q0_file


def find_scale_column(columns) -> str:
    """自动找到基金规模合计列。"""
    candidates = [c for c in columns if "基金规模合计" in c and "公告披露" in c]
    if not candidates:
        # 更宽松的匹配
        candidates = [c for c in columns if "规模" in c and "合计" in c]
    if not candidates:
        raise ValueError(f"未找到基金规模合计列，可用列: {columns[:20]}...")
    return candidates[0]


def find_data_sheet(path: str) -> str:
    """自动找到包含数据的工作表（不硬编码'万得'）。"""
    xl = pd.ExcelFile(path)
    sheet_names = xl.sheet_names

    # 优先匹配常见名称（不含空壳 Sheet1）
    for name in ["万得", "Wind", "数据", "基金列表"]:
        if name in sheet_names:
            logger.info(f"使用工作表: {name}")
            return name
    for name in sheet_names:
        if "季报" in name:
            logger.info(f"使用工作表: {name}")
            return name

    # 找含规模列/列数最多的 sheet
    best_sheet = sheet_names[0]
    best_score = -1
    for name in sheet_names:
        try:
            df = pd.read_excel(path, sheet_name=name, nrows=5)
            cols = [_norm_col(c) for c in df.columns]
            score = len(cols)
            if any("基金规模合计" in c for c in cols) or any("证券代码" in c for c in cols):
                score += 1000
            if score > best_score:
                best_score = score
                best_sheet = name
        except Exception:
            continue
    logger.info(f"使用工作表: {best_sheet}")
    return best_sheet


def parse_quarter_date(filename: str) -> date:
    """从文件名解析季报截止日期。"""
    # 尝试 YYYYQX 格式
    m = re.search(r"(\d{4})Q(\d)", filename)
    if m:
        year, q = int(m.group(1)), int(m.group(2))
        month_map = {1: 3, 2: 6, 3: 9, 4: 12}
        day_map = {1: 31, 2: 30, 3: 30, 4: 31}
        return date(year, month_map[q], day_map[q])

    # 尝试 YYYYMMDD 格式
    m = re.search(r"(\d{4})(\d{2})(\d{2})", filename)
    if m:
        return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))

    raise ValueError(f"无法从文件名解析日期: {filename}")


# ========== 分类逻辑 ==========

def classify_fund(row) -> str:
    """基金分类函数 — 基于 Wind 投资类型字段。"""
    l1 = str(row["投资类型(一级分类)"]) if pd.notna(row.get("投资类型(一级分类)")) else ""
    l2 = str(row["投资类型(二级分类)"]) if pd.notna(row.get("投资类型(二级分类)")) else ""
    l3 = str(row["投资类型[基金分类] 投资类型(三级分类)"]) if pd.notna(row.get("投资类型[基金分类] 投资类型(三级分类)")) else ""
    ff = str(row["基金全称"]) if pd.notna(row.get("基金全称")) else ""
    fs = str(row["证券简称"]) if pd.notna(row.get("证券简称")) else ""
    is_link = "联接" in ff or "联接" in fs
    is_quant = "量化" in ff or "量化" in fs

    # 国内公募REITs: 一级+二级分类均为REITs（三级分类为NaN）
    if "REITs" in l3 or (l1 == "REITs" and l2 == "REITs"):
        return CAT_REITS

    # 货币: 多级分类检查
    if (l3 == "货币市场型基金"
            or (l1 == "货币市场型基金" and l2 == "货币市场型基金")):
        return CAT_MONEY

    if is_link:
        return CAT_LINK
    if l3 in ["平衡混合型FOF基金", "债券型FOF基金", "偏债混合型FOF基金",
              "偏股混合型FOF基金", "目标日期型FOF基金", "股票型FOF基金"]:
        return CAT_FOF
    if "FOF" in l3:
        return CAT_FOF

    # QDII: 统一归为独立大类（九分类体系）
    qdii_types = [
        "国际(QDII)被动指数型股票基金",
        "国际(QDII)增强指数型股票基金",
        "国际(QDII)普通股票型基金",
        "国际(QDII)偏股混合型基金",
        "国际(QDII)灵活配置型基金",
        "国际(QDII)平衡混合型基金",
        "国际(QDII)宏观策略",
        "国际(QDII)普通债券型基金",
        "国际(QDII)偏债混合型基金",
        "国际(QDII)REITs",
    ]
    if l3 in qdii_types:
        return CAT_QDII

    if l3 == "被动指数型基金":
        return CAT_PASSIVE_STOCK
    if l3 == "增强指数型基金":
        return CAT_QUANT
    if l3 == "股票多空":
        return CAT_QUANT

    # 普通股票型/偏股混合型：名称含指数关键词 → 量化（规则表特殊量化归类）
    # 排除：定开/定期开放（如创业板两年定开，规则清单标黄，不应进量化）
    if l3 in ["普通股票型基金", "偏股混合型基金"]:
        idx_kw = ['中证1000', '中证500', '中证A500', '沪深300',
                  '中证全指', '上证科创板', '创业板', '科创板']
        is_dingkai = ("定开" in fs or "定开" in ff
                      or "定期开放" in fs or "定期开放" in ff)
        if (not is_dingkai) and any(k in fs or k in ff for k in idx_kw):
            return CAT_QUANT
        return CAT_QUANT if is_quant else CAT_ACTIVE_EQUITY

    # 平衡混合型：不按指数关键词进量化（规则表无此分流）
    if l3 == "平衡混合型基金":
        return CAT_QUANT if is_quant else CAT_ACTIVE_EQUITY

    # 灵活配置型：全部归主动权益（不再按名称关键词分流到固收+）
    # 仅名称含「量化」仍归量化（规则表量化选股型）
    if l3 == "灵活配置型基金":
        return CAT_QUANT if is_quant else CAT_ACTIVE_EQUITY

    if l3 in ["中长期纯债型基金", "短期纯债型基金"]:
        return CAT_CREDIT
    if l3 == "可转换债券型基金":
        return CAT_CREDIT_PLUS
    if l3 in ["被动指数型债券基金", "增强指数型债券基金"]:
        return CAT_PASSIVE_CREDIT
    if l3 in ["混合债券型一级基金", "混合债券型二级基金", "偏债混合型基金"]:
        return CAT_CREDIT_PLUS
    if "商品" in l3:
        return CAT_COMMODITY
    if "另类" in l3:
        return CAT_CREDIT
    return CAT_OTHER


def short_name(full_name: str) -> str:
    """公司全称 → 简称。"""
    m = {"上海国泰海通证券资产管理有限公司": "国泰海通",
         "上海国泰君安证券资产管理有限公司": "国泰君安资管"}
    if full_name in m:
        return m[full_name]
    return full_name.replace("基金管理有限公司", "基金").replace("管理股份有限公司", "")


def classify_link_underlying(row) -> str:
    l3 = ""
    if pd.notna(row.get("投资类型[基金分类] 投资类型(三级分类)")):
        l3 = str(row["投资类型[基金分类] 投资类型(三级分类)"])
    if "商品" in l3: return CAT_COMMODITY
    if "REITs" in l3: return CAT_REITS
    if "FOF" in l3: return CAT_FOF
    if "国际" in l3 or "QDII" in l3: return CAT_QDII
    if "债券" in l3:
        return CAT_PASSIVE_CREDIT if ("被动指数" in l3 or "增强指数" in l3) else CAT_CREDIT
    if "增强指数" in l3: return CAT_QUANT
    if "被动指数" in l3: return CAT_PASSIVE_STOCK
    return CAT_PASSIVE_STOCK


def process_quarter(
    df: pd.DataFrame,
    scale_col: str,
    link_residual_ratio: float | None = None,
    link_holdings: dict | None = None,
) -> dict:
    """汇总公司规模。联接残差优先级：
    1) link_residual_ratio 标量（含0=整只剔除）
    2) link_holdings 逐只残差（来自基金市值占净值比）
    3) 按底层类型默认残差
    """
    df = df.copy()
    df["目标分类"] = df.apply(classify_fund, axis=1)
    company_data = {}
    matched = missed = 0
    for _, r in df.iterrows():
        name, cat, scale = r["基金管理人"], r["目标分类"], r[scale_col]
        if pd.isna(name) or pd.isna(scale) or cat in [CAT_MONEY, CAT_OTHER]:
            continue
        scale = float(scale)
        if cat == CAT_LINK:
            u = classify_link_underlying(r)
            if link_residual_ratio is not None:
                ratio = float(link_residual_ratio)
            elif link_holdings:
                code = _fund_code6(r)
                if code in link_holdings:
                    ratio = float(link_holdings[code])
                    matched += 1
                else:
                    ratio = float(LINK_RESIDUAL_BY_UNDERLYING.get(u, ETF_LINK_RESIDUAL_RATIO))
                    missed += 1
            else:
                ratio = float(LINK_RESIDUAL_BY_UNDERLYING.get(u, ETF_LINK_RESIDUAL_RATIO))
            if ratio <= 0:
                continue
            cat, scale = u, scale * ratio
        if cat not in ALL_CATS:
            continue
        company_data.setdefault(name, {}).setdefault(cat, 0.0)
        company_data[name][cat] += scale
    for name in company_data:
        company_data[name][CAT_NON_MONETARY] = sum(company_data[name].get(c, 0) for c in ALL_CATS)
    if link_holdings is not None and link_residual_ratio is None:
        logger.info(f"联接残差: 持仓占比匹配 {matched} 只，回退默认比例 {missed} 只")
    return company_data


def calc_rankings(cq1: dict, cq0: dict) -> dict:
    """计算各类别排名。"""
    def _rank(cat):
        q1d = {n: d.get(cat, 0) for n, d in cq1.items()}
        q0d = {n: d.get(cat, 0) for n, d in cq0.items()}
        items = sorted([(n, v) for n, v in q1d.items() if v > 0], key=lambda x: -x[1])
        dr = pd.DataFrame(items, columns=["公司全称", "Q1规模"])
        dr["Q0规模"] = dr["公司全称"].map(q0d).fillna(0)
        dr["基金管理人"] = dr["公司全称"] if cat == CAT_NON_MONETARY else dr["公司全称"].apply(short_name)
        dr = dr.sort_values("Q1规模", ascending=False).reset_index(drop=True)
        dr["排名"] = range(1, len(dr) + 1)
        sc = "公司全称" if cat == CAT_NON_MONETARY else "基金管理人"
        q0_sorted = dr[dr["Q0规模"] > 0].sort_values("Q0规模", ascending=False)
        q0_map = dict(zip(q0_sorted[sc], range(1, len(q0_sorted) + 1)))
        dr["Q0排名"] = dr[sc].map(q0_map)
        # 上期无规模时 Q0排名 为 NaN，排名变化记为空（新进榜公司）
        dr["排名变化"] = dr["Q0排名"] - dr["排名"]
        dr["季度环比"] = dr.apply(
            lambda r: (r["Q1规模"] - r["Q0规模"]) / r["Q0规模"] if r["Q0规模"] else 0, axis=1)
        return dr[["排名", "基金管理人", "Q1规模", "Q0规模", "排名变化", "季度环比"]]

    return {cat: _rank(cat) for cat in RANK_CATS}


# ========== Excel生成 ==========

def generate_excel(
    q1_path: str,
    q0_path: str,
    output_dir: str | None = None,
    link_holdings_path: str | None = None,
) -> str:
    logger.info("=" * 60)
    logger.info("公募基金规模统计 — 自动化计算（规则修订）")
    logger.info("=" * 60)

    external_holdings = None
    ensure_xlsx(q1_path)
    ensure_xlsx(q0_path)
    if link_holdings_path:
        ensure_xlsx(link_holdings_path)
        external_holdings = load_link_holdings_ratios(link_holdings_path)

    q1_sheet = find_data_sheet(q1_path)
    q0_sheet = find_data_sheet(q0_path)

    # 读取
    df_q1 = pd.read_excel(q1_path, sheet_name=q1_sheet)
    df_q0 = pd.read_excel(q0_path, sheet_name=q0_sheet)
    df_q1 = df_q1.rename(columns=_norm_col)
    df_q0 = df_q0.rename(columns=_norm_col)

    scale_q1 = find_scale_column(df_q1.columns)
    scale_q0 = find_scale_column(df_q0.columns)
    logger.info(f"Q1规模列: {scale_q1}")
    logger.info(f"Q0规模列: {scale_q0}")

    # 联接残差：优先季报内嵌穿透列，其次外部持仓表，缺失回退类型默认
    def _merge_holdings(from_df, external):
        if from_df is None and external is None:
            return None
        out = dict(from_df or {})
        if external:
            for k, v in external.items():
                out.setdefault(k, v)
        return out

    holdings_q1 = _merge_holdings(extract_link_residuals_from_df(df_q1, scale_q1), external_holdings)
    holdings_q0 = _merge_holdings(extract_link_residuals_from_df(df_q0, scale_q0), external_holdings)
    if holdings_q1 is not None or holdings_q0 is not None:
        logger.info(
            "ETF联接口径: 逐只(1−持有基金市值占比/穿透)，缺失回退类型默认残差"
            f"（Q1={len(holdings_q1 or {})}只, Q0={len(holdings_q0 or {})}只）"
        )
    else:
        logger.info("ETF联接口径: 按底层类型默认残差（未提供穿透/持仓占比）")

    # 处理
    cq1 = process_quarter(df_q1, scale_q1, link_holdings=holdings_q1)
    cq0 = process_quarter(df_q0, scale_q0, link_holdings=holdings_q0)

    # 全市场汇总
    q1_sum = {c: sum(d.get(c, 0) for d in cq1.values()) for c in ALL_CATS + [CAT_NON_MONETARY]}
    q0_sum = {c: sum(d.get(c, 0) for d in cq0.values()) for c in ALL_CATS + [CAT_NON_MONETARY]}

    logger.info(f"Q1全市场非货合计: {q1_sum[CAT_NON_MONETARY]:.2f} 亿元")
    logger.info(f"Q0全市场非货合计: {q0_sum[CAT_NON_MONETARY]:.2f} 亿元")

    # 排名
    logger.info("计算排名...")
    rankings = calc_rankings(cq1, cq0)

    # 国泰海通数据
    gtht = {}
    for cat in RANK_CATS:
        m = "国泰海通" if cat == CAT_NON_MONETARY else "^国泰海通$"
        g = rankings[cat][rankings[cat]["基金管理人"].str.contains(m, na=False, regex=True)]
        if len(g) > 0:
            change = g.iloc[0]["排名变化"]
            gtht[cat] = {
                "排名": int(g.iloc[0]["排名"]),
                "Q1": float(g.iloc[0]["Q1规模"]),
                "Q0": float(g.iloc[0]["Q0规模"]),
                # 上期无排名时变化为空，避免 int(NaN) 报错
                "变化": int(change) if pd.notna(change) else None,
            }

    if gtht:
        logger.info("国泰海通数据:")
        for cat in RANK_CATS:
            d = gtht.get(cat)
            if d:
                logger.info(f"  {cat}: 排名={d['排名']}, Q1={d['Q1']:.4f}亿, Q0={d['Q0']:.4f}亿, 变化={d['变化']}")

    # 日期
    q1_date = parse_quarter_date(os.path.basename(q1_path))
    q0_date = parse_quarter_date(os.path.basename(q0_path))
    logger.info(f"Q1日期: {q1_date}")
    logger.info(f"Q0日期: {q0_date}")

    # 输出路径
    out_dir = output_dir or os.path.dirname(q1_path)
    os.makedirs(out_dir, exist_ok=True)
    output = os.path.join(out_dir, f"{q1_date.strftime('%Y%m%d')}公募基金规模统计_自动化_规则修订.xlsx")

    # ========== 生成Excel ==========
    wb = Workbook()
    hl_fill = PatternFill(start_color="FFE0E0", end_color="FFE0E0", fill_type="solid")
    hd_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    gt_font = Font(bold=True, color="C00000")
    tbrd = Border(
        left=Side(style="thin", color="D9D9D9"), right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"), bottom=Side(style="thin", color="D9D9D9"))
    nf = "#,##0.00"

    display_cats = ["非货合计（穿透剔除ETF联接）", "被动股票", "主动权益", "量化", "固收",
                    "固收+", "被动固收", "FOF", "REITs", "QDII"]
    key_map = {d: d.replace("合计（穿透剔除ETF联接）", "") for d in display_cats}
    key_map["非货合计（穿透剔除ETF联接）"] = CAT_NON_MONETARY

    # Sheet 1: 全市场及我司
    ws = wb.active
    ws.title = "全市场及我司规模增速对比"
    for ro, lb in [(0, "全市场规模总计（亿元）"), (12, "我司规模总计（亿元）"), (23, "我司规模排名及变化")]:
        ws.cell(row=ro + 1, column=1, value=lb)
        ws.cell(row=ro + 1, column=2, value=q1_date.strftime("%Y/%m/%d"))
        ws.cell(row=ro + 1, column=3, value=q0_date.strftime("%Y/%m/%d"))
        if ro < 23:
            ws.cell(row=ro + 1, column=5, value="规模增速" if ro == 12 else "全市场规模增速")
            ws.cell(row=ro + 1, column=6, value="季度环比")
        for i, dp in enumerate(display_cats):
            row = ro + 2 + i
            ky = key_map[dp]
            ws.cell(row=row, column=1, value=dp)
            if ro < 23:
                ws.cell(row=row, column=5, value=dp)
            if ro == 0:
                v1, v0 = q1_sum.get(ky, 0), q0_sum.get(ky, 0)
                ws.cell(row=row, column=2, value=round(v1, 2))
                ws.cell(row=row, column=3, value=round(v0, 2))
                ws.cell(row=row, column=6, value=round((v1 - v0) / v0, 4) if v0 else 0)
            elif ro == 12:
                d = gtht.get(ky, {})
                if d.get("Q1", 0) > 0:
                    ws.cell(row=row, column=2, value=round(d["Q1"], 2))
                    ws.cell(row=row, column=3, value=round(d["Q0"], 2))
                    ws.cell(row=row, column=6, value=round((d["Q1"] - d["Q0"]) / d["Q0"], 4) if d["Q0"] else 0)
            else:
                d = gtht.get(ky, {})
                if d.get("Q1", 0) > 0:
                    ws.cell(row=row, column=2, value=d["排名"])
                    ws.cell(row=row, column=3, value=d["变化"] if d["变化"] is not None else "")

    # Sheet 2: 各类基金规模
    ws2 = wb.create_sheet(title="各类基金规模")
    ws2.cell(row=1, column=1, value="规模总计（亿元）")
    ws2.cell(row=1, column=2, value=q1_date.strftime("%Y/%m/%d"))
    ws2.cell(row=1, column=3, value=q0_date.strftime("%Y/%m/%d"))
    for i, (c, dp) in enumerate(zip(
            [CAT_PASSIVE_STOCK, CAT_ACTIVE_EQUITY, CAT_QUANT, CAT_CREDIT,
             CAT_CREDIT_PLUS, CAT_PASSIVE_CREDIT, CAT_FOF, CAT_REITS],
            ["被动股票", "主动权益", "量化", "固收", "固收+", "被动固收", "FOF", "REITs"])):
        ws2.cell(row=i + 2, column=1, value=dp)
        ws2.cell(row=i + 2, column=2, value=round(q1_sum.get(c, 0), 2))
        ws2.cell(row=i + 2, column=3, value=round(q0_sum.get(c, 0), 2))

    # Sheet 3: 分类说明
    ws3 = wb.create_sheet(title="分类说明")
    for i, (cl, dc) in enumerate([
        ["分类", "包含的Wind三级分类及规则"],
        ["被动权益", "被动指数型基金（含ETF）；联接残差中底层为被动指数的部分"],
        ["主动权益", "普通股票型基金、偏股混合型基金、平衡混合型基金、灵活配置型基金（全部归主动权益，不再按名称分流）"],
        ["量化", "增强指数型基金、股票多空、名称含'量化'关键词、名称含指数关键词（中证1000/500/A500/沪深300/全指/科创板/创业板）的普通股票型/偏股混合型（定开除外）"],
        ["固收", "中长期纯债型基金、短期纯债型基金"],
        ["固收+", "混合债券型一级/二级基金、偏债混合型、可转债"],
        ["被动固收", "被动指数型债券基金（含ETF）、增强指数型债券基金；联接残差中底层为债券指数的部分"],
        ["FOF", "偏债混合型FOF、偏股混合型FOF、债券型FOF、目标日期型FOF、股票型FOF、平衡混合型FOF"],
        ["REITs", "基础设施公募REITs（一级/二级分类为REITs，三级分类可能为NaN）及QDII REITs"],
        ["", ""], ["排除项", ""],
        ["ETF联接", "名称含'联接'的基金：按(1−持有基金市值占比/季报穿透列)计入底层九分类；无占比时回退类型默认残差"],
        ["货币", "货币市场型基金（不计入非货）"],
        ["口径说明", "非货=九分类+联接残差（穿透剔除投向目标ETF的部分，避免与ETF双重计算）"],
        ["QDII", "所有QDII基金（排除在8分类外，但包含在非货合计中）"], ["商品", "商品型基金（排除）"],
    ]):
        ws3.cell(row=i + 1, column=1, value=cl)
        ws3.cell(row=i + 1, column=2, value=dc)

    # 排名 sheets
    sn_map = {CAT_NON_MONETARY: "非货（穿透剔除ETF联接）", CAT_PASSIVE_STOCK: "被动权益",
              CAT_ACTIVE_EQUITY: "主动权益", CAT_QUANT: "量化", CAT_CREDIT: "固收",
              CAT_CREDIT_PLUS: "固收+", CAT_PASSIVE_CREDIT: "被动固收", CAT_FOF: "FOF",
              CAT_REITS: "REITs", CAT_QDII: "QDII"}
    for cat in RANK_CATS:
        wr = wb.create_sheet(title=sn_map[cat])
        dr = rankings[cat]
        hs = ["排名", "基金管理人", q1_date.strftime("%Y/%m/%d"), q0_date.strftime("%Y/%m/%d"),
              "", "排名变化", "季度环比", "", "规模增速", "季度环比"]
        for c, h in enumerate(hs, 1):
            cell = wr.cell(row=1, column=c, value=h)
            cell.fill = hd_fill
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center")
            cell.border = tbrd
        for idx, rd in dr.iterrows():
            r = idx + 2
            is_my = "国泰海通" in str(rd["基金管理人"])
            wr.cell(row=r, column=1, value=int(rd["排名"]))
            wr.cell(row=r, column=2, value=rd["基金管理人"])
            wr.cell(row=r, column=3, value=round(rd["Q1规模"], 2))
            wr.cell(row=r, column=4, value=round(rd["Q0规模"], 2))
            rank_change_val = rd["排名变化"]
            wr.cell(row=r, column=6, value=int(rank_change_val) if pd.notna(rank_change_val) else "")
            wr.cell(row=r, column=7, value=rd["基金管理人"])
            wr.cell(row=r, column=9, value=rd["基金管理人"])
            wr.cell(row=r, column=10, value=round(rd["季度环比"], 4))
            wr.cell(row=r, column=3).number_format = nf
            wr.cell(row=r, column=4).number_format = nf
            if is_my:
                for c in range(1, 11):
                    cell = wr.cell(row=r, column=c)
                    cell.fill = hl_fill
                    cell.font = gt_font
                    cell.border = tbrd

    wb.save(output)
    logger.info(f"完成! Excel已保存: {output} ({len(wb.sheetnames)}个Sheet)")
    print(f"FILE={output}")
    return output


def default_output_directory() -> str:
    out = os.path.join(SCRIPT_DIR, "输出")
    Path(out).mkdir(parents=True, exist_ok=True)
    return out


# ========== 主入口 ==========

def main():
    parser = argparse.ArgumentParser(description="公募基金规模统计自动化（规则修订）")
    parser.add_argument("q1_file", nargs="?", help="当期季报 Excel 路径")
    parser.add_argument("q0_file", nargs="?", help="上期季报 Excel 路径")
    parser.add_argument("--input-dir", default=SCRIPT_DIR, help="自动查找季报时的目录")
    parser.add_argument("--output-dir", default=None, help="输出目录（默认：技能目录/输出）")
    parser.add_argument(
        "--link-holdings",
        help="ETF联接持仓占比表；不传则自动查找或用季报内穿透列",
    )
    args = parser.parse_args()

    try:
        if args.q1_file and args.q0_file:
            q1_path = args.q1_file if os.path.isabs(args.q1_file) else os.path.join(SCRIPT_DIR, args.q1_file)
            q0_path = args.q0_file if os.path.isabs(args.q0_file) else os.path.join(SCRIPT_DIR, args.q0_file)
            if not os.path.exists(q1_path):
                raise FileNotFoundError(f"找不到文件: {q1_path}")
            if not os.path.exists(q0_path):
                raise FileNotFoundError(f"找不到文件: {q0_path}")
            logger.info("手动模式: 使用指定文件")
            search_dir = os.path.dirname(q1_path) or SCRIPT_DIR
        else:
            logger.info(f"自动模式: 从 {args.input_dir} 查找季报文件")
            q1_path, q0_path = find_quarterly_files(args.input_dir)
            search_dir = args.input_dir

        holdings = args.link_holdings
        if holdings and not os.path.isabs(holdings):
            holdings = os.path.join(SCRIPT_DIR, holdings)
        if not holdings:
            holdings = find_link_holdings_file(search_dir)
            if not holdings:
                holdings = find_link_holdings_file(SCRIPT_DIR)

        out_dir = args.output_dir or default_output_directory()
        Path(out_dir).mkdir(parents=True, exist_ok=True)
        generate_excel(q1_path, q0_path, out_dir, link_holdings_path=holdings)

    except Exception as e:
        logger.error(f"错误: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
