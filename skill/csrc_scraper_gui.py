#!/usr/bin/env python3
# -*- coding: utf-8 -*-


import csv
import json
import logging
import os
import sys
import threading
import time
import re
from datetime import datetime, timedelta
from pathlib import Path
from typing import Dict, List, Optional, Tuple

try:
    import requests
except ImportError:
    print("=" * 55)
    print("  缺少必要的依赖库: requests")
    print("=" * 55)
    print()
    print("请先安装 requests，执行以下命令：")
    print()
    print("    pip3 install requests")
    print()
    print("如果权限不足，加上 --user 参数：")
    print()
    print("    pip3 install requests --user")
    print()
    print("=" * 55)
    sys.exit(1)

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill
except Exception:
    Workbook = None
OPENPYXL_AVAILABLE = Workbook is not None

# 可选：使用中国法定节假日/调休工作日计算（更准确）
try:
    from chinese_calendar import is_workday as _cn_is_workday  # type: ignore
except Exception:
    _cn_is_workday = None

DEFAULT_DATA_START_DATE = "2024-01-01"

# 日常一键勾选的对标机构（可改）
COMMON_BENCHMARKS = [
    "华夏", "易方达", "广发", "富国", "南方", "嘉实",
    "汇添富", "博时", "招商", "鹏华",
]

PREFS_FILENAME = "prefs.json"

_PROXY_ENV_KEYS = (
    "HTTP_PROXY", "HTTPS_PROXY", "http_proxy", "https_proxy",
    "ALL_PROXY", "all_proxy", "NO_PROXY", "no_proxy",
    "SOCKS_PROXY", "SOCKS5_PROXY", "socks_proxy", "socks5_proxy",
)


def _clear_process_proxy_env() -> None:
    """启动时清除进程内代理环境变量，避免走已关闭的本地代理。"""
    for key in _PROXY_ENV_KEYS:
        os.environ.pop(key, None)


def _format_network_error(exc: Exception) -> str:
    """将常见网络错误转为更易理解的提示。"""
    msg = str(exc)
    if isinstance(exc, requests.exceptions.ProxyError) or "ProxyError" in msg or "Unable to connect to proxy" in msg:
        return (
            "网络请求失败：检测到本地代理不可用（与 Wi-Fi 无关）。"
            "程序已配置为直连证监会官网，请关闭系统/终端里的 HTTP 代理后重试，"
            "或重新启动本程序。"
        )
    if "SSLError" in msg or "CERTIFICATE_VERIFY_FAILED" in msg:
        return "SSL 证书校验失败，请检查本机网络环境或系统时间是否正确。"
    return msg


_clear_process_proxy_env()

# ============================================================
# 核心爬虫逻辑（与命令行版本共享）
# ============================================================

BASE_URL = "https://neris.csrc.gov.cn/alappr-delare"
API_ENDPOINTS = {
    "progress": "/home/approval-progress/v1/list",
    "scrolling": "/home/approval-progress/v1/scrolling-list",
    "result": "/home/aprvResult/v1/list",
    "review": "/home/baprs-pub/v1/list",
    "stats": "/home/applyHome/getStatisticsCounts",
}

DEFAULT_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
    "Referer": "https://neris.csrc.gov.cn/alappr-delare-front/home/toPubFlow",
    "Origin": "https://neris.csrc.gov.cn",
}

REQUEST_DELAY = 0.15
MAX_RETRIES = 3
RETRY_DELAY = 2
PAGE_SIZE = 1000  # 接口实测支持；显著减少翻页次数

# 基金分类
FUND_CATEGORIES = ["货币型", "债券型", "混合型", "股票型", "REITs", "QDII", "ETF", "FOF"]

TASK_ORDER_MAP = {
    "接收材料": 1,
    "补正通知": 2,
    "受理通知": 3,
    "书面反馈": 4,  # 一次/二次/三次书面反馈统一为“书面反馈”
    "行政许可决定书": 5,
    "申请人主动撤销": 6,
}
ACCEPT_MATERIAL_STEP_NAMES = frozenset({"接收材料", "接受材料"})

REG_TYPE_APPLY = "申请注册"
REG_TYPE_CHANGE = "变更注册"

GTHT_KEYWORDS = [
    "国泰海通",
    "国泰君安",  # 历史/别名兜底
]

DEFAULT_BENCHMARK_INSTITUTIONS = {
    # 头部公募（关键词用于匹配 CSRC 标题里的申请单位全称）
    "华夏": ["华夏基金管理有限公司", "华夏基金"],
    "易方达": ["易方达基金管理有限公司", "易方达基金"],
    "广发": ["广发基金管理有限公司", "广发基金"],
    "富国": ["富国基金管理有限公司", "富国基金"],
    "汇添富": ["汇添富基金管理股份有限公司", "汇添富基金"],
    "博时": ["博时基金管理有限公司", "博时基金"],
    "招商": ["招商基金管理有限公司", "招商基金"],
    "鹏华": ["鹏华基金管理有限公司", "鹏华基金"],
    "南方": ["南方基金管理股份有限公司", "南方基金管理有限公司", "南方基金"],
    "嘉实": ["嘉实基金管理有限公司", "嘉实基金"],
    "国泰基金": ["国泰基金管理有限公司", "国泰基金"],
    "华安": ["华安基金管理有限公司", "华安基金"],
    "天弘": ["天弘基金管理有限公司", "天弘基金"],
    "银华": ["银华基金管理股份有限公司", "银华基金"],
    "平安": ["平安基金管理有限公司", "平安基金"],
    "中欧": ["中欧基金管理有限公司", "中欧基金"],
    "工银瑞信": ["工银瑞信基金管理有限公司", "工银瑞信"],
    "建信": ["建信基金管理有限责任公司", "建信基金"],
    "景顺长城": ["景顺长城基金管理有限公司", "景顺长城"],
    "永赢": ["永赢基金管理有限公司", "永赢基金"],
    "万家": ["万家基金管理有限公司", "万家基金"],
    "大成": ["大成基金管理有限公司", "大成基金"],
    "华泰柏瑞": ["华泰柏瑞基金管理有限公司", "华泰柏瑞"],
    "华宝": ["华宝基金管理有限公司", "华宝基金"],
    "中银": ["中银基金管理有限公司", "中银基金"],
    "创金合信": ["创金合信基金管理有限公司", "创金合信"],
    "兴业": ["兴业基金管理有限公司", "兴业基金"],
    "浦银安盛": ["浦银安盛基金管理有限公司", "浦银安盛"],
    "前海开源": ["前海开源基金管理有限公司", "前海开源"],
    "民生加银": ["民生加银基金管理有限公司", "民生加银"],
    "交银施罗德": ["交银施罗德基金管理有限公司", "交银施罗德"],
    "长城": ["长城基金管理有限公司", "长城基金"],
    "海富通": ["海富通基金管理有限公司", "海富通"],
    "兴证全球": ["兴证全球基金管理有限公司", "兴证全球", "兴全基金", "兴证全球基金"],
    "东方红": ["上海东方证券资产管理有限公司", "东方证券资产管理", "东方红资产管理", "东方红"],
    "信达澳亚": ["信达澳亚基金管理有限公司", "信达澳亚", "信澳"],
    "国寿安保": ["国寿安保基金管理有限公司", "国寿安保"],
    "中金": ["中金基金管理有限公司", "中金基金"],
    "国投瑞银": ["国投瑞银基金管理有限公司", "国投瑞银"],
    "银河": ["银河基金管理有限公司", "银河基金"],
    "长信": ["长信基金管理有限责任公司", "长信基金"],
    "财通资管": ["财通证券资产管理有限公司", "财通证券资产管理", "财通资管", "财通证券资管"],
    "财通基金": ["财通基金管理有限公司", "财通基金"],
    "中银国际": ["中银国际证券股份有限公司", "中银国际证券", "中银国际"],
    "华泰资管": [
        "华泰证券（上海）资产管理有限公司",
        "华泰证券(上海)资产管理有限公司",
        "华泰证券资产管理有限公司",
        "华泰证券资产管理",
    ],
    "东财基金": ["西藏东财基金管理有限公司", "东财基金"],
    "泰康": ["泰康资产管理有限责任公司", "泰康资产"],
    "人保资产": ["中国人保资产管理有限公司", "人保资产"],
}

# 对标筛选：机构简称易撞车时，若申请单位命中下列排除词则不算该机构
INSTITUTION_MATCH_EXCLUDES: Dict[str, List[str]] = {
    "中银": ["中银国际", "中银证券", "中银保险"],
    "国泰基金": ["国泰海通", "国泰君安", "海通证券", "海通资管", "国泰海通资管"],
    "华泰柏瑞": ["华泰证券", "华泰资管", "华泰紫金", "华泰证券资产管理"],
    "华泰资管": ["华泰柏瑞", "华泰柏瑞基金"],
    "财通基金": ["财通证券", "财通资管", "财通证券资管", "财通证券资产管理"],
    "财通资管": ["财通基金", "财通基金管理"],
    "平安": ["中国平安", "平安银行", "平安养老保险", "平安证券"],
    "兴业": ["兴业银行", "兴业证券", "兴业信托", "兴业消费金融"],
    "中金": ["中金公司", "中国国际金融", "中金财富"],
    "银河": ["银河证券", "中国银河证券"],
    "长城": ["长城证券", "长城国瑞"],
    "招商": ["招商证券", "招商局"],
    "中信": ["中信证券", "中信建投"],
}


def load_benchmark_institutions() -> Dict[str, List[str]]:
    """对标机构清单（内置在本文件，不读外部 json）。"""
    return dict(sorted(DEFAULT_BENCHMARK_INSTITUTIONS.items(), key=lambda kv: kv[0]))


# 与 csrc_scraper_gui.py 同目录：A 列填写要从「同批对比_全量」中排除的产品名（可写简称，子串匹配）
GROUP_COMPARE_EXCLUDE_FILENAME = "同批对比_排除产品.xlsx"
_GROUP_COMPARE_EXCLUDE_HEADER_HINTS = frozenset(
    {"产品名称", "产品名", "名称", "product", "product_name", "name"}
)

# 「ETF新主题」Sheet 排除产品（子串匹配产品名 / 原文标题）
ETF_NEW_THEME_EXCLUDE_PRODUCTS = frozenset(
    {
        "东方红中证东方红红利低波动交易型开放式指数证券投资基金",
        "东方红中证东方红红利低波动",
    }
)


def is_etf_new_theme_excluded(product_name: str, content: str = "") -> bool:
    """ETF新主题：命中排除名单则不输出。"""
    hay = f"{product_name or ''} {content or ''}"
    if not hay.strip():
        return False
    for name in ETF_NEW_THEME_EXCLUDE_PRODUCTS:
        if name and name in hay:
            return True
    return False


def script_directory() -> str:
    return os.path.dirname(os.path.abspath(__file__))


def default_output_directory() -> str:
    """技能目录下的「输出」；不存在则回退脚本目录。"""
    out = os.path.join(script_directory(), "输出")
    try:
        Path(out).mkdir(parents=True, exist_ok=True)
        return out
    except Exception:
        return script_directory()


def prefs_path() -> str:
    return os.path.join(script_directory(), PREFS_FILENAME)


def load_prefs() -> Dict:
    path = prefs_path()
    if not os.path.exists(path):
        return {}
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def save_prefs(data: Dict) -> None:
    try:
        with open(prefs_path(), "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def group_compare_exclude_products_path() -> str:
    return os.path.join(script_directory(), GROUP_COMPARE_EXCLUDE_FILENAME)


def export_output_base_name() -> str:
    """导出文件名：证监会网页数据 + 运行当日日期（YYYYMMDD）。"""
    return f"证监会网页数据{datetime.now().strftime('%Y%m%d')}"


def export_benchmark_xlsx_path(main_xlsx_path: str) -> str:
    """对标同业 Excel 路径：主文件名 + _对标同业。"""
    stem, ext = os.path.splitext(main_xlsx_path)
    return f"{stem}_对标同业{ext or '.xlsx'}"


def _write_benchmark_workbook(
    xlsx_path: str,
    *,
    benchmark_data: List[Dict],
    selected_benchmarks: List[str],
    summary_fields: List[str],
) -> Dict[str, int]:
    """将对标机构各 Sheet 写入独立 Excel。"""
    benchmark_summary_rows = [summarize_progress(r) for r in benchmark_data]
    sort_summary_rows_by_date_desc(benchmark_summary_rows)
    inst_map = load_benchmark_institutions()

    wb = Workbook()
    counts: Dict[str, int] = {}
    first_sheet = True
    for inst in selected_benchmarks:
        keywords = inst_map.get(inst) or [inst]
        inst_rows = [
            r
            for r in benchmark_summary_rows
            if match_institution(r.get("content", ""), keywords, inst_key=inst)
        ]
        counts[inst] = len(inst_rows)
        inst_rows = group_with_blank_rows(
            inst_rows, ["reg_type", "fund_type"], within_group_date_desc=True
        )
        sheet_name = f"对标_{inst}"[:31]
        if first_sheet:
            ws = wb.active
            ws.title = sheet_name
            first_sheet = False
        else:
            ws = wb.create_sheet(sheet_name)
        ws.append(summary_fields)
        for r in inst_rows:
            ws.append([r.get(k, "") for k in summary_fields])

    Path(xlsx_path).parent.mkdir(parents=True, exist_ok=True)
    wb.save(xlsx_path)
    return counts


def ensure_group_compare_exclude_template() -> str:
    """若排除名单 Excel 不存在则创建带表头的模板，返回文件路径。"""
    path = group_compare_exclude_products_path()
    if os.path.exists(path):
        return path
    if not OPENPYXL_AVAILABLE:
        return path
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "排除产品"
    ws.append(["产品名称"])
    ws.append(["示例-请删除本行，下面填写要排除的产品名（支持简称）"])
    wb.save(path)
    return path


def load_group_compare_exclude_products() -> set:
    """读取同目录排除名单 Excel 第一列（跳过表头行）。"""
    path = group_compare_exclude_products_path()
    if not os.path.exists(path) or not OPENPYXL_AVAILABLE:
        return set()
    try:
        from openpyxl import load_workbook

        wb = load_workbook(path, read_only=True, data_only=True)
        ws = wb.active
        names: set = set()
        for row in ws.iter_rows(min_row=1, max_col=1, values_only=True):
            val = row[0] if row else None
            if val is None:
                continue
            s = str(val).strip()
            if not s:
                continue
            if s in _GROUP_COMPARE_EXCLUDE_HEADER_HINTS:
                continue
            if s.startswith("示例"):
                continue
            names.add(s)
        wb.close()
        return names
    except Exception:
        return set()


def is_group_compare_excluded_content(content: str, exclude_names: set) -> bool:
    """命中排除名单则该行不出现在「同批对比_全量」。"""
    if not exclude_names:
        return False
    product = extract_product_name_from_content(content)
    haystack = f"{product} {content or ''}".strip()
    for name in exclude_names:
        token = (name or "").strip()
        if not token:
            continue
        if token in haystack or (product and product in token):
            return True
    return False


def classify_fund(name: str) -> Optional[str]:
    """根据产品名称自动分类基金类型，非基金产品返回 None。"""
    name = name.replace("公开募集基金募集申请注册-", "").replace("公开募集基金变更注册-", "")

    non_fund_keywords = [
        "向特定对象发行股票", "公开发行股票", "公司债券",
        "合格境外投资者资格", "合格境内机构投资者资格",
        "公募基金管理公司变更", "公募基金管理公司申请",
        "从事公募基金销售的机构", "期货公司", "证券公司",
    ]
    if any(k in name for k in non_fund_keywords):
        return None

    if "交易型开放式" in name and "联接" not in name:
        return "ETF"
    if any(k in name for k in ["不动产投资信托", "基础设施", "REITs"]):
        return "REITs"
    if "QDII" in name:
        return "QDII"
    if any(k in name for k in ["基金中基金", "FOF"]):
        return "FOF"
    if any(k in name for k in ["货币", "同业存单"]):
        return "货币型"
    if any(k in name for k in ["债券", "纯债", "短债", "中短债", "可转债", "可转换债券"]):
        return "债券型"
    if any(k in name for k in ["混合", "灵活配置"]):
        return "混合型"
    if any(k in name for k in ["股票", "指数增强", "增强型"]):
        return "股票型"
    if "联接" in name:
        return "股票型"
    if "指数" in name and "债券" not in name:
        return "股票型"

    if "基金" in name:
        if any(k in name for k in ["亚洲", "全球", "海外", "港元", "美元"]):
            return "QDII"
        return "其他"

    return None


def classify_reg_type(text: str) -> str:
    t = (text or "").strip()
    # 页面标题里常见的两种写法，但有时分隔符可能不是半角“-”
    if "公开募集基金变更注册" in t or "变更注册" in t:
        return REG_TYPE_CHANGE
    if "公开募集基金募集申请注册" in t or "募集申请注册" in t or "申请注册" in t:
        return REG_TYPE_APPLY
    return ""


def normalize_task_name(name: str) -> str:
    n = (name or "").strip()
    if not n:
        return ""
    if "书面反馈" in n:
        return "书面反馈"
    return n


def task_order(task_name: str) -> Optional[int]:
    return TASK_ORDER_MAP.get(normalize_task_name(task_name))


def record_task_order_num(row: Dict) -> int:
    try:
        return int(str(row.get("task_order") or "").strip())
    except Exception:
        return 0


def extract_accept_material_date(row: Dict) -> str:
    """从 progress 或 all_task_names / all_finish_dates 中取「接收材料」完成日期。"""
    progress = row.get("progress")
    if progress:
        for p in progress:
            n = normalize_task_name(str(p.get("task_name") or ""))
            if n in ACCEPT_MATERIAL_STEP_NAMES:
                return (p.get("finish_date") or "").strip()
        return ""

    names = [x.strip() for x in str(row.get("all_task_names") or "").split(" > ")]
    dates = [x.strip() for x in str(row.get("all_finish_dates") or "").split(" > ")]
    for i, n in enumerate(names):
        if normalize_task_name(n) in ACCEPT_MATERIAL_STEP_NAMES and i < len(dates):
            return dates[i]
    return ""


def build_latest_submission_step1_row(latest: Dict) -> Optional[Dict]:
    """最后一次提交若已离开接收材料，补一行仅展示该次「接收材料」步骤（不含中间过程）。"""
    if record_task_order_num(latest) <= 1:
        return None
    accept_date = extract_accept_material_date(latest)
    if not accept_date:
        return None
    rr = dict(latest)
    rr["task_name"] = "接收材料"
    rr["task_order"] = 1
    rr["finish_date"] = accept_date
    rr["accept_material_date"] = accept_date
    return rr


def extract_applicant_from_title(title: str) -> str:
    """
    从 showCntnt / matter_name 里提取“关于XXX的《...》”中的 XXX。
    提取失败则返回空字符串。
    """
    t = (title or "").strip()
    if not t:
        return ""
    m = re.search(r"关于(.+?)的《", t)
    if not m:
        return ""
    return m.group(1).strip()


def extract_product_name_from_content(content: str) -> str:
    """
    从标题中抽取《...》里的产品名称部分。
    常见格式：
      关于XXX的《公开募集基金募集申请注册-产品名》
    返回尽量干净的“产品名”（去掉前缀说明）。
    """
    t = (content or "").strip()
    if not t:
        return ""
    m = re.search(r"《(.+?)》", t)
    if not m:
        return ""
    inside = m.group(1).strip()
    for prefix in ("公开募集基金募集申请注册-", "公开募集基金变更注册-", "公开募集不动产投资信托基金募集申请注册-"):
        if inside.startswith(prefix):
            inside = inside[len(prefix) :].strip()
    # 若仍包含 '-'，取最后一段作为产品名
    if "-" in inside:
        inside = inside.split("-")[-1].strip()
    return inside


INDEX_KEYWORDS = [
    "中证",
    "国证",
    "上证",
    "深证",
    "沪深",
    "创业板",
    "科创板",
    "恒生",
    "标普",
    "纳斯达克",
    "道琼斯",
    "日经",
    "MSCI",
]


def extract_etf_theme_key(product_name: str) -> str:
    """
    ETF 主题/指数 key（用于判定“是否新主题”）。
    口径：尽量从指数/主题关键词起截断，并截到“交易型开放式”之前。
    """
    name = (product_name or "").strip()
    if not name:
        return ""

    # 仅对 ETF 主体做 key：截断到“交易型开放式”之前
    if "交易型开放式" in name:
        name = name.split("交易型开放式")[0].strip()

    # 去掉常见尾巴
    for suf in ("指数证券投资基金", "指数型证券投资基金", "交易型开放式指数证券投资基金"):
        if name.endswith(suf):
            name = name[: -len(suf)].strip()

    # 从指数关键词开始截（去掉发行人前缀如“华夏/广发/富国…”）
    pos = None
    for kw in INDEX_KEYWORDS:
        p = name.find(kw)
        if p != -1:
            pos = p if pos is None else min(pos, p)
    if pos is not None and pos > 0:
        name = name[pos:].strip()

    return name


def build_etf_new_theme_sheet_rows(records: List[Dict]) -> List[Dict]:
    """
    ETF「新主题」Sheet 用行：
    - 按 theme_key：只保留该主题下 application_date 最早的一批；
    - 若最早日期当天有多家公司（多条记录），则该日全部保留；
    - 输出顺序：各主题按其「最早出现日期」升序，主题内按日期、机构、标题排序；主题之间插空行。
    """
    by_theme: Dict[str, List[Dict]] = {}
    for r in records:
        t = (r.get("theme_key") or "").strip()
        if not t:
            continue
        by_theme.setdefault(t, []).append(r)

    filtered: Dict[str, List[Dict]] = {}
    for tkey, rows in by_theme.items():
        dates = [(rr.get("application_date") or "").strip() for rr in rows if (rr.get("application_date") or "").strip()]
        if not dates:
            continue
        min_d = min(dates)
        filtered[tkey] = [rr for rr in rows if (rr.get("application_date") or "").strip() == min_d]

    def _theme_first_date(tk: str) -> str:
        rs = filtered.get(tk) or []
        ds = [(rr.get("application_date") or "").strip() for rr in rs if (rr.get("application_date") or "").strip()]
        return min(ds) if ds else "9999-99-99"

    ordered_themes = sorted(filtered.keys(), key=lambda k: (_theme_first_date(k), k))
    out: List[Dict] = []
    prev_tk = None
    for tkey in ordered_themes:
        rows = sorted(
            filtered[tkey],
            key=lambda rr: (
                (rr.get("application_date") or "").strip(),
                (rr.get("applicant") or "").strip(),
                (rr.get("content") or "").strip(),
            ),
        )
        if prev_tk is not None:
            out.append({})
        out.extend(rows)
        prev_tk = tkey
    return out


def build_duplicate_product_registrations_rows(summary_rows: List[Dict], cutoff: str = "2026-01-01") -> List[Dict]:
    """
    从 summary_rows 中找出“产品名称重复注册”的情况：
    - 仅统计 application_date >= cutoff 的记录
    - 用 extract_product_name_from_content 抽取 product_name
    - product_name 相同且出现次数 >=2 的，全部列出（组间插空行）
    """
    rows: List[Dict] = []
    for r in summary_rows:
        d = (r.get("application_date") or "").strip()
        if not d or d < cutoff:
            continue
        product = (r.get("product_name") or "").strip()
        if not product:
            product = extract_product_name_from_content(r.get("content", ""))
        product = (product or "").strip()
        if not product:
            continue
        rr = dict(r)
        rr["product_name"] = product
        rr["applicant"] = rr.get("applicant") or extract_applicant_from_title(r.get("content", ""))
        rows.append(rr)

    by_name: Dict[str, List[Dict]] = {}
    for r in rows:
        by_name.setdefault(r["product_name"], []).append(r)

    dup_names = [k for k, v in by_name.items() if len(v) >= 2]
    if not dup_names:
        return []

    def _min_date(name: str) -> str:
        ds = [(x.get("application_date") or "").strip() for x in by_name.get(name, []) if (x.get("application_date") or "").strip()]
        return min(ds) if ds else "9999-99-99"

    dup_names.sort(key=lambda n: (_min_date(n), n))

    out: List[Dict] = []
    first = True
    for name in dup_names:
        items = by_name[name]
        items.sort(
            key=lambda x: (
                (x.get("application_date") or "").strip(),
                (x.get("reg_type") or "").strip(),
                (x.get("fund_type") or "").strip(),
                (x.get("applicant") or "").strip(),
                (x.get("content") or "").strip(),
            )
        )
        if not first:
            out.append({})
        out.extend(items)
        first = False
    return out


def build_duplicate_content_groups(summary_rows: List[Dict], cutoff: str = "2026-01-01") -> Dict[str, List[Dict]]:
    """
    找出 content 完全相同的重复记录（application_date >= cutoff），按 content 分组返回。
    组内按 application_date 倒序（最新在前）。
    """
    rows: List[Dict] = []
    for r in summary_rows:
        d = (r.get("application_date") or "").strip()
        if not d or d < cutoff:
            continue
        c = (r.get("content") or "").strip()
        if not c:
            continue
        rr = dict(r)
        rr["applicant"] = rr.get("applicant") or extract_applicant_from_title(c)
        rr["product_name"] = rr.get("product_name") or extract_product_name_from_content(c)
        rows.append(rr)

    by_content: Dict[str, List[Dict]] = {}
    for r in rows:
        by_content.setdefault(r["content"], []).append(r)

    by_content = {k: v for k, v in by_content.items() if len(v) >= 2}
    for k, v in by_content.items():
        v.sort(
            key=lambda x: (
                (x.get("application_date") or "").strip(),
                (x.get("reg_type") or "").strip(),
                (x.get("fund_type") or "").strip(),
                (x.get("applicant") or "").strip(),
            ),
            reverse=True,
        )
        by_content[k] = v

    return by_content


def _parse_ymd(s: str) -> Optional[datetime]:
    ss = (s or "").strip()
    if not ss:
        return None
    try:
        return datetime.strptime(ss[:10], "%Y-%m-%d")
    except Exception:
        return None


def business_days_elapsed(start_ymd: str, end_ymd: Optional[str] = None) -> Optional[int]:
    """
    计算从 start_ymd 到 end_ymd（默认今天）之间已过去的“工作日”数量。
    口径：
    - 不含 start 当天，从次日开始计数；
    - 优先按中国法定节假日 + 调休补班（若安装了 chinese_calendar）；
    - 若未安装，则退回到“仅排除周末”的口径。
    """
    start_dt = _parse_ymd(start_ymd)
    if not start_dt:
        return None
    end_dt = _parse_ymd(end_ymd) if end_ymd else datetime.now()
    if not end_dt:
        end_dt = datetime.now()
    if end_dt.date() <= start_dt.date():
        return 0

    def _is_business_day(day) -> bool:
        # chinese_calendar：工作日（含调休补班）返回 True；法定节假日返回 False
        if _cn_is_workday is not None:
            try:
                return bool(_cn_is_workday(day))
            except Exception:
                # 库异常时兜底
                return day.weekday() < 5
        # 兜底口径：仅排除周末
        return day.weekday() < 5

    d = start_dt.date() + timedelta(days=1)
    end_date = end_dt.date()
    cnt = 0
    while d <= end_date:
        if _is_business_day(d):
            cnt += 1
        d += timedelta(days=1)
    return cnt


def normalize_institution_text(text: str) -> str:
    """统一括号/空格后再做机构名包含匹配。"""
    if not text:
        return ""
    t = str(text).strip()
    t = t.translate(str.maketrans({
        "（": "(", "）": ")",
        "【": "(", "】": ")",
        "　": " ",
        "：": ":",
    }))
    t = re.sub(r"\s+", "", t)
    return t


def match_institution(title: str, keywords: List[str], inst_key: str = "") -> bool:
    """
    优先用「关于XXX的《」中的申请单位 XXX 匹配对标机构；
    避免仅靠产品名里的指数/主题关键词误命中其它机构。
    """
    applicant = extract_applicant_from_title(title)
    applicant_n = normalize_institution_text(applicant)
    title_n = normalize_institution_text(title or "")

    excludes = INSTITUTION_MATCH_EXCLUDES.get((inst_key or "").strip(), [])
    check_hay = applicant_n or title_n
    for ex in excludes:
        ex_n = normalize_institution_text(ex)
        if ex_n and ex_n in check_hay:
            return False

    kw_sorted = sorted(
        [k for k in keywords if (k or "").strip()],
        key=lambda x: len(normalize_institution_text(x)),
        reverse=True,
    )
    kw_norm = [(k, normalize_institution_text(k)) for k in kw_sorted]

    if applicant_n:
        for _, kn in kw_norm:
            if kn and kn in applicant_n:
                return True
        return False

    # 标题格式异常、解析不出申请单位时，才回退到全文匹配
    hay = title_n
    return any(kn and kn in hay for _, kn in kw_norm)


def group_with_blank_rows(
    rows: List[Dict],
    group_keys: List[str],
    *,
    within_group_date_desc: bool = False,
) -> List[Dict]:
    """按 group_keys 分组；组与组之间插入空行（{}）。"""
    out: List[Dict] = []
    if not rows:
        return out

    def _group_key(r: Dict):
        return tuple((r.get(k) or "") for k in group_keys)

    buckets: Dict[tuple, List[Dict]] = {}
    for r in rows:
        buckets.setdefault(_group_key(r), []).append(r)

    groups_sorted = sorted(buckets.keys())
    for gi, g in enumerate(groups_sorted):
        chunk = buckets[g]
        if within_group_date_desc:
            chunk.sort(
                key=lambda r: (r.get("application_date") or "0000-00-00"),
                reverse=True,
            )
        else:
            chunk.sort(
                key=lambda r: (
                    (r.get("application_date") or ""),
                    (r.get("content") or ""),
                )
            )
        out.extend(chunk)
        if gi != len(groups_sorted) - 1:
            out.append({})
    return out


def sort_summary_rows_by_date_desc(rows: List[Dict]) -> None:
    """进度汇总等 sheet：公示日期从新到旧；同日保持 reg_type / fund_type / content 顺序。"""
    rows.sort(
        key=lambda r: (
            (r.get("reg_type") or ""),
            (r.get("fund_type") or ""),
            (r.get("content") or ""),
        )
    )
    rows.sort(
        key=lambda r: (r.get("application_date") or "0000-00-00"),
        reverse=True,
    )


def is_gtht_record(title: str) -> bool:
    applicant = extract_applicant_from_title(title)
    hay = (applicant or "") + " " + (title or "")
    return any(k in hay for k in GTHT_KEYWORDS)


def _dedupe_rows_by_latest_content(rows: List[Dict]) -> List[Dict]:
    """同一 content（标题）可能重复出现，只保留 application_date 最新的那条。"""
    def _row_key_for_latest(r: Dict) -> tuple:
        app_date = (r.get("application_date") or "").strip()
        finish_date = (r.get("finish_date") or "").strip()
        try:
            order = int(str(r.get("task_order") or "").strip())
        except Exception:
            order = -1
        return (app_date, order, finish_date)

    latest_by_content: Dict[str, Dict] = {}
    for r in rows:
        c = (r.get("content") or "").strip()
        if not c:
            continue
        prev = latest_by_content.get(c)
        if prev is None or _row_key_for_latest(r) > _row_key_for_latest(prev):
            latest_by_content[c] = r
    return list(latest_by_content.values()) if latest_by_content else rows


def build_leaders_vs_gtht_rows(summary_rows: List[Dict]) -> Tuple[List[Dict], List[str]]:
    """
    按 (reg_type, fund_type, application_date) 分组：
    - 找到国泰海通在该组的最大 task_order
    - 筛出同组同日里 task_order > 国泰海通 的其他记录
    输出时包含“国泰海通基准行”作为参照物。
    """
    rows = _dedupe_rows_by_latest_content([dict(r) for r in summary_rows])

    # 构建每组国泰海通的基准进度（取该组国泰海通最大 task_order 的那一行）
    base_by_group: Dict[tuple, Dict] = {}
    for r in rows:
        group = (
            (r.get("reg_type") or "").strip(),
            (r.get("fund_type") or "").strip(),
            (r.get("application_date") or "").strip(),
        )
        try:
            order = int(str(r.get("task_order") or "").strip())
        except Exception:
            continue
        if not is_gtht_record(r.get("content", "")):
            continue
        prev = base_by_group.get(group)
        if prev is None:
            base_by_group[group] = dict(r)
        else:
            try:
                prev_order = int(str(prev.get("task_order") or "").strip())
            except Exception:
                prev_order = -1
            if order > prev_order:
                base_by_group[group] = dict(r)

    # 收集每组领先行
    peers_by_group: Dict[tuple, List[Dict]] = {}
    for r in rows:
        group = (
            (r.get("reg_type") or "").strip(),
            (r.get("fund_type") or "").strip(),
            (r.get("application_date") or "").strip(),
        )
        base = base_by_group.get(group)
        if base is None:
            continue
        try:
            base_order = int(str(base.get("task_order") or "").strip())
        except Exception:
            continue
        if is_gtht_record(r.get("content", "")):
            continue
        try:
            order = int(str(r.get("task_order") or "").strip())
        except Exception:
            continue
        if order <= base_order:
            continue
        rr = dict(r)
        rr["_peer_order"] = order
        peers_by_group.setdefault(group, []).append(rr)

    # 组内放一起，组间插空行
    out_rows: List[Dict] = []
    groups_sorted = sorted(peers_by_group.keys())
    for idx, group in enumerate(groups_sorted):
        base = base_by_group.get(group) or {}
        try:
            base_order = int(str(base.get("task_order") or "").strip())
        except Exception:
            base_order = None

        base_row = dict(base)
        base_row["compare_role"] = "gtht_baseline"
        base_row["gtht_task_order"] = str(base_order) if base_order is not None else ""
        out_rows.append(base_row)

        peers = peers_by_group[group]
        peers.sort(key=lambda r: (r.get("_peer_order", -1)), reverse=True)
        for p in peers:
            p2 = dict(p)
            p2.pop("_peer_order", None)
            p2["compare_role"] = "peer_ahead"
            p2["gtht_task_order"] = str(base_order) if base_order is not None else ""
            out_rows.append(p2)

        if idx != len(groups_sorted) - 1:
            out_rows.append({})  # 空行分隔组

    fieldnames = [
        "application_id",
        "content",
        "reg_type",
        "fund_type",
        "application_date",
        "progress_count",
        "task_name",
        "task_order",
        "finish_date",
        "all_task_names",
        "all_finish_dates",
        "compare_role",
        "gtht_task_order",
    ]
    return out_rows, fieldnames


def build_gtht_group_all_rows(
    summary_rows: List[Dict],
    exclude_product_names: Optional[set] = None,
) -> Tuple[List[Dict], List[str]]:
    """
    按 (reg_type, fund_type, application_date) 分组：
    - 只要组内包含国泰海通产品，就把该组同日同类的所有产品都输出（不再按 task 进度筛选）
    - 组内排序：国泰海通在前，其它在后；组与组之间插空行
    - exclude_product_names：同目录 Excel 中的产品名，命中后整行不出现在本 sheet
    """
    exclude = exclude_product_names or set()
    rows = _dedupe_rows_by_latest_content([dict(r) for r in summary_rows])

    groups: Dict[tuple, List[Dict]] = {}
    for r in rows:
        group = (
            (r.get("reg_type") or "").strip(),
            (r.get("fund_type") or "").strip(),
            (r.get("application_date") or "").strip(),
        )
        groups.setdefault(group, []).append(r)

    out: List[Dict] = []
    group_keys = list(groups.keys())
    group_keys.sort(key=lambda g: ((g[0] or ""), (g[1] or "")))
    group_keys.sort(key=lambda g: (g[2] or "0000-00-00"), reverse=True)
    emitted = 0
    for g in group_keys:
        items = groups[g]
        if exclude:
            items = [
                x
                for x in items
                if not is_group_compare_excluded_content(x.get("content", ""), exclude)
            ]
        if not items:
            continue
        gtht_items_all = [x for x in items if is_gtht_record(x.get("content", ""))]
        if not gtht_items_all:
            continue

        # 如果国泰海通该组进度已到“行政许可决定书”（task_order=5 / task_name 命中），则跳过该组，避免冗长
        def _is_decision_step(r: Dict) -> bool:
            try:
                if int(str(r.get("task_order") or "").strip()) == 5:
                    return True
            except Exception:
                pass
            return normalize_task_name(str(r.get("task_name") or "")) == "行政许可决定书"

        if any(_is_decision_step(x) for x in gtht_items_all):
            continue

        # 国泰海通在前，其余按 content 排序
        gtht_items = list(gtht_items_all)
        peer_items = [x for x in items if not is_gtht_record(x.get("content", ""))]
        gtht_items.sort(key=lambda r: (r.get("content") or ""))
        peer_items.sort(key=lambda r: (r.get("content") or ""))

        if emitted > 0:
            out.append({})
        for x in gtht_items:
            rr = dict(x)
            rr["compare_role"] = "gtht"
            out.append(rr)
        for x in peer_items:
            rr = dict(x)
            rr["compare_role"] = "peer"
            out.append(rr)
        emitted += 1

    fields = [
        "application_id",
        "content",
        "reg_type",
        "fund_type",
        "application_date",
        "progress_count",
        "task_name",
        "task_order",
        "finish_date",
        "all_task_names",
        "all_finish_dates",
        "compare_role",
    ]
    return out, fields


def export_progress_workbook_xlsx(
    data: List[Dict],
    xlsx_path: str,
    leaders_rows: Optional[List[Dict]] = None,
    leaders_fieldnames: Optional[List[str]] = None,
    selected_benchmarks: Optional[List[str]] = None,
    benchmark_data: Optional[List[Dict]] = None,
    benchmark_xlsx_path: Optional[str] = None,
    etf_theme_compare_data: Optional[List[Dict]] = None,
    etf_output_end_date: Optional[str] = None,
) -> Dict[str, object]:
    if not OPENPYXL_AVAILABLE:
        raise RuntimeError("缺少 openpyxl，无法导出 xlsx（请执行 pip3 install openpyxl）")

    ensure_group_compare_exclude_template()
    exclude_names = load_group_compare_exclude_products()

    summary_rows = [summarize_progress(r) for r in data]
    sort_summary_rows_by_date_desc(summary_rows)
    summary_fields = [
        "application_id",
        "content",
        "reg_type",
        "fund_type",
        "application_date",
        "progress_count",
        "task_name",
        "task_order",
        "finish_date",
        "all_task_names",
        "all_finish_dates",
    ]

    bond_rows = [r for r in summary_rows if (r.get("fund_type") or "").strip() == "债券型"]
    # 债券类全量（第 2 个 sheet）：I 列 finish_date 从新到旧
    bond_rows.sort(
        key=lambda r: (r.get("finish_date") or "0000-00-00"),
        reverse=True,
    )
    # 进度汇总（第 1 个 sheet）：B 列 content 须含「基金」
    summary_sheet_rows = [
        r for r in summary_rows if "基金" in str(r.get("content") or "")
    ]

    Path(xlsx_path).parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()

    # 高亮样式：国泰海通产品行
    gtht_fill = PatternFill("solid", fgColor="FFF2CC")  # 浅黄色
    gtht_font = Font(bold=True)
    # 重复注册：最新一条已到受理通知及之后
    dup_reg_fill = PatternFill("solid", fgColor="D9E1F2")  # 浅蓝色
    dup_reg_font = Font(bold=True)

    ws_sum = wb.active
    ws_sum.title = "进度汇总"
    ws_sum.append(summary_fields)
    for r in summary_sheet_rows:
        ws_sum.append([r.get(k, "") for k in summary_fields])
        # 若该行属于国泰海通，高亮
        if is_gtht_record(r.get("content", "")):
            row_idx = ws_sum.max_row
            for col in range(1, len(summary_fields) + 1):
                cell = ws_sum.cell(row=row_idx, column=col)
                cell.fill = gtht_fill
                cell.font = gtht_font

    ws_bond_all = wb.create_sheet("债券类全量")
    ws_bond_all.append(summary_fields)
    for r in bond_rows:
        ws_bond_all.append([r.get(k, "") for k in summary_fields])
        if is_gtht_record(r.get("content", "")):
            row_idx = ws_bond_all.max_row
            for col in range(1, len(summary_fields) + 1):
                cell = ws_bond_all.cell(row=row_idx, column=col)
                cell.fill = gtht_fill
                cell.font = gtht_font

    if leaders_rows is not None and leaders_fieldnames is not None:
        ws_cmp = wb.create_sheet("对比_gtht_vs_leaders")
        ws_cmp.append(leaders_fieldnames)
        for r in leaders_rows:
            ws_cmp.append([r.get(k, "") for k in leaders_fieldnames])
            if (r.get("compare_role") or "") == "gtht_baseline":
                row_idx = ws_cmp.max_row
                for col in range(1, len(leaders_fieldnames) + 1):
                    cell = ws_cmp.cell(row=row_idx, column=col)
                    cell.fill = gtht_fill
                    cell.font = gtht_font

    # 同批对比（全量）：E 列 application_date 从新到旧（按对比组排序，组内国泰海通在前）
    group_all_rows, group_all_fields = build_gtht_group_all_rows(
        summary_rows, exclude_product_names=exclude_names
    )
    ws_all = wb.create_sheet("同批对比_全量")
    ws_all.append(group_all_fields)
    for r in group_all_rows:
        ws_all.append([r.get(k, "") for k in group_all_fields])
        if (r.get("compare_role") or "") == "gtht":
            row_idx = ws_all.max_row
            for col in range(1, len(group_all_fields) + 1):
                cell = ws_all.cell(row=row_idx, column=col)
                cell.fill = gtht_fill
                cell.font = gtht_font

    # ETF 新主题（2026-01-01 起）：
    # - 对比口径：用全历史数据做 seen_before（etf_theme_compare_data；未提供则退回到当前 data）
    # - 输出口径：从 2026-01-01 起（可选限制到 etf_output_end_date），每主题只保留最早发布日；
    #   若最早日有多家公司则该日全部列出（见 build_etf_new_theme_sheet_rows）
    cutoff = "2026-01-01"
    etf_base = etf_theme_compare_data if etf_theme_compare_data is not None else data
    etf_base_summary = [summarize_progress(r) for r in etf_base]
    etf_rows = [r for r in etf_base_summary if (r.get("fund_type") or "").strip() == "ETF"]
    seen_before = set()
    for r in etf_rows:
        d = (r.get("application_date") or "").strip()
        if d and d < cutoff:
            theme = extract_etf_theme_key(extract_product_name_from_content(r.get("content", "")))
            if theme:
                seen_before.add(theme)

    new_theme_records: List[Dict] = []
    for r in etf_rows:
        d = (r.get("application_date") or "").strip()
        if not d or d < cutoff:
            continue
        if etf_output_end_date and d > etf_output_end_date:
            continue
        product = extract_product_name_from_content(r.get("content", ""))
        if is_etf_new_theme_excluded(product, r.get("content", "")):
            continue
        theme = extract_etf_theme_key(product)
        if not theme:
            continue
        if theme in seen_before:
            continue
        rr = dict(r)
        rr["applicant"] = extract_applicant_from_title(r.get("content", ""))
        rr["product_name"] = product
        rr["theme_key"] = theme
        new_theme_records.append(rr)

    grouped_new = build_etf_new_theme_sheet_rows(new_theme_records)

    etf_fields = [
        "theme_key",
        "application_date",
        "reg_type",
        "fund_type",
        "applicant",
        "product_name",
        "content",
        "task_name",
        "task_order",
        "finish_date",
    ]
    ws_etf = wb.create_sheet("ETF新主题_20260101起")
    ws_etf.append(etf_fields)
    for r in grouped_new:
        ws_etf.append([r.get(k, "") for k in etf_fields])

    # ETF - 接收材料超过 7 个工作日未推进
    # 条件：fund_type=ETF 且当前 task_name=接收材料（兼容“接受材料”写法），并且从 finish_date（无则 application_date）起已过去 >=7 个工作日
    def _is_accept_material_step(tn: str) -> bool:
        n = normalize_task_name(str(tn or ""))
        return n in ("接收材料", "接受材料")

    today_ymd = datetime.now().strftime("%Y-%m-%d")
    accept_stuck_rows: List[Dict] = []
    # 若同一 content 在后续日期又出现（疑似重新注册），则剔除较早那条“卡在接收材料”的记录
    latest_app_date_by_content: Dict[str, str] = {}
    for rr in etf_rows:
        c = (rr.get("content") or "").strip()
        d = (rr.get("application_date") or "").strip()
        if not c or not d:
            continue
        prev = latest_app_date_by_content.get(c)
        if prev is None or d > prev:
            latest_app_date_by_content[c] = d

    for r in etf_rows:
        # 仅统计 2026 年起
        app_date = (r.get("application_date") or "").strip()
        if app_date < cutoff:
            continue
        c = (r.get("content") or "").strip()
        if c and latest_app_date_by_content.get(c) and app_date < latest_app_date_by_content.get(c):
            continue
        if not _is_accept_material_step(r.get("task_name", "")):
            continue
        base = (r.get("finish_date") or "").strip() or (r.get("application_date") or "").strip()
        bd = business_days_elapsed(base, today_ymd)
        if bd is None:
            continue
        if bd >= 7:
            rr = dict(r)
            rr["applicant"] = rr.get("applicant") or extract_applicant_from_title(rr.get("content", ""))
            rr["product_name"] = rr.get("product_name") or extract_product_name_from_content(rr.get("content", ""))
            rr["waiting_business_days"] = bd
            rr["base_date_for_waiting"] = base
            accept_stuck_rows.append(rr)

    accept_stuck_rows.sort(
        key=lambda x: (
            int(x.get("waiting_business_days") or 0),
            (x.get("base_date_for_waiting") or ""),
            (x.get("application_date") or ""),
            (x.get("content") or ""),
        ),
        reverse=True,
    )

    ws_accept = wb.create_sheet("ETF_接收材料超7工作日")
    accept_fields = [
        "waiting_business_days",
        "base_date_for_waiting",
        "application_date",
        "reg_type",
        "fund_type",
        "task_name",
        "task_order",
        "finish_date",
        "applicant",
        "product_name",
        "content",
        "application_id",
    ]
    ws_accept.append(accept_fields)
    for r in accept_stuck_rows:
        ws_accept.append([r.get(k, "") for k in accept_fields])

    # 2026 起：content 完全相同的重复记录（展示最新一条，其余折叠可展开）
    dup_cutoff = "2026-01-01"
    dup_base = etf_base_summary  # 全量 progress 的汇总（若传入全历史则为全历史）
    dup_groups = build_duplicate_content_groups(dup_base, cutoff=dup_cutoff)
    ws_dup = wb.create_sheet("重复注册_2026起")
    try:
        ws_dup.sheet_properties.outlinePr.summaryBelow = True
        ws_dup.sheet_view.showOutlineSymbols = True
    except Exception:
        pass

    dup_fields = [
        "dup_count",
        "is_latest",
        "application_date",
        "accept_material_date",
        "reg_type",
        "fund_type",
        "applicant",
        "product_name",
        "task_name",
        "task_order",
        "finish_date",
        "content",
        "application_id",
    ]
    ws_dup.append(dup_fields)
    # 表头中文备注（字段名仍为英文 key）
    dup_header_labels = {
        "accept_material_date": "接收材料日期",
    }
    for col_idx, key in enumerate(dup_fields, start=1):
        label = dup_header_labels.get(key)
        if label:
            ws_dup.cell(row=1, column=col_idx, value=label)

    def _group_sort_key(item):
        content, rows = item
        d = (rows[0].get("application_date") or "").strip() if rows else "9999-99-99"
        return (d, content)

    for content, rows in sorted(dup_groups.items(), key=_group_sort_key, reverse=True):
        dup_count = len(rows)
        if not rows:
            continue

        latest = rows[0]
        accept_date_latest = extract_accept_material_date(latest)

        rr = dict(latest)
        rr["dup_count"] = dup_count
        rr["is_latest"] = "Y"
        rr["accept_material_date"] = accept_date_latest if record_task_order_num(latest) > 1 else ""
        ws_dup.append([rr.get(k, "") for k in dup_fields])
        latest_row_idx = ws_dup.max_row
        highlight_rows: List[int] = []

        step1_row = build_latest_submission_step1_row(latest)
        if step1_row is not None:
            step1_row["dup_count"] = dup_count
            step1_row["is_latest"] = ""
            ws_dup.append([step1_row.get(k, "") for k in dup_fields])
            highlight_rows = [latest_row_idx, ws_dup.max_row]

        start_hidden_row = ws_dup.max_row + 1
        for other in rows[1:]:
            rr2 = dict(other)
            rr2["dup_count"] = dup_count
            rr2["is_latest"] = ""
            rr2["accept_material_date"] = extract_accept_material_date(other)
            ws_dup.append([rr2.get(k, "") for k in dup_fields])
        end_hidden_row = ws_dup.max_row

        for row_idx in highlight_rows:
            for col in range(1, len(dup_fields) + 1):
                cell = ws_dup.cell(row=row_idx, column=col)
                cell.fill = dup_reg_fill
                cell.font = dup_reg_font

        # 折叠其它重复行（Excel 左侧可点 “+/-” 展开/收起）
        if end_hidden_row >= start_hidden_row:
            try:
                ws_dup.row_dimensions.group(start_hidden_row, end_hidden_row, hidden=True, outline_level=1)
            except Exception:
                for ri in range(start_hidden_row, end_hidden_row + 1):
                    try:
                        ws_dup.row_dimensions[ri].hidden = True
                        ws_dup.row_dimensions[ri].outlineLevel = 1
                    except Exception:
                        pass
        # 组间空行
        ws_dup.append(["" for _ in dup_fields])

    wb.save(xlsx_path)

    benchmark_counts: Dict[str, int] = {}
    benchmark_file = ""
    if selected_benchmarks:
        if benchmark_data is None:
            raise RuntimeError("对标机构数据未提供")
        bpath = benchmark_xlsx_path or export_benchmark_xlsx_path(xlsx_path)
        benchmark_counts = _write_benchmark_workbook(
            bpath,
            benchmark_data=benchmark_data,
            selected_benchmarks=selected_benchmarks,
            summary_fields=summary_fields,
        )
        benchmark_file = bpath

    return {
        "group_compare_exclude_path": group_compare_exclude_products_path(),
        "group_compare_exclude_count": len(exclude_names),
        "benchmark_xlsx_path": benchmark_file,
        "benchmark_counts": benchmark_counts,
    }


def export_progress_leaders_csv(input_csv_path: str, output_csv_path: str) -> int:
    """兼容旧用法：从 summary CSV 生成对比行（仍可用于手工脚本）。"""
    with open(input_csv_path, "r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        rows = list(reader)
    out_rows, fieldnames = build_leaders_vs_gtht_rows(rows)

    Path(output_csv_path).parent.mkdir(parents=True, exist_ok=True)
    with open(output_csv_path, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(out_rows)

    return len(out_rows)


class CSRCClient:
    """CSRC API 客户端。"""

    def __init__(self, delay: float = REQUEST_DELAY):
        self.session = requests.Session()
        self.session.trust_env = False
        self.session.proxies = {"http": None, "https": None}
        self.session.headers.update(DEFAULT_HEADERS)
        self.delay = delay
        self._last_request_time = 0
        self._verify_ssl = True

    def _rate_limit(self):
        elapsed = time.time() - self._last_request_time
        if elapsed < self.delay:
            time.sleep(self.delay - elapsed)
        self._last_request_time = time.time()

    def request(
        self,
        method: str,
        endpoint: str,
        params: Optional[Dict] = None,
        json_data: Optional[Dict] = None,
        retries: int = MAX_RETRIES,
    ) -> Dict:
        url = f"{BASE_URL}{endpoint}"
        last_error = None

        for attempt in range(1, retries + 1):
            try:
                self._rate_limit()
                req_kwargs = {
                    "timeout": 30,
                    "proxies": {"http": None, "https": None},
                    "verify": self._verify_ssl,
                }
                if method.upper() == "GET":
                    resp = self.session.get(url, params=params, **req_kwargs)
                else:
                    resp = self.session.post(url, json=json_data, **req_kwargs)
                resp.raise_for_status()
                data = resp.json()
                if data.get("code") != "0000":
                    raise RuntimeError(
                        f"API错误: {data.get('code')} - {data.get('message')}"
                    )
                return data
            except requests.exceptions.SSLError as e:
                if self._verify_ssl:
                    self._verify_ssl = False
                    last_error = e
                    continue
                last_error = e
                if attempt < retries:
                    time.sleep(RETRY_DELAY * attempt)
            except Exception as e:
                last_error = e
                if attempt < retries:
                    time.sleep(RETRY_DELAY * attempt)

        raise last_error

    def get_progress_list(
        self, page_num: int = 1, page_size: int = PAGE_SIZE, query: str = ""
    ) -> Dict:
        params = {
            "appMatrCde": "",
            "queryCondition": query,
            "pageNum": page_num,
            "pageSize": page_size,
        }
        return self.request("GET", API_ENDPOINTS["progress"], params=params)

    def get_result_list(self, page_num: int = 1, page_size: int = PAGE_SIZE) -> Dict:
        params = {"pageNum": page_num, "pageSize": page_size}
        return self.request("GET", API_ENDPOINTS["result"], params=params)

    def get_review_list(self, page_num: int = 1, page_size: int = PAGE_SIZE) -> Dict:
        params = {"pageNum": page_num, "pageSize": page_size}
        return self.request("GET", API_ENDPOINTS["review"], params=params)

    def get_stats(self) -> Dict:
        return self.request("GET", API_ENDPOINTS["stats"])


def extract_progress_records(raw_data: Dict) -> List[Dict]:
    records = raw_data.get("data", {}).get("records", [])
    result = []
    for r in records:
        progress_list = r.get("aprvSchdPubFlowViewResultList", [])
        progress_summary = []
        for p in progress_list:
            progress_summary.append({
                "task_name": normalize_task_name(p.get("taskName", "")),
                "finish_date": p.get("fnshDate", ""),
                "is_link": p.get("isLink", ""),
            })
        content = r.get("showCntnt", "")
        fund_type = classify_fund(content)
        if fund_type is None:
            continue
        result.append({
            "application_id": r.get("alAppLtCde", ""),
            "content": content,
            "reg_type": classify_reg_type(content),
            "fund_type": fund_type,
            "application_date": r.get("appDate", ""),
            "progress_count": len(progress_list),
            "progress": progress_summary,
        })
    return result


def extract_result_records(raw_data: Dict) -> List[Dict]:
    records = raw_data.get("data", {}).get("records", [])
    result = []
    for r in records:
        matter_name = r.get("appMatrName", "")
        fund_type = classify_fund(matter_name)
        if fund_type is None:
            continue
        result.append({
            "result_id": r.get("aprvResultPubCde", ""),
            "application_id": r.get("alAppLtCde", ""),
            "applicant": r.get("apptName", ""),
            "matter_name": matter_name,
            "reg_type": classify_reg_type(matter_name),
            "fund_type": fund_type,
            "description": r.get("dsptName", ""),
            "decision_date": r.get("dsptDate", ""),
            "file_code": r.get("alFileCde", ""),
        })
    return result


def extract_review_records(raw_data: Dict) -> List[Dict]:
    records = raw_data.get("data", {}).get("records", [])
    result = []
    for r in records:
        result.append({
            "review_id": r.get("doPcAprsId", ""),
            "department": r.get("ppdmDescrib", ""),
            "matter_code": r.get("appMatrCde", ""),
            "matter_name": r.get("matrName", ""),
            "reviewer": r.get("aprsPsnName", ""),
            "rating": r.get("ovalStftTypeCde", ""),
            "content": r.get("aprsCntnt", ""),
            "review_time": r.get("aprsTime", ""),
        })
    return result


def _get_record_date(record: Dict, mode: str) -> str:
    if mode == "progress":
        return record.get("application_date", "")
    elif mode == "result":
        return record.get("decision_date", "")
    elif mode == "review":
        return record.get("review_time", "")[:10]
    return ""


def _match_date(record: Dict, date_filter: str, mode: str) -> bool:
    return _get_record_date(record, mode) == date_filter


def _match_date_range(record: Dict, mode: str, start_date: str, end_date: str) -> bool:
    d = _get_record_date(record, mode)
    return start_date <= d <= end_date


def _has_records_before_start(records: List[Dict], mode: str, start_date: str) -> bool:
    if not records:
        return False
    for r in records:
        d = _get_record_date(r, mode)
        if d >= start_date:
            return False
    return True


def flatten_progress(record: Dict) -> List[Dict]:
    base = {
        "application_id": record["application_id"],
        "content": record["content"],
        "reg_type": record.get("reg_type", ""),
        "fund_type": record.get("fund_type", ""),
        "application_date": record["application_date"],
    }
    if not record.get("progress"):
        return [base]
    rows = []
    for i, p in enumerate(record["progress"]):
        row = base.copy()
        row.update({
            "task_seq": i + 1,
            "task_name": p["task_name"],
            "task_order": task_order(p.get("task_name", "")) or "",
            "finish_date": p["finish_date"],
            "is_link": p["is_link"],
        })
        rows.append(row)
    return rows


def summarize_progress(record: Dict) -> Dict:
    base = {
        "application_id": record.get("application_id", ""),
        "content": record.get("content", ""),
        "reg_type": record.get("reg_type", ""),
        "fund_type": record.get("fund_type", ""),
        "application_date": record.get("application_date", ""),
    }

    progress = record.get("progress") or []
    if not progress:
        base.update(
            {
                "progress_count": 0,
                "task_name": "",
                "task_order": "",
                "finish_date": "",
                "all_task_names": "",
                "all_finish_dates": "",
            }
        )
        return base

    task_names = [p.get("task_name", "") for p in progress if p.get("task_name", "")]
    finish_dates = [p.get("finish_date", "") for p in progress if p.get("finish_date", "")]

    last = None
    for p in reversed(progress):
        if p.get("task_name") or p.get("finish_date"):
            last = p
            break
    last = last or {}

    base.update(
        {
            "progress_count": len(progress),
            "task_name": last.get("task_name", ""),
            "task_order": task_order(last.get("task_name", "")) or "",
            "finish_date": last.get("finish_date", ""),
            "all_task_names": " > ".join(task_names),
            "all_finish_dates": " > ".join(finish_dates),
        }
    )
    return base


def export_csv(
    data: List[Dict],
    filepath: str,
    mode: str = "progress",
    progress_layout: str = "summary",
):
    Path(filepath).parent.mkdir(parents=True, exist_ok=True)
    if not data:
        return 0

    if mode == "progress":
        if progress_layout == "flatten":
            flat_data = []
            for record in data:
                flat_data.extend(flatten_progress(record))
            fieldnames = [
                "application_id",
                "content",
                "reg_type",
                "fund_type",
                "application_date",
                "task_seq",
                "task_name",
                "task_order",
                "finish_date",
                "is_link",
            ]
            write_data = flat_data
        else:
            summary_data = [summarize_progress(r) for r in data]
            fieldnames = [
                "application_id",
                "content",
                "reg_type",
                "fund_type",
                "application_date",
                "progress_count",
                "task_name",
                "task_order",
                "finish_date",
                "all_task_names",
                "all_finish_dates",
            ]
            write_data = summary_data
    elif mode == "result":
        fieldnames = [
            "result_id", "application_id", "applicant", "matter_name",
            "reg_type", "fund_type", "description", "decision_date", "file_code",
        ]
        write_data = data
    elif mode == "review":
        fieldnames = [
            "review_id", "department", "matter_code", "matter_name",
            "reviewer", "rating", "content", "review_time",
        ]
        write_data = data
    else:
        raise ValueError(f"未知模式: {mode}")

    with open(filepath, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(write_data)

    return len(write_data)


def export_json(data: List[Dict], filepath: str):
    Path(filepath).parent.mkdir(parents=True, exist_ok=True)
    with open(filepath, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


# ============================================================
# GUI 界面
# ============================================================

import tkinter as tk
from tkinter import ttk, filedialog
import subprocess

FONT_TITLE = ("PingFang SC" if sys.platform == "darwin" else "Microsoft YaHei", 16, "bold")
FONT_SUBTITLE = ("PingFang SC" if sys.platform == "darwin" else "Microsoft YaHei", 10)
FONT_FRAME = ("PingFang SC" if sys.platform == "darwin" else "Microsoft YaHei", 12, "bold")
FONT_BTN = ("PingFang SC" if sys.platform == "darwin" else "Microsoft YaHei", 13, "bold")
FONT_BTN2 = ("PingFang SC" if sys.platform == "darwin" else "Microsoft YaHei", 13)
FONT_LOG = ("Menlo" if sys.platform == "darwin" else "Consolas", 10)


def _open_path(path: str) -> None:
    """用系统默认方式打开文件或文件夹。"""
    if not path or not os.path.exists(path):
        return
    try:
        if sys.platform == "darwin":
            subprocess.Popen(["open", path])
        elif sys.platform.startswith("win"):
            os.startfile(path)  # type: ignore[attr-defined]
        else:
            subprocess.Popen(["xdg-open", path])
    except Exception:
        pass


class GUILogHandler(logging.Handler):
    """将日志输出到 Tkinter Text 控件的 Handler。"""

    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget
        self.setFormatter(logging.Formatter(
            "[%(asctime)s] %(levelname)s | %(message)s",
            datefmt="%H:%M:%S",
        ))

    def emit(self, record):
        msg = self.format(record) + "\n"
        self.text_widget.insert("end", msg)
        self.text_widget.see("end")


class CSRCScraperGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("CSRC 行政许可审批进度爬虫")
        # 更适合“对标机构 + 日志”同时展示的默认尺寸
        self.root.geometry("980x820" if sys.platform == "darwin" else "920x780")
        self.root.minsize(860, 680)
        self._prefs = load_prefs()
        self._last_exported_files: List[str] = []

        if sys.platform == "darwin":
            self.root.tk.call("tk", "scaling", 2.0)

        # 统一 ttk 主题与基础样式（让界面更“像软件”）
        try:
            style = ttk.Style()
            # macOS 默认主题控件偏“空”，clam 更耐看且跨平台一致
            style.theme_use("clam")
            style.configure("TLabel", font=("", 11))
            style.configure("TLabelframe.Label", font=FONT_FRAME)
            style.configure("TButton", font=("", 11))
        except Exception:
            pass

        self._build_ui()
        self._setup_logging()
        self._apply_prefs()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    def _build_ui(self):
        padx, pady = 16, 10
        self.root.configure(bg="#f6f7fb")

        # --- 标题 ---
        tk.Label(self.root, text="中国证监会行政许可审批进度爬虫", font=FONT_TITLE, fg="#1a73e8", bg="#f6f7fb").pack(pady=(pady, 4))
        tk.Label(self.root, text="https://neris.csrc.gov.cn/alappr-delare-front/home/toPubFlow", font=FONT_SUBTITLE, fg="#666666", bg="#f6f7fb").pack(pady=(0, 4))

        # 快捷预设：每天打开先点一下即可，再微调对标机构
        preset_bar = tk.Frame(self.root, bg="#f6f7fb")
        preset_bar.pack(fill="x", padx=padx, pady=(0, pady))
        tk.Label(preset_bar, text="快捷：", bg="#f6f7fb", fg="#555555", font=("", 11)).pack(side="left")
        for text, key in [
            ("今日", "today"),
            ("昨日", "yesterday"),
            ("日常全量(自开始日)", "all"),
            ("自定义范围", "range"),
        ]:
            ttk.Button(preset_bar, text=text, width=14, command=lambda k=key: self._apply_preset(k)).pack(
                side="left", padx=(6, 0)
            )

        # 主分割：上（配置）/ 下（操作+日志）可拖拽调整高度，避免窗口变矮时看不到按钮
        self._main_pane = tk.PanedWindow(self.root, orient="vertical", sashrelief="raised")
        self._main_pane.pack(fill="both", expand=True, padx=padx, pady=(0, pady))

        # --- 配置区域（可滚动，避免窗口较小时“显示不全”）---
        cfg_outer = tk.Frame(self._main_pane, bg="#f6f7fb")
        self._main_pane.add(cfg_outer, stretch="never")

        self._cfg_canvas = tk.Canvas(cfg_outer, highlightthickness=0, bg="#f6f7fb")
        cfg_scrollbar = tk.Scrollbar(cfg_outer, orient="vertical", command=self._cfg_canvas.yview)
        self._cfg_canvas.configure(yscrollcommand=cfg_scrollbar.set)
        self._cfg_canvas.pack(side="left", fill="both", expand=True)
        cfg_scrollbar.pack(side="right", fill="y")

        cfg_inner = tk.Frame(self._cfg_canvas)
        self._cfg_window_id = self._cfg_canvas.create_window((0, 0), window=cfg_inner, anchor="nw")

        def _on_cfg_frame_configure(_evt=None):
            self._cfg_canvas.configure(scrollregion=self._cfg_canvas.bbox("all"))

        def _on_cfg_canvas_configure(evt):
            # 让内部宽度跟随 canvas，避免横向挤压
            try:
                self._cfg_canvas.itemconfigure(self._cfg_window_id, width=evt.width)
            except Exception:
                pass

        cfg_inner.bind("<Configure>", _on_cfg_frame_configure)
        self._cfg_canvas.bind("<Configure>", _on_cfg_canvas_configure)

        def _on_cfg_mousewheel(event):
            delta = event.delta
            if delta == 0:
                return
            step = -1 if delta > 0 else 1
            self._cfg_canvas.yview_scroll(step, "units")

        self._cfg_canvas.bind("<Enter>", lambda _e: self._cfg_canvas.bind_all("<MouseWheel>", _on_cfg_mousewheel))
        self._cfg_canvas.bind("<Leave>", lambda _e: self._cfg_canvas.unbind_all("<MouseWheel>"))

        frame = ttk.LabelFrame(cfg_inner, text="爬取配置", padding=(padx, pady))
        frame.pack(fill="x", expand=True, padx=0, pady=(0, 8))

        # 统一左侧标签宽度与对齐
        def _mk_label(parent, text: str):
            return ttk.Label(parent, text=f"{text}：", width=10, anchor="e")

        # 模式选择
        row1 = ttk.Frame(frame)
        row1.pack(fill="x", pady=4)
        _mk_label(row1, "爬取模式").pack(side="left")
        self.mode_var = tk.StringVar(value="progress")
        combo = ttk.Combobox(row1, textvariable=self.mode_var, values=["progress", "result", "review"], state="readonly", width=18)
        combo.pack(side="left", padx=(8, 0))
        ttk.Label(row1, text="（进度 / 结果 / 评价）", foreground="#777777").pack(side="left", padx=(10, 0))

        # 基金类型筛选
        row1b = ttk.Frame(frame)
        row1b.pack(fill="x", pady=4)
        _mk_label(row1b, "基金类型").pack(side="left")
        self.fund_type_var = tk.StringVar(value="全部")
        ttk.Combobox(
            row1b,
            textvariable=self.fund_type_var,
            values=["全部", "货币型", "债券型", "混合型", "股票型", "REITs", "QDII", "ETF", "FOF"],
            state="readonly",
            width=16,
        ).pack(side="left", padx=(8, 0))
        ttk.Label(row1b, text="（可选，只导出指定类型）", foreground="#777777").pack(side="left", padx=(10, 0))

        ttk.Separator(frame, orient="horizontal").pack(fill="x", pady=8)

        # 日期类型
        row2 = ttk.Frame(frame)
        row2.pack(fill="x", pady=(0, 6))
        _mk_label(row2, "日期类型").pack(side="left")
        self.date_type_var = tk.StringVar(value="today")
        for val, label in [("today", "今日"), ("yesterday", "昨日"), ("range", "日期范围"), ("single", "单日"), ("all", "全部历史")]:
            ttk.Radiobutton(row2, text=label, variable=self.date_type_var, value=val, command=self._on_date_type_change).pack(side="left", padx=(10, 0))

        # 开始日期（始终可见；不修改则默认 2024-01-01）
        row2b = ttk.Frame(frame)
        row2b.pack(fill="x", pady=4)
        _mk_label(row2b, "开始日期").pack(side="left")
        self.data_start_entry = ttk.Entry(row2b, width=14, justify="center")
        self.data_start_entry.pack(side="left", padx=(8, 0))
        self.data_start_entry.insert(0, DEFAULT_DATA_START_DATE)
        ttk.Label(row2b, text="（默认 2024-01-01；「全部历史」亦从此日起）", foreground="#777777").pack(
            side="left", padx=(10, 0)
        )

        # --- 日期输入区：范围（仅结束日期；起始日见上方「开始日期」）---
        self.range_frame = ttk.Frame(frame)
        self.range_frame.pack(fill="x", pady=4)
        _mk_label(self.range_frame, "结束日期").pack(side="left")
        self.end_entry = ttk.Entry(self.range_frame, width=14, justify="center")
        self.end_entry.pack(side="left", padx=(8, 0))
        self.end_entry.insert(0, datetime.now().strftime("%Y-%m-%d"))

        # --- 日期输入区：单日 ---
        self.single_frame = ttk.Frame(frame)
        _mk_label(self.single_frame, "目标日期").pack(side="left")
        self.single_entry = ttk.Entry(self.single_frame, width=14, justify="center")
        self.single_entry.pack(side="left", padx=(8, 0))
        self.single_entry.insert(0, datetime.now().strftime("%Y-%m-%d"))

        ttk.Separator(frame, orient="horizontal").pack(fill="x", pady=8)

        # 导出格式（默认多 Sheet Excel 台账）
        row4 = ttk.Frame(frame)
        row4.pack(fill="x", pady=(0, 6))
        _mk_label(row4, "导出格式").pack(side="left")
        self.format_var = tk.StringVar(value="excel")
        for val, label in [
            ("excel", "多Sheet Excel（推荐）"),
            ("json", "JSON"),
            ("both", "Excel + JSON"),
            ("csv", "仅 CSV"),
        ]:
            ttk.Radiobutton(row4, text=label, variable=self.format_var, value=val).pack(side="left", padx=(10, 0))

        # 进度导出布局（仅 progress 生效）
        row4b = ttk.Frame(frame)
        row4b.pack(fill="x", pady=4)
        _mk_label(row4b, "进度导出").pack(side="left")
        self.progress_layout_var = tk.StringVar(value="summary")
        ttk.Radiobutton(
            row4b,
            text="汇总台账（多Sheet）",
            variable=self.progress_layout_var,
            value="summary",
        ).pack(side="left", padx=(4, 10))
        ttk.Radiobutton(
            row4b,
            text="明细（每个流程节点一行）",
            variable=self.progress_layout_var,
            value="flatten",
        ).pack(side="left", padx=(4, 10))

        # 对标机构（滚动 + 搜索，避免挤掉底部控件）
        bench_frame = ttk.LabelFrame(frame, text="对标机构（勾选后另存对标同业专表）")
        bench_frame.pack(fill="x", expand=False, pady=6)

        bench_top = ttk.Frame(bench_frame)
        bench_top.pack(fill="x", padx=8, pady=(6, 4))
        ttk.Label(bench_top, text="搜索：", foreground="#555555").pack(side="left")
        self.bench_search_var = tk.StringVar(value="")
        bench_search_entry = ttk.Entry(bench_top, textvariable=self.bench_search_var, width=20)
        bench_search_entry.pack(side="left", padx=(8, 0))
        self.bench_search_var.trace_add("write", lambda *_: self._render_benchmark_checkboxes())

        ttk.Button(bench_top, text="常用对标", command=self._select_common_benchmarks, width=8).pack(
            side="left", padx=(10, 0)
        )
        ttk.Button(bench_top, text="全选可见", command=self._select_visible_benchmarks, width=8).pack(
            side="left", padx=(6, 0)
        )
        ttk.Button(bench_top, text="清空", command=self._clear_benchmarks, width=6).pack(
            side="left", padx=(6, 0)
        )
        ttk.Button(
            bench_top,
            text="查看清单",
            command=self._view_benchmark_institutions,
            width=8,
        ).pack(side="left", padx=(6, 0))

        # 可滚动区域
        bench_scroll = ttk.Frame(bench_frame)
        bench_scroll.pack(fill="both", expand=True, padx=8, pady=(0, 8))

        self.bench_vars = {}
        # 固定一个较小高度，避免把后续输入区/按钮/日志挤没；内容通过滚动查看
        self._bench_canvas = tk.Canvas(bench_scroll, height=150, highlightthickness=0)
        self._bench_scrollbar = tk.Scrollbar(bench_scroll, orient="vertical", command=self._bench_canvas.yview)
        self._bench_canvas.configure(yscrollcommand=self._bench_scrollbar.set)
        self._bench_scrollbar.pack(side="right", fill="y")
        self._bench_canvas.pack(side="left", fill="both", expand=True)

        self._bench_container = tk.Frame(self._bench_canvas)
        self._bench_window_id = self._bench_canvas.create_window((0, 0), window=self._bench_container, anchor="nw")

        def _on_frame_configure(_evt=None):
            self._bench_canvas.configure(scrollregion=self._bench_canvas.bbox("all"))

        def _on_canvas_configure(evt):
            # 让内部 frame 宽度跟随 canvas，避免横向挤压
            try:
                self._bench_canvas.itemconfigure(self._bench_window_id, width=evt.width)
            except Exception:
                pass

        self._bench_container.bind("<Configure>", _on_frame_configure)
        self._bench_canvas.bind("<Configure>", _on_canvas_configure)

        # 鼠标滚轮支持（mac/windows）
        def _on_mousewheel(event):
            # macOS: event.delta 通常是 +/-1 或 +/-120；Windows: +/-120
            delta = event.delta
            if delta == 0:
                return
            step = -1 if delta > 0 else 1
            self._bench_canvas.yview_scroll(step, "units")

        # 仅在鼠标悬停在对标列表区域时滚动，避免全局绑定影响其它控件
        self._bench_canvas.bind("<Enter>", lambda _e: self._bench_canvas.bind_all("<MouseWheel>", _on_mousewheel))
        self._bench_canvas.bind("<Leave>", lambda _e: self._bench_canvas.unbind_all("<MouseWheel>"))
        # 对标机构列表内置在本 py 文件，启动时加载一次
        self._bench_inst_map_cache = load_benchmark_institutions()
        self._render_benchmark_checkboxes()

        # 对标机构时间范围（独立于上面的日期范围）
        row4d = ttk.Frame(frame)
        row4d.pack(fill="x", pady=4)
        _mk_label(row4d, "对标时间").pack(side="left")
        self.bench_start_entry = ttk.Entry(row4d, width=14, justify="center")
        self.bench_start_entry.pack(side="left", padx=(8, 0))
        self.bench_start_entry.insert(0, DEFAULT_DATA_START_DATE)
        ttk.Label(row4d, text="至").pack(side="left", padx=(12, 0))
        self.bench_end_entry = ttk.Entry(row4d, width=14, justify="center")
        self.bench_end_entry.pack(side="left", padx=(8, 0))
        self.bench_end_entry.insert(0, datetime.now().strftime("%Y-%m-%d"))
        ttk.Label(row4d, text="（仅影响对标机构 Sheet）", foreground="#777777").pack(
            side="left", padx=(8, 0)
        )

        # 输出位置（默认技能目录/输出；可改）
        row5 = ttk.Frame(frame)
        row5.pack(fill="x", pady=4)
        _mk_label(row5, "输出目录").pack(side="left")
        self.output_dir_var = tk.StringVar(value=default_output_directory())
        ttk.Entry(row5, textvariable=self.output_dir_var, width=42).pack(side="left", padx=(8, 0))
        ttk.Button(row5, text="选择…", width=6, command=self._choose_output_dir).pack(side="left", padx=(8, 0))
        ttk.Label(row5, text="文件名：证监会网页数据+当日", foreground="#777777").pack(side="left", padx=(8, 0))

        row5b = ttk.Frame(frame)
        row5b.pack(fill="x", pady=2)
        self.open_when_done_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            row5b,
            text="完成后打开输出文件夹（并尝试打开 Excel）",
            variable=self.open_when_done_var,
        ).pack(side="left", padx=(88, 0))

        row5c = ttk.Frame(frame)
        row5c.pack(fill="x", pady=2)
        # 默认关闭：全历史翻页是日常最慢的一步
        self.etf_full_history_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            row5c,
            text="ETF新主题：额外抓全历史对比（很慢，日常可不勾）",
            variable=self.etf_full_history_var,
        ).pack(side="left", padx=(88, 0))

        # 底部：按钮在上、日志在下且占满剩余高度
        ops_frame = tk.Frame(self._main_pane)
        self._main_pane.add(ops_frame, stretch="always")

        # --- 按钮行（更紧凑，减少空白）---
        btn_frame = tk.Frame(ops_frame)
        btn_frame.pack(fill="x", pady=(6, 8))

        style = ttk.Style()
        style.configure("Run.TButton", font=FONT_BTN)
        style.configure("Stop.TButton", font=FONT_BTN2)

        self.run_btn = ttk.Button(btn_frame, text="开始爬取", style="Run.TButton", command=self._on_run, width=14)
        self.run_btn.pack(side="left", padx=(0, 10))

        self.stop_btn = ttk.Button(btn_frame, text="停止", style="Stop.TButton", command=self._on_stop, width=10)
        self.stop_btn.state(["disabled"])
        self.stop_btn.pack(side="left")

        ttk.Button(btn_frame, text="打开输出目录", command=self._open_output_dir, width=12).pack(
            side="left", padx=(16, 0)
        )

        # --- 日志区域 ---
        log_frame = ttk.LabelFrame(ops_frame, text="运行日志", padding=(padx, 8))
        log_frame.pack(fill="both", expand=True)

        self.log_text = tk.Text(log_frame, height=10, font=FONT_LOG, state="normal", wrap="word",
                                bg="#fafafa", fg="#333333", padx=6, pady=4)
        scrollbar = tk.Scrollbar(log_frame, command=self.log_text.yview)
        self.log_text.config(yscrollcommand=scrollbar.set)
        self.log_text.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # 首次布局后再放 sash（此时 PanedWindow 高度才可靠）
        self.root.after_idle(self._init_main_pane_sash)

        # --- 状态栏 ---
        self.status_var = tk.StringVar(value="就绪")
        tk.Label(self.root, textvariable=self.status_var, font=("", 10), fg="#666666",
                 bd=1, relief="sunken", anchor="w", padx=8).pack(fill="x", side="bottom")

        self._on_date_type_change()

    def _init_main_pane_sash(self):
        """把配置区/底部区的初始高度切到更合理的比例；若尚未完成布局则稍后重试。"""
        try:
            h = self._main_pane.winfo_height()
            if h <= 1:
                self.root.after(50, self._init_main_pane_sash)
                return
            # 约 55% 给配置区；窗口很矮时保证底部至少 ~300px（按钮+日志+边距）
            y = int(h * 0.55)
            min_bottom = 300
            y = min(y, max(h - min_bottom, int(h * 0.30)))
            self._main_pane.sash_place(0, 0, y)
        except Exception:
            pass

    def _on_date_type_change(self):
        """根据日期类型切换显示范围/单日/无输入。"""
        dtype = self.date_type_var.get()
        self.range_frame.pack_forget()
        self.single_frame.pack_forget()
        if dtype == "range":
            self.range_frame.pack(fill="x", pady=4)
        elif dtype == "single":
            self.single_frame.pack(fill="x", pady=4)

    def _apply_preset(self, key: str):
        """一键切换日期类型；日常场景少点几下。"""
        today = datetime.now().strftime("%Y-%m-%d")
        if key == "today":
            self.date_type_var.set("today")
        elif key == "yesterday":
            self.date_type_var.set("yesterday")
        elif key == "all":
            self.date_type_var.set("all")
            if not self.data_start_entry.get().strip():
                self.data_start_entry.delete(0, "end")
                self.data_start_entry.insert(0, DEFAULT_DATA_START_DATE)
        elif key == "range":
            self.date_type_var.set("range")
            self.end_entry.delete(0, "end")
            self.end_entry.insert(0, today)
        self._on_date_type_change()
        self.mode_var.set("progress")
        self.format_var.set("excel")
        self.progress_layout_var.set("summary")
        self.status_var.set(f"已选快捷：{key}（可再勾选对标机构后点开始）")

    def _choose_output_dir(self):
        path = filedialog.askdirectory(
            title="选择输出目录",
            initialdir=self.output_dir_var.get() or default_output_directory(),
        )
        if path:
            self.output_dir_var.set(path)

    def _open_output_dir(self):
        path = (self.output_dir_var.get() or "").strip() or default_output_directory()
        Path(path).mkdir(parents=True, exist_ok=True)
        _open_path(path)

    def _select_common_benchmarks(self):
        if self.bench_search_var.get().strip():
            self.bench_search_var.set("")
        self._bench_sticky = set(COMMON_BENCHMARKS)
        self._render_benchmark_checkboxes()
        self.status_var.set(f"已勾选常用对标 {len(COMMON_BENCHMARKS)} 家")

    def _select_visible_benchmarks(self):
        if not hasattr(self, "_bench_sticky"):
            self._bench_sticky = set()
        for name, var in getattr(self, "bench_vars", {}).items():
            var.set(True)
            self._bench_sticky.add(name)

    def _clear_benchmarks(self):
        self._bench_sticky = set()
        for var in getattr(self, "bench_vars", {}).values():
            var.set(False)

    def _collect_prefs(self) -> Dict:
        self._sync_bench_sticky()
        selected = sorted(getattr(self, "_bench_sticky", set()) or [])
        return {
            "mode": self.mode_var.get(),
            "fund_type": self.fund_type_var.get(),
            "date_type": self.date_type_var.get(),
            "data_start": self.data_start_entry.get().strip(),
            "end_date": self.end_entry.get().strip(),
            "single_date": self.single_entry.get().strip(),
            "format": self.format_var.get(),
            "progress_layout": self.progress_layout_var.get(),
            "benchmarks": selected,
            "bench_start": self.bench_start_entry.get().strip(),
            "bench_end": self.bench_end_entry.get().strip(),
            "output_dir": self.output_dir_var.get().strip(),
            "open_when_done": bool(self.open_when_done_var.get()),
            "etf_full_history": bool(self.etf_full_history_var.get()),
        }

    def _apply_prefs(self):
        p = self._prefs or {}
        if not p:
            return
        try:
            if p.get("mode"):
                self.mode_var.set(p["mode"])
            if p.get("fund_type"):
                self.fund_type_var.set(p["fund_type"])
            if p.get("date_type"):
                self.date_type_var.set(p["date_type"])
                self._on_date_type_change()
            if p.get("data_start"):
                self.data_start_entry.delete(0, "end")
                self.data_start_entry.insert(0, p["data_start"])
            if p.get("end_date"):
                self.end_entry.delete(0, "end")
                self.end_entry.insert(0, p["end_date"])
            if p.get("single_date"):
                self.single_entry.delete(0, "end")
                self.single_entry.insert(0, p["single_date"])
            fmt = p.get("format")
            if fmt == "csv" and p.get("progress_layout") == "summary":
                # 旧偏好：csv+summary 实际出 Excel，迁到 excel
                fmt = "excel"
            if fmt in ("excel", "json", "both", "csv"):
                self.format_var.set(fmt)
            if p.get("progress_layout"):
                self.progress_layout_var.set(p["progress_layout"])
            if p.get("bench_start"):
                self.bench_start_entry.delete(0, "end")
                self.bench_start_entry.insert(0, p["bench_start"])
            if p.get("bench_end"):
                self.bench_end_entry.delete(0, "end")
                self.bench_end_entry.insert(0, p["bench_end"])
            if p.get("output_dir"):
                self.output_dir_var.set(p["output_dir"])
            if "open_when_done" in p:
                self.open_when_done_var.set(bool(p["open_when_done"]))
            if "etf_full_history" in p:
                self.etf_full_history_var.set(bool(p["etf_full_history"]))
            saved_benches = set(p.get("benchmarks") or [])
            if saved_benches:
                self._bench_sticky = set(saved_benches)
                self._render_benchmark_checkboxes()
        except Exception:
            pass

    def _on_close(self):
        save_prefs(self._collect_prefs())
        self.root.destroy()

    def _setup_logging(self):
        self.logger = logging.getLogger("csrc_scraper")
        self.logger.setLevel(logging.DEBUG)
        self.logger.handlers.clear()
        gui_handler = GUILogHandler(self.log_text)
        gui_handler.setLevel(logging.INFO)
        self.logger.addHandler(gui_handler)

    def _sync_bench_sticky(self):
        """把当前可见勾选合并进 sticky，避免搜索过滤时丢掉已选机构。"""
        if not hasattr(self, "_bench_sticky"):
            self._bench_sticky = set()
        for k, v in getattr(self, "bench_vars", {}).items():
            if v.get():
                self._bench_sticky.add(k)
            else:
                self._bench_sticky.discard(k)

    def _render_benchmark_checkboxes(self):
        # 清空并重绘
        if not hasattr(self, "_bench_container"):
            return
        self._sync_bench_sticky()
        for w in list(self._bench_container.winfo_children()):
            w.destroy()
        inst_map = getattr(self, "_bench_inst_map_cache", None) or load_benchmark_institutions()
        q = (getattr(self, "bench_search_var", None).get().strip() if hasattr(self, "bench_search_var") else "")
        if q:
            inst_map = {k: v for k, v in inst_map.items() if q in k or any(q in x for x in v)}
        if not hasattr(self, "_bench_sticky"):
            self._bench_sticky = set()
        self.bench_vars = {}

        max_per_row = 3
        row = 0
        col = 0
        for inst in inst_map.keys():
            if col >= max_per_row:
                row += 1
                col = 0
            v = tk.BooleanVar(value=inst in self._bench_sticky)
            self.bench_vars[inst] = v
            cb = tk.Checkbutton(
                self._bench_container,
                text=inst,
                variable=v,
                font=("", 10),
                command=self._sync_bench_sticky,
            )
            cb.grid(row=row, column=col, padx=(0, 10), pady=1, sticky="w")
            col += 1

        # 更新滚动区域
        if hasattr(self, "_bench_canvas"):
            try:
                self._bench_container.update_idletasks()
                self._bench_canvas.configure(scrollregion=self._bench_canvas.bbox("all"))
            except Exception:
                pass

    def _view_benchmark_institutions(self):
        """只读查看内置对标机构（修改请编辑 py 内 DEFAULT_BENCHMARK_INSTITUTIONS）。"""
        win = tk.Toplevel(self.root)
        win.title("对标机构清单（内置）")
        win.geometry("760x420" if sys.platform == "darwin" else "720x400")

        n = len(DEFAULT_BENCHMARK_INSTITUTIONS)
        tk.Label(
            win,
            text=(
                f"共 {n} 家机构，配置写在本文件 DEFAULT_BENCHMARK_INSTITUTIONS 中。"
                " 增删机构请用编辑器改该字典后重启 GUI，无需 json 配置文件。"
            ),
            font=("", 10),
            fg="#555",
            wraplength=720,
            justify="left",
        ).pack(anchor="w", padx=12, pady=(10, 4))

        txt = tk.Text(win, font=("Menlo" if sys.platform == "darwin" else "Consolas", 10), wrap="none")
        txt.pack(fill="both", expand=True, padx=12, pady=8)
        txt.insert("1.0", json.dumps(DEFAULT_BENCHMARK_INSTITUTIONS, ensure_ascii=False, indent=2))
        txt.configure(state="disabled")

        btns = tk.Frame(win)
        btns.pack(fill="x")
        tk.Button(btns, text="关闭", font=("", 10), command=win.destroy).pack(
            side="right", padx=12, pady=(0, 12)
        )

    def _validate_date(self, date_str):
        try:
            datetime.strptime(date_str, "%Y-%m-%d")
            return True
        except ValueError:
            return False

    def _get_data_start_date(self) -> str:
        return self.data_start_entry.get().strip() or DEFAULT_DATA_START_DATE

    def _on_run(self):
        dtype = self.date_type_var.get()
        mode = self.mode_var.get()
        fmt = self.format_var.get()
        data_start = self._get_data_start_date()

        if not self._validate_date(data_start):
            self.logger.error("开始日期格式错误，应为 YYYY-MM-DD（默认 %s）", DEFAULT_DATA_START_DATE)
            return

        if dtype == "range":
            start = data_start
            end = self.end_entry.get().strip()
            if not self._validate_date(end):
                self.logger.error("结束日期格式错误，应为 YYYY-MM-DD")
                return
            if start > end:
                self.logger.error("开始日期不能晚于结束日期")
                return
        elif dtype == "single":
            single = self.single_entry.get().strip()
            if not self._validate_date(single):
                self.logger.error("日期格式错误，应为 YYYY-MM-DD")
                return

        self.run_btn.config(state="disabled")
        self.stop_btn.config(state="normal")
        self.status_var.set("正在爬取...")

        thread = threading.Thread(target=self._do_scrape, daemon=True)
        thread.start()

    def _on_stop(self):
        self._stopped = True
        self.status_var.set("已请求停止")

    def _do_scrape(self):
        dtype = self.date_type_var.get()
        mode = self.mode_var.get()
        fmt = self.format_var.get()
        # excel / both / 旧版 csv+summary → 走多 Sheet 台账
        want_excel = fmt in ("excel", "both") or (
            fmt == "csv" and mode == "progress" and self.progress_layout_var.get() == "summary"
        )
        want_json = fmt in ("json", "both")
        want_plain_csv = fmt == "csv" and not want_excel
        output_dir = (self.output_dir_var.get() or "").strip() or default_output_directory()
        Path(output_dir).mkdir(parents=True, exist_ok=True)
        self._stopped = False
        self._last_exported_files = []
        start_time = time.time()
        save_prefs(self._collect_prefs())

        try:
            client = CSRCClient(delay=REQUEST_DELAY)
            # ETF 新主题需要“全历史”对比口径：缓存一份全历史 progress（仅用于主题基准），避免依赖用户选择的日期范围
            etf_theme_compare_base = None

            try:
                stats = client.get_stats()
                d = stats.get("data", {})
                self.logger.info("网站统计 - 审批总数: %s, 已办结: %s", d.get("countsl", "N/A"), d.get("countbj", "N/A"))
            except Exception as e:
                self.logger.warning("获取统计信息失败: %s", e)

            data_start = self._get_data_start_date()
            today_str = datetime.now().strftime("%Y-%m-%d")

            if dtype == "range":
                start_date = data_start
                end_date = self.end_entry.get().strip()
                self.logger.info("开始爬取 %s 至 %s 的 %s 数据", start_date, end_date, mode)
                data = self._fetch_date_range(client, mode, start_date, end_date)
                date_suffix = f"{start_date}_to_{end_date}"
            elif dtype == "single":
                date_str = self.single_entry.get().strip()
                self.logger.info("开始爬取 %s 的 %s 数据", date_str, mode)
                data = self._fetch_daily(client, mode, date_str)
                date_suffix = date_str
            elif dtype == "today":
                date_str = datetime.now().strftime("%Y-%m-%d")
                self.logger.info("开始爬取今日 (%s) 的 %s 数据", date_str, mode)
                data = self._fetch_daily(client, mode, date_str)
                date_suffix = date_str
            elif dtype == "yesterday":
                date_str = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
                self.logger.info("开始爬取昨日 (%s) 的 %s 数据", date_str, mode)
                data = self._fetch_daily(client, mode, date_str)
                date_suffix = date_str
            elif dtype == "all":
                end_date = today_str
                self.logger.info(
                    "开始爬取 %s 至 %s 的 %s 数据（全部历史，自开始日期起）",
                    data_start,
                    end_date,
                    mode,
                )
                data = self._fetch_date_range(client, mode, data_start, end_date)
                date_suffix = f"{data_start}_to_{end_date}"
                start_date = data_start
            else:
                return

            # 仅当导出 progress 汇总 Excel 时，才额外拉全历史用于 ETF 新主题对比（避免无谓耗时）
            etf_output_end_date = None
            if mode == "progress":
                if dtype == "range":
                    etf_output_end_date = end_date
                elif dtype in ("single", "today", "yesterday"):
                    etf_output_end_date = date_suffix
                elif dtype == "all":
                    etf_output_end_date = end_date

                if (
                    want_excel
                    and self.progress_layout_var.get() == "summary"
                    and bool(self.etf_full_history_var.get())
                ):
                    try:
                        self.logger.info("ETF新主题：开始抓取全历史数据用于对比（不受日期范围影响）...")
                        etf_theme_compare_base = self._fetch_all(client, "progress")
                        self.logger.info("ETF新主题：全历史共抓到 %d 条 progress 记录", len(etf_theme_compare_base))
                    except Exception as e:
                        self.logger.warning("ETF新主题：全历史对比数据抓取失败，将退回到当前日期范围：%s", e)
                elif want_excel and self.progress_layout_var.get() == "summary":
                    self.logger.info(
                        "ETF新主题：未勾选全历史对比，将仅用当前日期范围估算（更快）；需要精确口径请勾选后再跑"
                    )

            if self._stopped:
                self.logger.info("爬取已停止")
                return

            # 基金类型统计与筛选
            if mode in ("progress", "result"):
                from collections import Counter
                type_counts = Counter(r.get("fund_type", "未知") for r in data)
                self.logger.info("-" * 50)
                self.logger.info("基金类型分布:")
                for ft in FUND_CATEGORIES:
                    if ft in type_counts:
                        self.logger.info("  %s: %d 条", ft, type_counts[ft])
                if "其他" in type_counts:
                    self.logger.info("  其他: %d 条", type_counts["其他"])
                self.logger.info("-" * 50)

                ft_filter = self.fund_type_var.get()
                if ft_filter and ft_filter != "全部":
                    data = [r for r in data if r.get("fund_type") == ft_filter]
                    self.logger.info("已筛选为 %s，剩余 %d 条", ft_filter, len(data))

            base_name = export_output_base_name()
            exported_files = []

            if want_excel and mode == "progress" and self.progress_layout_var.get() == "summary":
                xlsx_path = os.path.join(output_dir, f"{base_name}.xlsx")
                leaders_rows = None
                leaders_fields = None
                try:
                    summary_rows = [summarize_progress(r) for r in data]
                    leaders_rows, leaders_fields = build_leaders_vs_gtht_rows(summary_rows)
                    self._sync_bench_sticky()
                    selected_benchmarks = sorted(getattr(self, "_bench_sticky", set()) or [])
                    benchmark_data = None
                    if selected_benchmarks:
                        bs = getattr(self, "bench_start_entry", None).get().strip()
                        be = getattr(self, "bench_end_entry", None).get().strip()
                        if (not self._validate_date(bs)) or (not self._validate_date(be)) or bs > be:
                            self.logger.error("对标时间范围格式错误，应为 YYYY-MM-DD 且起始<=结束")
                            return
                        # 对标数据始终按「对标时间」独立抓取，不受主范围/基金类型筛选影响
                        self.logger.info("对标机构：按时间范围 %s 至 %s 独立抓取...", bs, be)
                        benchmark_data = self._fetch_date_range(client, mode, bs, be)
                        self.logger.info("对标范围内共抓到 %d 条原始记录", len(benchmark_data))
                    benchmark_xlsx_path = (
                        os.path.join(output_dir, f"{base_name}_对标同业.xlsx")
                        if selected_benchmarks
                        else None
                    )
                    export_stats = export_progress_workbook_xlsx(
                        data,
                        xlsx_path,
                        leaders_rows=leaders_rows,
                        leaders_fieldnames=leaders_fields,
                        selected_benchmarks=selected_benchmarks,
                        benchmark_data=benchmark_data,
                        benchmark_xlsx_path=benchmark_xlsx_path,
                        etf_theme_compare_data=etf_theme_compare_base,
                        etf_output_end_date=etf_output_end_date,
                    )
                    exclude_count = int(export_stats.get("group_compare_exclude_count") or 0)
                    exclude_path = export_stats.get("group_compare_exclude_path") or group_compare_exclude_products_path()
                    if exclude_count:
                        self.logger.info(
                            "同批对比_全量：已排除 %d 个产品（仅本 Sheet，其它 Sheet 不受影响；名单：%s）",
                            exclude_count,
                            exclude_path,
                        )
                    bench_counts = export_stats.get("benchmark_counts") or {}
                    for inst_name, cnt in bench_counts.items():
                        self.logger.info("对标机构 %s：筛选命中 %d 条", inst_name, cnt)
                    self.logger.info(
                        "Excel 导出完成: %s（Sheet：进度汇总 / 债券类全量 / 对比_gtht_vs_leaders / 同批对比_全量 / ETF 等）",
                        xlsx_path,
                    )
                    exported_files.append(xlsx_path)
                    bench_xlsx = export_stats.get("benchmark_xlsx_path") or ""
                    if bench_xlsx:
                        self.logger.info(
                            "对标同业 Excel: %s（Sheet：%s）",
                            bench_xlsx,
                            " / ".join([f"对标_{x}" for x in selected_benchmarks]),
                        )
                        exported_files.append(bench_xlsx)
                except Exception as e:
                    self.logger.error("Excel 导出失败，将回退为 CSV: %s", e, exc_info=True)
                    csv_path = os.path.join(output_dir, f"{base_name}.csv")
                    count = export_csv(
                        data,
                        csv_path,
                        mode=mode,
                        progress_layout="summary",
                    )
                    self.logger.info("CSV 导出完成（回退）: %s (共 %d 条)", csv_path, count)
                    exported_files.append(csv_path)

                    if leaders_rows is not None and leaders_fields is not None:
                        try:
                            leaders_csv = os.path.join(output_dir, f"{base_name}_leaders_vs_gtht.csv")
                            Path(leaders_csv).parent.mkdir(parents=True, exist_ok=True)
                            with open(leaders_csv, "w", encoding="utf-8-sig", newline="") as f:
                                w = csv.DictWriter(f, fieldnames=leaders_fields)
                                w.writeheader()
                                w.writerows(leaders_rows)
                            self.logger.info("对比 CSV 导出完成（回退）: %s", leaders_csv)
                            exported_files.append(leaders_csv)
                        except Exception as e2:
                            self.logger.warning("对比 CSV 导出失败: %s", e2)
            elif want_plain_csv or (
                want_excel and not (mode == "progress" and self.progress_layout_var.get() == "summary")
            ):
                csv_path = os.path.join(output_dir, f"{base_name}.csv")
                count = export_csv(
                    data,
                    csv_path,
                    mode=mode,
                    progress_layout=self.progress_layout_var.get()
                    if mode == "progress"
                    else "summary",
                )
                self.logger.info("CSV 导出完成: %s (共 %d 条)", csv_path, count)
                exported_files.append(csv_path)

            if want_json:
                json_path = os.path.join(output_dir, f"{base_name}.json")
                export_json(data, json_path)
                self.logger.info("JSON 导出完成: %s (共 %d 条)", json_path, len(data))
                exported_files.append(json_path)

            self.logger.info("=" * 50)
            self.logger.info("爬取完成! 共 %d 条记录，耗时 %.2f 秒", len(data), time.time() - start_time)
            self.logger.info("输出目录: %s", os.path.abspath(output_dir))
            for f in exported_files:
                self.logger.info("  -> %s", f)
            self.logger.info("=" * 50)
            self._last_exported_files = list(exported_files)
            self.status_var.set(f"完成 - 共 {len(data)} 条")
            if self.open_when_done_var.get():
                self.root.after(200, lambda: self._reveal_outputs(output_dir, exported_files))

        except Exception as e:
            friendly = _format_network_error(e)
            self.logger.error("爬取失败: %s", friendly, exc_info=True)
            self.status_var.set("爬取失败")
        finally:
            self.run_btn.config(state="normal")
            self.stop_btn.config(state="disabled")

    def _reveal_outputs(self, output_dir: str, exported_files: List[str]):
        _open_path(output_dir)
        for f in exported_files:
            if f.lower().endswith(".xlsx"):
                _open_path(f)
                break

    def _fetch_all(self, client, mode):
        if mode == "progress":
            extractor = extract_progress_records
        elif mode == "result":
            extractor = extract_result_records
        else:
            extractor = extract_review_records

        all_records = []
        resp = client.get_progress_list(page_num=1) if mode == "progress" else (client.get_result_list(page_num=1) if mode == "result" else client.get_review_list(page_num=1))
        all_records.extend(extractor(resp))
        total = resp.get("data", {}).get("total", 0)
        total_pages = (total + PAGE_SIZE - 1) // PAGE_SIZE
        self.logger.info("总记录数: %d, 总页数: %d", total, total_pages)

        for page in range(2, total_pages + 1):
            if self._stopped:
                break
            self.logger.info("获取第 %d / %d 页...", page, total_pages)
            resp = client.get_progress_list(page_num=page) if mode == "progress" else (client.get_result_list(page_num=page) if mode == "result" else client.get_review_list(page_num=page))
            all_records.extend(extractor(resp))

        return all_records

    def _fetch_daily(self, client, mode, date_str):
        if mode == "progress":
            extractor = extract_progress_records
        elif mode == "result":
            extractor = extract_result_records
        else:
            extractor = extract_review_records

        all_records = []
        page = 1
        consecutive_empty = 0

        while consecutive_empty < 3:
            if self._stopped:
                break
            resp = client.get_progress_list(page_num=page) if mode == "progress" else (client.get_result_list(page_num=page) if mode == "result" else client.get_review_list(page_num=page))
            records = extractor(resp)
            filtered = [r for r in records if _match_date(r, date_str, mode)]
            all_records.extend(filtered)

            if not filtered:
                consecutive_empty = consecutive_empty + 1 if _has_records_before_start(records, mode, date_str) else 0
            else:
                consecutive_empty = 0

            if records and _has_records_before_start(records, mode, date_str):
                break
            page += 1
            if page > 500:
                break

        self.logger.info("%s 共找到 %d 条记录", date_str, len(all_records))
        return all_records

    def _fetch_date_range(self, client, mode, start_date, end_date):
        if mode == "progress":
            extractor = extract_progress_records
        elif mode == "result":
            extractor = extract_result_records
        else:
            extractor = extract_review_records

        all_records = []
        page = 1

        while True:
            if self._stopped:
                break
            resp = client.get_progress_list(page_num=page) if mode == "progress" else (client.get_result_list(page_num=page) if mode == "result" else client.get_review_list(page_num=page))
            records = extractor(resp)
            if not records:
                break

            filtered = [r for r in records if _match_date_range(r, mode, start_date, end_date)]
            all_records.extend(filtered)

            # 接口按时间倒序分页。若当前页所有记录都早于 start_date，后续页也必然更早，可立刻停止。
            if _has_records_before_start(records, mode, start_date):
                self.logger.info("已翻到早于起始日的数据（第 %d 页），停止翻页", page)
                break

            if page % 10 == 0:
                self.logger.info("已翻页到第 %d 页，范围内累计 %d 条", page, len(all_records))

            page += 1
            if page > 2000:
                break

        self.logger.info("%s 至 %s 共找到 %d 条记录", start_date, end_date, len(all_records))
        return all_records


def main():
    root = tk.Tk()
    app = CSRCScraperGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
