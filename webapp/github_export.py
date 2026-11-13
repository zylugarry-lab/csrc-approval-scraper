#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""从 GitHub issue 正文里的参数生成与网页版相同的 Excel。"""

from __future__ import annotations

import json
import os
import re
import shutil
import sys
import time
import zipfile
from datetime import datetime
from pathlib import Path

os.environ["TZ"] = "Asia/Shanghai"
if hasattr(time, "tzset"):
    time.tzset()

from csrc_scraper_gui import load_benchmark_institutions
from period_job import run_period_job
from run_cli import run_job

FUND_TYPES = ["全部", "货币型", "债券型", "混合型", "股票型", "REITs", "QDII", "ETF", "FOF"]
DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")
MONTH_RE = re.compile(r"^\d{4}-\d{2}$")


def parse_body(raw: str) -> dict:
    if len(raw) > 20000:
        raise ValueError("请求内容过长")
    match = re.search(r"```csrc\s*(\{.*?\})\s*```", raw, re.S)
    if not match:
        raise ValueError("没有找到导出参数")
    data = json.loads(match.group(1))
    if not isinstance(data, dict):
        raise ValueError("参数格式不对")
    return data


def _date(value: str, label: str) -> str:
    text = (value or "").strip()
    if not DATE_RE.match(text):
        raise ValueError(f"{label}应为 YYYY-MM-DD")
    datetime.strptime(text, "%Y-%m-%d")
    return text


def _benches(req: dict) -> list:
    inst = load_benchmark_institutions()
    raw = req.get("benchmarks") or []
    if not isinstance(raw, list):
        raise ValueError("对标机构格式不对")
    out = []
    for name in raw:
        if not isinstance(name, str) or name not in inst:
            raise ValueError(f"未知对标机构：{name}")
        if name not in out:
            out.append(name)
    return out


def _collect(result: dict) -> list:
    files = [Path(p) for p in (result.get("files") or []) if p and Path(p).is_file()]
    if not files:
        raise RuntimeError("未生成 Excel")
    return files


def _daily(req: dict, work: Path) -> list:
    date_type = (req.get("date_type") or "today").strip()
    if date_type == "today":
        spec = "today"
    elif date_type == "yesterday":
        spec = "yesterday"
    elif date_type == "single":
        spec = _date(req.get("single_date"), "日期")
    elif date_type == "range":
        start = _date(req.get("start_date"), "开始日期")
        end = _date(req.get("end_date"), "结束日期")
        if start > end:
            raise ValueError("开始日期不能晚于结束日期")
        spec = f"{start}:{end}"
    else:
        raise ValueError("未知日期类型")
    fund = req.get("fund_type") or "全部"
    if fund not in FUND_TYPES:
        raise ValueError("未知基金类型")
    result = run_job(
        date_spec=spec,
        benches=_benches(req),
        output_dir=str(work),
        etf_full_history=bool(req.get("etf_full_history")),
        mode="progress",
        fund_type=None if fund == "全部" else fund,
    )
    return _collect(result)


def _period(req: dict, work: Path) -> list:
    start = _date(req.get("start_date"), "开始日期")
    end = _date(req.get("end_date"), "结束日期")
    if start > end:
        raise ValueError("开始日期不能晚于结束日期")
    month = (req.get("approve_month") or "").strip() or None
    if month and not MONTH_RE.match(month):
        raise ValueError("获批月份应为 YYYY-MM")
    include_progress = bool(req.get("include_progress", True))
    include_result = bool(req.get("include_result", True))
    if not include_progress and not include_result:
        raise ValueError("请至少勾选一个数据源")
    result = run_period_job(
        start_date=start,
        end_date=end,
        output_dir=str(work),
        approve_month=month,
        include_progress=include_progress,
        include_result=include_result,
    )
    return _collect(result)


def run(req: dict, out_dir: Path) -> Path:
    kind = req.get("kind")
    work = out_dir / "_files"
    if work.exists():
        shutil.rmtree(work)
    work.mkdir(parents=True)
    if kind == "daily":
        files = _daily(req, work)
        zip_name = "证监会网页数据.zip"
    elif kind == "period":
        files = _period(req, work)
        zip_name = "区间双源.zip"
    else:
        raise ValueError("未知导出类型")

    out_dir.mkdir(parents=True, exist_ok=True)
    if len(files) == 1:
        dest = out_dir / files[0].name
        shutil.copy2(files[0], dest)
    else:
        dest = out_dir / zip_name
        with zipfile.ZipFile(dest, "w", zipfile.ZIP_DEFLATED) as bundle:
            for path in files:
                bundle.write(path, path.name)
    shutil.rmtree(work, ignore_errors=True)
    return dest


def main() -> int:
    out = Path(sys.argv[1] if len(sys.argv) > 1 else "export-out")
    try:
        produced = run(parse_body(os.environ.get("ISSUE_BODY", "")), out)
    except Exception as exc:
        print(f"导出失败: {exc}", file=sys.stderr)
        return 1
    print(produced)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
