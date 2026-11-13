#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""区间双源拉取：审批进度公示 + 行政许可事项，可选某月获批清单。"""

from __future__ import annotations

import time
from collections import Counter
from datetime import datetime, timedelta
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from openpyxl import Workbook

from csrc_scraper_gui import (
    CSRCClient,
    build_leaders_vs_gtht_rows,
    export_progress_workbook_xlsx,
    normalize_task_name,
    summarize_progress,
)
from run_cli import _fetch_date_range, _setup_logger


RESULT_FIELDS = [
    "decision_date",
    "applicant",
    "matter_name",
    "fund_type",
    "reg_type",
    "description",
    "file_code",
    "application_id",
    "result_id",
]

PROG_FIELDS = [
    "decision_date",
    "application_date",
    "content",
    "fund_type",
    "reg_type",
    "task_name",
    "finish_date",
    "all_task_names",
    "all_finish_dates",
    "application_id",
]


def _write_rows(path: Path, sheet_name: str, rows: List[Dict], fields: List[str]) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name[:31]
    ws.append(fields)
    for r in rows:
        ws.append([r.get(f, "") for f in fields])
    wb.save(path)


def _decision_date_from_progress(rec: Dict) -> str:
    prog = rec.get("progress") or []
    for p in reversed(prog):
        name = normalize_task_name(p.get("task_name") or "")
        if name == "行政许可决定书":
            return (p.get("finish_date") or "").strip()
    dates = []
    for p in prog:
        name = normalize_task_name(p.get("task_name") or "")
        if name == "行政许可决定书":
            dates.append((p.get("finish_date") or "").strip())
    return dates[-1] if dates else ""


def _month_bounds(ym: str) -> Tuple[str, str]:
    """ym = YYYY-MM -> (YYYY-MM-01, YYYY-MM-last)."""
    start = f"{ym}-01"
    y, m = map(int, ym.split("-"))
    if m == 12:
        end = f"{y}-12-31"
    else:
        end = (datetime(y, m + 1, 1) - timedelta(days=1)).strftime("%Y-%m-%d")
    return start, end


def export_month_approvals(
    progress: List[Dict],
    results: List[Dict],
    month: str,
    out_path: Path,
) -> Dict:
    oct_s, oct_e = _month_bounds(month)
    oct_result = [
        r for r in results if oct_s <= (r.get("decision_date") or "") <= oct_e
    ]
    oct_result = sorted(
        oct_result,
        key=lambda r: (r.get("decision_date") or "", r.get("applicant") or ""),
    )

    oct_progress = []
    for r in progress:
        d = _decision_date_from_progress(r)
        if oct_s <= d <= oct_e:
            s = summarize_progress(r)
            s["decision_date"] = d
            oct_progress.append(s)
    oct_progress = sorted(
        oct_progress,
        key=lambda r: (r.get("decision_date") or "", r.get("content") or ""),
    )

    wb = Workbook()
    ws3 = wb.active
    ws3.title = f"{month}获批汇总"
    ws3.append(["获批月份", month])
    ws3.append(["窗口", f"{oct_s} ~ {oct_e}"])
    ws3.append([])
    ws3.append(["来源", "数量"])
    ws3.append(["行政许可结果（decision_date）", len(oct_result)])
    ws3.append(["审批进度（行政许可决定书 finish_date）", len(oct_progress)])
    ws3.append([])
    ws3.append(["按基金类型（行政许可结果）", "数量"])
    for k, v in Counter(r.get("fund_type") or "未知" for r in oct_result).most_common():
        ws3.append([k, v])
    ws3.append([])
    ws3.append(["按申请人（行政许可结果）", "数量"])
    for k, v in Counter(r.get("applicant") or "未知" for r in oct_result).most_common():
        ws3.append([k, v])

    ws1 = wb.create_sheet(f"{month}获批_行政许可结果")
    ws1.append(RESULT_FIELDS)
    for r in oct_result:
        ws1.append([r.get(f, "") for f in RESULT_FIELDS])

    ws2 = wb.create_sheet(f"{month}获批_审批进度决定书")
    ws2.append(PROG_FIELDS)
    for r in oct_progress:
        ws2.append([r.get(f, "") for f in PROG_FIELDS])

    wb.save(out_path)
    return {
        "result_count": len(oct_result),
        "progress_count": len(oct_progress),
        "by_fund_type": dict(Counter(r.get("fund_type") or "未知" for r in oct_result)),
    }


def run_period_job(
    start_date: str,
    end_date: str,
    output_dir: str,
    approve_month: Optional[str] = None,
    include_progress: bool = True,
    include_result: bool = True,
) -> Dict:
    logger = _setup_logger()
    out = Path(output_dir)
    out.mkdir(parents=True, exist_ok=True)
    client = CSRCClient(delay=0.05)
    t0 = time.time()
    files: List[str] = []
    progress: List[Dict] = []
    results: List[Dict] = []
    stats: Dict = {}

    tag = f"{start_date.replace('-', '')}_{end_date.replace('-', '')}"

    if include_progress:
        logger.info("拉取审批进度公示 %s~%s", start_date, end_date)
        progress = _fetch_date_range(client, "progress", start_date, end_date, logger)
        xlsx = out / f"审批进度公示_{tag}.xlsx"
        leaders_rows, leaders_fields = build_leaders_vs_gtht_rows(
            [summarize_progress(r) for r in progress]
        )
        export_progress_workbook_xlsx(
            progress,
            str(xlsx),
            leaders_rows=leaders_rows,
            leaders_fieldnames=leaders_fields,
            etf_output_end_date=end_date,
        )
        files.append(str(xlsx))
        stats["progress_count"] = len(progress)

    if include_result:
        logger.info("拉取行政许可事项 %s~%s", start_date, end_date)
        results = _fetch_date_range(client, "result", start_date, end_date, logger)
        xlsx = out / f"行政许可事项_{tag}.xlsx"
        rows = sorted(
            results,
            key=lambda r: (r.get("decision_date") or "", r.get("applicant") or ""),
            reverse=True,
        )
        _write_rows(xlsx, "行政许可结果", rows, RESULT_FIELDS)
        files.append(str(xlsx))
        stats["result_count"] = len(results)

    if approve_month:
        # 若只勾了某一侧，仍尽量基于已拉数据生成；缺一侧则补拉该月所在总区间
        if not include_progress:
            progress = _fetch_date_range(client, "progress", start_date, end_date, logger)
        if not include_result:
            results = _fetch_date_range(client, "result", start_date, end_date, logger)
        month_name = approve_month.replace("-", "")
        xlsx = out / f"{approve_month}获批基金清单.xlsx"
        month_stats = export_month_approvals(progress, results, approve_month, xlsx)
        files.append(str(xlsx))
        stats["approve_month"] = approve_month
        stats["month"] = month_stats
        logger.info(
            "%s 获批：结果 %d / 进度决定书 %d",
            approve_month,
            month_stats["result_count"],
            month_stats["progress_count"],
        )

    elapsed = time.time() - t0
    logger.info("区间双源完成: %d 个文件, %.1fs", len(files), elapsed)
    return {
        "files": files,
        "elapsed": elapsed,
        "start": start_date,
        "end": end_date,
        "stats": stats,
    }
