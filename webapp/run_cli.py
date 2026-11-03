#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""无界面快速爬取：供 WorkBuddy 问答选参后调用。

日常默认：今日 + 不抓 ETF 全历史 + 对标时间与主区间一致（避免从 2024 翻到现在）。
"""

from __future__ import annotations

import argparse
import logging
import os
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timedelta
from pathlib import Path
from typing import Dict, List, Optional, Tuple

# 与 GUI 同目录
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from csrc_scraper_gui import (  # noqa: E402
    COMMON_BENCHMARKS,
    CSRCClient,
    DEFAULT_DATA_START_DATE,
    _has_records_before_start,
    _match_date,
    _match_date_range,
    build_leaders_vs_gtht_rows,
    default_output_directory,
    export_output_base_name,
    export_progress_workbook_xlsx,
    extract_progress_records,
    extract_result_records,
    extract_review_records,
    load_benchmark_institutions,
    summarize_progress,
)

# CLI 加速参数（接口实测支持 pageSize=1000）
CLI_PAGE_SIZE = 1000
CLI_DELAY = 0.05
CLI_WORKERS = 5


def _setup_logger() -> logging.Logger:
    logger = logging.getLogger("csrc_cli")
    logger.setLevel(logging.INFO)
    if not logger.handlers:
        h = logging.StreamHandler(sys.stdout)
        h.setFormatter(logging.Formatter("[%(asctime)s] %(levelname)s | %(message)s", "%H:%M:%S"))
        logger.addHandler(h)
    return logger


def _parse_date_spec(spec: str) -> tuple[str, str, str]:
    """返回 (kind, start, end)。kind: today|yesterday|single|range。"""
    s = (spec or "today").strip().lower()
    today = datetime.now().strftime("%Y-%m-%d")
    yesterday = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
    if s in ("today", "今日", "今天"):
        return "today", today, today
    if s in ("yesterday", "昨日", "昨天"):
        return "yesterday", yesterday, yesterday
    if ":" in s or "至" in s or "到" in s:
        for sep in (":", "至", "到"):
            if sep in s:
                a, b = [x.strip() for x in s.split(sep, 1)]
                datetime.strptime(a, "%Y-%m-%d")
                datetime.strptime(b, "%Y-%m-%d")
                if a > b:
                    raise ValueError(f"开始日期不能晚于结束日期: {a} > {b}")
                return "range", a, b
    datetime.strptime(s, "%Y-%m-%d")
    return "single", s, s


def _parse_benches(raw: Optional[str], use_common: bool) -> List[str]:
    inst_map = load_benchmark_institutions()
    if use_common:
        return [x for x in COMMON_BENCHMARKS if x in inst_map]
    if not raw or not raw.strip():
        return []
    if raw.strip() in ("常用", "common", "常用对标"):
        return [x for x in COMMON_BENCHMARKS if x in inst_map]
    out = []
    for part in raw.replace("，", ",").replace("、", ",").split(","):
        name = part.strip()
        if not name:
            continue
        if name in inst_map:
            out.append(name)
            continue
        # 模糊：简称命中
        hits = [k for k in inst_map if name in k or k in name]
        if len(hits) == 1:
            out.append(hits[0])
        elif name in hits:
            out.append(name)
        else:
            raise ValueError(f"未知对标机构「{name}」。可选: {', '.join(inst_map.keys())}")
    # 去重保序
    seen = set()
    uniq = []
    for x in out:
        if x not in seen:
            seen.add(x)
            uniq.append(x)
    return uniq


def _extractor(mode: str):
    if mode == "progress":
        return extract_progress_records
    if mode == "result":
        return extract_result_records
    return extract_review_records


def _fetch_one_page(mode: str, page_num: int, page_size: int) -> Tuple[int, List[Dict]]:
    """独立 Session 拉一页，便于并发。"""
    client = CSRCClient(delay=0)
    if mode == "progress":
        resp = client.get_progress_list(page_num=page_num, page_size=page_size)
    elif mode == "result":
        resp = client.get_result_list(page_num=page_num, page_size=page_size)
    else:
        resp = client.get_review_list(page_num=page_num, page_size=page_size)
    return page_num, _extractor(mode)(resp)


def _fetch_daily(client: CSRCClient, mode: str, date_str: str, logger: logging.Logger) -> List[Dict]:
    all_records: List[Dict] = []
    page = 1
    consecutive_empty = 0
    while consecutive_empty < 3:
        resp = (
            client.get_progress_list(page_num=page, page_size=CLI_PAGE_SIZE)
            if mode == "progress"
            else (
                client.get_result_list(page_num=page, page_size=CLI_PAGE_SIZE)
                if mode == "result"
                else client.get_review_list(page_num=page, page_size=CLI_PAGE_SIZE)
            )
        )
        records = _extractor(mode)(resp)
        filtered = [r for r in records if _match_date(r, date_str, mode)]
        all_records.extend(filtered)
        if not filtered:
            consecutive_empty = consecutive_empty + 1 if _has_records_before_start(records, mode, date_str) else 0
        else:
            consecutive_empty = 0
        if records and _has_records_before_start(records, mode, date_str):
            break
        page += 1
        if page > 200:
            break
        if page % 3 == 0:
            logger.info("单日翻页中…第 %d 页，已命中 %d 条", page, len(all_records))
        if CLI_DELAY:
            time.sleep(CLI_DELAY)
    logger.info("%s 共找到 %d 条记录", date_str, len(all_records))
    return all_records


def _fetch_date_range(
    client: CSRCClient,
    mode: str,
    start_date: str,
    end_date: str,
    logger: logging.Logger,
) -> List[Dict]:
    """大 pageSize + 有限并发翻页；按页序处理，一旦整页早于 start 即停。"""
    del client  # 并发时每页独立 Session
    all_records: List[Dict] = []
    page = 1
    stop = False
    max_page = 500
    while not stop and page <= max_page:
        batch = list(range(page, min(page + CLI_WORKERS, max_page + 1)))
        t_batch = time.time()
        results: Dict[int, List[Dict]] = {}
        with ThreadPoolExecutor(max_workers=CLI_WORKERS) as pool:
            futs = [pool.submit(_fetch_one_page, mode, p, CLI_PAGE_SIZE) for p in batch]
            for fut in as_completed(futs):
                pnum, records = fut.result()
                results[pnum] = records
        logger.info(
            "并发拉取第 %d–%d 页完成（%.1fs）",
            batch[0],
            batch[-1],
            time.time() - t_batch,
        )
        for p in sorted(results):
            records = results[p]
            if not records:
                stop = True
                break
            filtered = [r for r in records if _match_date_range(r, mode, start_date, end_date)]
            all_records.extend(filtered)
            if _has_records_before_start(records, mode, start_date):
                logger.info("已翻到早于起始日的数据（第 %d 页），停止翻页", p)
                stop = True
                break
        page += CLI_WORKERS
        if CLI_DELAY and not stop:
            time.sleep(CLI_DELAY)

    logger.info("%s 至 %s 共找到 %d 条记录", start_date, end_date, len(all_records))
    return all_records


def run_job(
    date_spec: str = "today",
    benches: Optional[List[str]] = None,
    output_dir: Optional[str] = None,
    etf_full_history: bool = False,
    mode: str = "progress",
) -> Dict:
    logger = _setup_logger()
    kind, start, end = _parse_date_spec(date_spec)
    benches = benches or []
    out_dir = output_dir or default_output_directory()
    Path(out_dir).mkdir(parents=True, exist_ok=True)

    logger.info(
        "参数: 日期=%s (%s~%s), 对标=%s, ETF全历史=%s, pageSize=%d, workers=%d",
        date_spec,
        start,
        end,
        benches or "无",
        etf_full_history,
        CLI_PAGE_SIZE,
        CLI_WORKERS,
    )
    t0 = time.time()
    client = CSRCClient(delay=CLI_DELAY)

    if kind in ("today", "yesterday", "single"):
        data = _fetch_daily(client, mode, start, logger)
        etf_end = start
    else:
        data = _fetch_date_range(client, mode, start, end, logger)
        etf_end = end

    etf_theme_compare_base = None
    if etf_full_history:
        logger.info("ETF新主题：抓取全历史（较慢）...")
        # 复用区间抓取自默认起点
        etf_theme_compare_base = _fetch_date_range(
            client, "progress", DEFAULT_DATA_START_DATE, datetime.now().strftime("%Y-%m-%d"), logger
        )
    else:
        logger.info("跳过 ETF 全历史对比（更快）")

    benchmark_data = None
    if benches:
        # 对标与主区间一致，避免默默从 2024 翻到现在
        logger.info("对标机构按主区间 %s~%s 抓取: %s", start, end, ", ".join(benches))
        if start == end:
            benchmark_data = _fetch_daily(client, mode, start, logger)
        else:
            benchmark_data = _fetch_date_range(client, mode, start, end, logger)

    base_name = export_output_base_name()
    xlsx_path = os.path.join(out_dir, f"{base_name}.xlsx")
    leaders_rows, leaders_fields = build_leaders_vs_gtht_rows([summarize_progress(r) for r in data])
    bench_xlsx = os.path.join(out_dir, f"{base_name}_对标同业.xlsx") if benches else None

    export_stats = export_progress_workbook_xlsx(
        data,
        xlsx_path,
        leaders_rows=leaders_rows,
        leaders_fieldnames=leaders_fields,
        selected_benchmarks=benches or None,
        benchmark_data=benchmark_data,
        benchmark_xlsx_path=bench_xlsx,
        etf_theme_compare_data=etf_theme_compare_base,
        etf_output_end_date=etf_end,
    )
    elapsed = time.time() - t0
    files = [xlsx_path]
    if export_stats.get("benchmark_xlsx_path"):
        files.append(export_stats["benchmark_xlsx_path"])

    logger.info("完成: %d 条, 耗时 %.1f 秒", len(data), elapsed)
    for f in files:
        logger.info("  -> %s", f)
    return {
        "count": len(data),
        "elapsed": elapsed,
        "files": files,
        "start": start,
        "end": end,
        "benches": benches,
    }


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="证监会审批进度爬虫（无界面快速版）")
    parser.add_argument(
        "--date",
        default="today",
        help="today / yesterday / YYYY-MM-DD / YYYY-MM-DD:YYYY-MM-DD",
    )
    parser.add_argument("--bench", default="", help="对标机构，逗号分隔；或写「常用」")
    parser.add_argument("--bench-common", action="store_true", help="勾选常用对标机构")
    parser.add_argument("--out", default="", help="输出目录，默认技能目录/输出")
    parser.add_argument("--etf-full-history", action="store_true", help="ETF新主题额外抓全历史（很慢）")
    parser.add_argument("--list-bench", action="store_true", help="列出可选对标机构后退出")
    args = parser.parse_args(argv)

    if args.list_bench:
        for k in load_benchmark_institutions():
            print(k)
        return 0

    try:
        benches = _parse_benches(args.bench, args.bench_common)
        result = run_job(
            date_spec=args.date,
            benches=benches,
            output_dir=args.out or None,
            etf_full_history=bool(args.etf_full_history),
        )
    except Exception as e:
        print(f"错误: {e}", file=sys.stderr)
        return 1

    print("OK")
    print(f"COUNT={result['count']}")
    print(f"ELAPSED={result['elapsed']:.1f}")
    for f in result["files"]:
        print(f"FILE={f}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
