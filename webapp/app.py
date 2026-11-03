#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""证监会审批爬虫 · 网页版 GUI（FastAPI）

启动：
    pip3 install -r requirements.txt
    python3 app.py
浏览器打开 http://127.0.0.1:8787
"""

from __future__ import annotations

import io
import os
import shutil
import tempfile
import traceback
import zipfile
from datetime import datetime
from pathlib import Path
from typing import List, Optional

from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field

from csrc_scraper_gui import COMMON_BENCHMARKS, load_benchmark_institutions
from run_cli import run_job

ROOT = Path(__file__).resolve().parent
STATIC = ROOT / "static"
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)

app = FastAPI(title="证监会每日审批爬虫 · Web")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)
app.mount("/static", StaticFiles(directory=str(STATIC)), name="static")


class ScrapeRequest(BaseModel):
    date_type: str = Field("today", description="today|yesterday|single|range")
    single_date: Optional[str] = None
    start_date: Optional[str] = None
    end_date: Optional[str] = None
    fund_type: str = "全部"
    benchmarks: List[str] = Field(default_factory=list)
    etf_full_history: bool = False


def _date_spec(req: ScrapeRequest) -> str:
    dt = (req.date_type or "today").strip().lower()
    if dt == "today":
        return "today"
    if dt == "yesterday":
        return "yesterday"
    if dt == "single":
        if not req.single_date:
            raise HTTPException(400, "请填写目标日期")
        return req.single_date.strip()
    if dt == "range":
        if not req.start_date or not req.end_date:
            raise HTTPException(400, "请填写起止日期")
        a, b = req.start_date.strip(), req.end_date.strip()
        if a > b:
            raise HTTPException(400, "开始日期不能晚于结束日期")
        return f"{a}:{b}"
    raise HTTPException(400, f"未知日期类型: {req.date_type}")


@app.get("/", response_class=HTMLResponse)
def index():
    return HTMLResponse((STATIC / "index.html").read_text(encoding="utf-8"))


@app.get("/api/meta")
def meta():
    inst = load_benchmark_institutions()
    return {
        "institutions": list(inst.keys()),
        "common_benchmarks": [x for x in COMMON_BENCHMARKS if x in inst],
        "fund_types": ["全部", "货币型", "债券型", "混合型", "股票型", "REITs", "QDII", "ETF", "FOF"],
        "today": datetime.now().strftime("%Y-%m-%d"),
    }


@app.post("/api/scrape")
def scrape(req: ScrapeRequest):
    inst_map = load_benchmark_institutions()
    benches = []
    for name in req.benchmarks:
        if name not in inst_map:
            raise HTTPException(400, f"未知对标机构：{name}")
        benches.append(name)

    date_spec = _date_spec(req)
    tmp = tempfile.mkdtemp(prefix="csrc_web_")
    try:
        result = run_job(
            date_spec=date_spec,
            benches=benches,
            output_dir=tmp,
            etf_full_history=bool(req.etf_full_history),
            mode="progress",
        )
        files = [Path(p) for p in (result.get("files") or []) if p and Path(p).is_file()]
        if not files:
            files = sorted(Path(tmp).glob("*.xlsx"))
        if not files:
            raise HTTPException(500, "未生成 Excel")

        # 持久化一份到 output，避免临时目录被清
        stamp_dir = OUTPUT / datetime.now().strftime("%Y%m%d_%H%M%S")
        stamp_dir.mkdir(parents=True, exist_ok=True)
        saved = []
        for p in files:
            dest = stamp_dir / p.name
            shutil.copy2(p, dest)
            saved.append(dest)

        if len(saved) == 1:
            path = saved[0]
            return FileResponse(
                path,
                media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                filename=path.name,
            )

        mem = io.BytesIO()
        with zipfile.ZipFile(mem, "w", zipfile.ZIP_DEFLATED) as zf:
            for p in saved:
                zf.write(p, p.name)
        mem.seek(0)
        stamp = datetime.now().strftime("%Y%m%d")
        return StreamingResponse(
            mem,
            media_type="application/zip",
            headers={"Content-Disposition": f'attachment; filename="证监会网页数据{stamp}.zip"'},
        )
    except HTTPException:
        raise
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(500, f"抓取失败: {e}") from e
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def main():
    import uvicorn

    host = os.environ.get("HOST", "0.0.0.0")
    port = int(os.environ.get("PORT", "8787"))
    print(f"网页版 GUI: http://127.0.0.1:{port}")
    uvicorn.run("app:app", host=host, port=port, reload=False, log_level="info")


if __name__ == "__main__":
    main()
