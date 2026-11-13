#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""这台电脑开着时，发现同事提交的导出请求就启动生成。"""

from __future__ import annotations

import json
import subprocess
import sys
import time
import urllib.request

BIN = "https://extendsclass.com/api/json-storage/bin/fdabeee"
REPO = "zylugarry-lab/csrc-approval-scraper"


def pending() -> int:
    req = urllib.request.Request(
        BIN + "?t=" + str(time.time()),
        headers={"User-Agent": "csrc-watch"},
    )
    with urllib.request.urlopen(req, timeout=30) as resp:
        text = resp.read().decode("utf-8", "replace")
    cut = text.find("<!")
    if cut != -1:
        text = text[:cut]
    data = json.loads(text)
    items = data.get("requests") if isinstance(data, dict) else []
    return sum(1 for item in items if isinstance(item, dict) and not item.get("done"))


def busy() -> bool:
    out = subprocess.check_output(
        [
            "gh",
            "run",
            "list",
            "-R",
            REPO,
            "--workflow",
            "csrc-export.yml",
            "--limit",
            "1",
            "--json",
            "status",
        ],
        text=True,
    )
    runs = json.loads(out or "[]")
    if not runs:
        return False
    return runs[0].get("status") in {"queued", "in_progress", "waiting", "pending", "requested"}


def main() -> int:
    if pending() < 1 or busy():
        return 0
    subprocess.check_call(["gh", "workflow", "run", "csrc-export.yml", "-R", REPO])
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(exc, file=sys.stderr)
        raise SystemExit(1)
