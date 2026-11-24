#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""处理页面提交的导出请求，生成与网页版相同的 Excel。"""

from __future__ import annotations

import json
import os
import re
import sys
import time
import urllib.error
import urllib.request
from pathlib import Path

BIN = "https://extendsclass.com/api/json-storage/bin/fdabeee"
ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "approval" / "out"


def _parse_body(text: str) -> dict:
    cut = text.find("<!")
    if cut != -1:
        text = text[:cut]
    data = json.loads(text)
    if not isinstance(data, dict):
        return {"requests": []}
    if not isinstance(data.get("requests"), list):
        data["requests"] = []
    return data


def _read(req, timeout=25, attempts=3) -> bytes:
    last = None
    for attempt in range(1, attempts + 1):
        try:
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                return resp.read()
        except (TimeoutError, urllib.error.URLError, ConnectionError) as exc:
            last = exc
            if attempt < attempts:
                time.sleep(2 * attempt)
    raise last


def load_bin() -> dict:
    req = urllib.request.Request(
        BIN + "?t=" + str(time.time()),
        headers={"User-Agent": "csrc-export"},
    )
    return _parse_body(_read(req).decode("utf-8", "replace"))


def save_bin(data: dict) -> None:
    body = json.dumps({"requests": data.get("requests") or []}, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(
        BIN,
        data=body,
        method="PUT",
        headers={"Content-Type": "application/json", "User-Agent": "csrc-export"},
    )
    _read(req)


def pending_items(data: dict) -> list:
    found = []
    for item in data["requests"]:
        if not isinstance(item, dict) or item.get("done"):
            continue
        request_id = str(item.get("id") or "")
        if not re.fullmatch(r"[a-z0-9]{8,40}", request_id):
            item["done"] = True
            item["error"] = "请求编号不对"
            continue
        if (OUT / f"{request_id}.json").is_file():
            item["done"] = True
            continue
        found.append(item)
    return found[:3]


def _github_output(key: str, value: str) -> None:
    path = os.environ.get("GITHUB_OUTPUT")
    if not path:
        return
    with open(path, "a", encoding="utf-8") as handle:
        handle.write(f"{key}={value}\n")


def main() -> int:
    os.environ["TZ"] = "Asia/Shanghai"
    if hasattr(time, "tzset"):
        time.tzset()
    try:
        data = load_bin()
    except (TimeoutError, urllib.error.URLError, ConnectionError) as exc:
        if "--peek" in sys.argv:
            print(f"队列暂时读不到，这轮跳过：{exc}")
            _github_output("pending", "0")
            return 0
        print(f"队列读不到：{exc}", file=sys.stderr)
        return 1
    todo = pending_items(data)
    if "--peek" in sys.argv:
        _github_output("pending", "1" if todo else "0")
        print(f"pending={len(todo)}")
        return 0
    if not todo:
        print("没有待生成的请求")
        _github_output("produced", "0")
        return 0

    sys.path.insert(0, str(Path(__file__).resolve().parent))
    from github_export import export_to_out

    OUT.mkdir(parents=True, exist_ok=True)
    produced = 0
    for item in todo:
        request_id = str(item["id"])
        payload = item.get("payload")
        try:
            if not isinstance(payload, dict):
                raise ValueError("参数格式不对")
            export_to_out(payload, request_id, OUT)
            item["done"] = True
            item.pop("error", None)
            produced += 1
        except Exception as exc:
            tries = int(item.get("tries") or 0) + 1
            item["tries"] = tries
            item["error"] = str(exc)
            if tries >= 2:
                item["done"] = True
                (OUT / f"{request_id}.json").write_text(
                    json.dumps({"error": str(exc)}, ensure_ascii=False),
                    encoding="utf-8",
                )
            print(f"导出失败 {request_id}: {exc}", file=sys.stderr)
    data["requests"] = data["requests"][-40:]
    save_bin(data)
    publish = produced > 0 or any(item.get("done") and item.get("error") for item in todo)
    _github_output("produced", "1" if publish else "0")
    print(f"produced={produced}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
