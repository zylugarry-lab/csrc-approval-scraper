---
name: 证监会每日审批爬虫
description: 在 WorkBuddy 对话里问清时间段与对标机构后，用无界面 CLI 快速导出证监会审批多 Sheet Excel。用户说证监会每日、审批进度爬虫、neris 台账时使用。默认不要弹 Tk GUI；不要无界面擅自抓全历史。
agent_created: true
---

# 证监会每日审批爬虫

## 推荐流程（默认）

不要先弹 GUI。在对话里用**两轮短问**收齐参数，再跑 CLI。

### 第 1 问：时间段

让用户选其一（默认建议「今日」）：

- 今日
- 昨日
- 某一天：`YYYY-MM-DD`
- 区间：`YYYY-MM-DD:YYYY-MM-DD`（慢，仅用户明确要长区间时）

### 第 2 问：对标机构

- 不要对标
- 常用对标（华夏/易方达/广发/富国/南方/嘉实/汇添富/博时/招商/鹏华）
- 或用户点名若干家（逗号分隔）

可用命令列出机构：

```bash
/usr/local/bin/python3 "${CODEBUDDY_SKILL_DIR}/run_cli.py" --list-bench
```

### 然后只跑 CLI（禁止代写爬虫逻辑）

依赖：

```bash
/usr/local/bin/python3 -m pip install -r "${CODEBUDDY_SKILL_DIR}/requirements.txt"
```

示例：

```bash
# 今日，不对标（最快）
/usr/local/bin/python3 "${CODEBUDDY_SKILL_DIR}/run_cli.py" --date today

# 昨日 + 常用对标
/usr/local/bin/python3 "${CODEBUDDY_SKILL_DIR}/run_cli.py" --date yesterday --bench-common

# 指定日 + 指定机构
/usr/local/bin/python3 "${CODEBUDDY_SKILL_DIR}/run_cli.py" --date 2026-10-28 --bench 华夏,易方达,广发
```

跑完把 `FILE=` 路径告诉用户。默认输出在 `${CODEBUDDY_SKILL_DIR}/输出/`。

## 速度规则

- CLI 已用 `pageSize=1000` + 有限并发翻页；`2024-01-01` 至今约十余秒级（视网络而定）。
- 默认 **不要** 加 `--etf-full-history`（全历史很慢）。
- 对标时间与主区间一致；不要默默从 2024 翻到现在。
- 用户说「全量 / 从 2024 / 日常全量」才用长区间；长区间仍比单日慢，但已明显加速。

## 何时才开 GUI

仅当用户明确说「打开 GUI / 图形界面 / 弹窗选」时：

```bash
/usr/local/bin/python3 "${CODEBUDDY_SKILL_DIR}/csrc_scraper_gui.py"
```

## 导出

- `证监会网页数据YYYYMMDD.xlsx`（多 Sheet）
- 有对标时另有 `_对标同业.xlsx`
