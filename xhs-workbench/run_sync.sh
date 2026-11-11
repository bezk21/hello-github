#!/usr/bin/env bash
# 定时任务入口（macOS / Linux）。cron 示例：
#   */30 * * * * cd /你的路径/xhs-workbench && bash run_sync.sh >> sync_cron.log 2>&1
cd "$(dirname "$0")" || exit 1
PY=python3.11
command -v "$PY" >/dev/null 2>&1 || PY=python3
"$PY" sync.py --once
