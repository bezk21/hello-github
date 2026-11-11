#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
小红书投流工作台 · 本地数据底表同步脚本

用法（在工作台目录执行）：
    python3.11 sync.py --init          生成示例底表到 data/，用于第一次跑通链路
    python3.11 sync.py --inspect      列出所有底表的 sheet 与表头列名（字段对齐时用）
    python3.11 sync.py --once         单次同步：读取 data/ 底表 -> 生成 data.js
    python3.11 sync.py --watch 60     每 60 秒检查目录，文件有变化就自动同步（常驻）

定时方式二选一：
    A) 系统定时任务调用 --once（cron / launchd / Windows 任务计划程序）
    B) 直接常驻运行 --watch 60（最简单，无需配置系统任务）

依赖：openpyxl（pip3 install openpyxl）；CSV 底表仅用标准库。
数据不上传任何服务器，全部在本地完成。
"""
import argparse
import csv
import hashlib
import json
import sys
import time
from datetime import datetime
from pathlib import Path

BASE = Path(__file__).resolve().parent

# 数值型字段（字符串字段保持原样）
NUM_FIELDS = {
    "kpi": {"当期值", "目标值"},
    "trend": {"搜索量", "小红星GMV"},
    "accounts": {"消耗", "展现", "点击", "互动量", "加购", "成交GMV"},
    "finance": {"预算", "已消耗"},
    "todos": set(),
}
STATUS_VALUES = {"未开始", "进行中", "已完成", "已延期"}


def log(msg):
    line = "[%s] %s" % (datetime.now().strftime("%Y-%m-%d %H:%M:%S"), msg)
    print(line)
    try:
        cfg = load_config()
        with open(BASE / cfg.get("log_file", "sync.log"), "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass


def load_config():
    with open(BASE / "config.json", "r", encoding="utf-8") as f:
        return json.load(f)


def norm(s):
    return str(s or "").strip().lower().replace(" ", "").replace("\n", "").replace("\t", "")


def parse_num(v):
    if v is None or v == "":
        return 0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(",", "").replace("，", "").replace("¥", "").replace("￥", "").replace("%", "").replace(" ", "")
    mult = 1
    if s.endswith("亿"):
        mult, s = 1e8, s[:-1]
    elif s.endswith("万"):
        mult, s = 1e4, s[:-1]
    try:
        return float(s) * mult
    except ValueError:
        return 0


def source_files(cfg):
    d = BASE / cfg.get("input_dir", "data")
    if not d.exists():
        return []
    globs = [cfg.get("file_glob", "*.xlsx")] + cfg.get("extra_globs", [])
    out = []
    for g in globs:
        out.extend(d.glob(g))
    out = [p for p in out if p.is_file() and not p.name.startswith("~$") and p.suffix.lower() != ".js"]
    return sorted(set(out), key=lambda p: p.name)


def read_sheets(path):
    """返回 [(sheet名, [表头], [[单元格...], ...]), ...]"""
    suf = path.suffix.lower()
    if suf == ".csv":
        raw = None
        for enc in ("utf-8-sig", "gbk", "utf-8"):
            try:
                with open(path, "r", encoding=enc, newline="") as f:
                    raw = list(csv.reader(f))
                break
            except UnicodeDecodeError:
                continue
        if raw is None:
            return []
        rows = [r for r in raw if any(str(c).strip() for c in r)]
        if not rows:
            return []
        raw_headers = [str(c).strip() for c in rows[0]]
        return [(path.stem, [norm(c) for c in raw_headers], rows[1:], raw_headers)]
    try:
        from openpyxl import load_workbook
    except ImportError:
        log("缺少 openpyxl，请执行：pip3 install openpyxl")
        sys.exit(1)
    wb = load_workbook(path, data_only=True, read_only=True)
    result = []
    for ws in wb.worksheets:
        rows = []
        for r in ws.iter_rows(values_only=True):
            if any(c is not None and str(c).strip() != "" for c in r):
                rows.append(["" if c is None else c for c in r])
        if rows:
            raw_headers = [str(c).strip() for c in rows[0]]
            result.append((ws.title, [norm(c) for c in raw_headers], rows[1:], raw_headers))
    wb.close()
    return result


def build_colmap(headers, aliases):
    """返回 {工作台字段: 列索引}；headers 已 norm 过"""
    hidx = {}
    for i, h in enumerate(headers):
        if h and h not in hidx:
            hidx[h] = i
    cmap = {}
    for field, names in aliases.items():
        for n in names:
            key = norm(n)
            if key in hidx:
                cmap[field] = hidx[key]
                break
    return cmap


def collect(cfg, sheets, table_name):
    """从所有 sheet 中挑出最匹配该板块的一张，返回 (rows, sheet名, 命中数)"""
    spec = cfg["tables"].get(table_name, {})
    aliases = spec.get("columns", {})
    required = spec.get("required", [])
    best = None
    for sname, headers, body, _raw in sheets:
        cmap = build_colmap(headers, aliases)
        if any(r not in cmap for r in required):
            continue
        hits = len(cmap)
        if hits < cfg.get("min_match", 2):
            continue
        if best is None or hits > best[2]:
            best = (body, cmap, hits, sname)
    if best is None:
        return []
    body, cmap, _, sname = best
    nums = NUM_FIELDS.get(table_name, set())
    out = []
    for row in body:
        obj = {}
        for field, idx in cmap.items():
            v = row[idx] if idx < len(row) else ""
            obj[field] = parse_num(v) if field in nums else str(v).strip()
        if table_name == "todos":
            if not obj.get("事项"):
                continue
            if norm(obj.get("状态")) and obj["状态"] not in STATUS_VALUES:
                obj["状态"] = "未开始"
            for k in ("流程", "计划时间点", "负责人", "状态", "备注"):
                obj.setdefault(k, "")
        elif table_name == "accounts":
            if not obj.get("子账户"):
                continue
            for k in ("消耗", "展现", "点击", "互动量", "加购", "成交GMV"):
                obj.setdefault(k, 0)
        elif table_name == "kpi":
            if not obj.get("指标"):
                continue
            obj.setdefault("当期值", 0)
            obj.setdefault("目标值", 0)
        elif table_name == "trend":
            if not obj.get("月份"):
                continue
            obj.setdefault("搜索量", 0)
            obj.setdefault("小红星GMV", 0)
        elif table_name == "finance":
            if not obj.get("项目"):
                continue
            obj.setdefault("预算", 0)
            obj.setdefault("已消耗", 0)
        out.append(obj)
    return out


def sync_once(cfg, verbose=True):
    files = source_files(cfg)
    if not files:
        log("未找到数据底表，请把 Excel/CSV 放进 %s 目录（或先执行 --init）" % (BASE / cfg.get("input_dir", "data")))
        return False
    all_sheets = []
    for p in files:
        try:
            for sname, headers, body, _raw in read_sheets(p):
                all_sheets.append((sname, headers, body, _raw))
        except Exception as e:
            log("读取失败 %s: %s" % (p.name, e))
    if not all_sheets:
        log("底表均为空或无法解析")
        return False

    payload = {
        "updatedAt": datetime.now().isoformat(timespec="seconds"),
        "sourceFiles": [p.name for p in files],
        "tables": [],
    }
    summary = []
    for t in ("kpi", "trend", "accounts", "finance", "todos"):
        rows = collect(cfg, all_sheets, t)
        if not rows:
            summary.append("%s:跳过(未匹配)" % t)
            continue
        if t == "finance":
            budget = cfg.get("finance_default_budget", 0)
            kw = cfg.get("budget_row_keyword", "总预算")
            projects = []
            for r in rows:
                if kw in str(r.get("项目", "")):
                    budget = parse_num(r.get("预算", budget))
                else:
                    projects.append(r)
            payload["finance"] = {"总预算": budget, "projects": projects}
        else:
            payload[t] = rows
        payload["tables"].append(t)
        summary.append("%s:%d行" % (t, len(rows)))

    if not payload["tables"]:
        log("没有任何板块匹配成功。请执行 --inspect 查看底表列名，并修改 config.json 的 columns 别名")
        return False

    core = {k: v for k, v in payload.items() if k != "updatedAt"}
    payload["fingerprint"] = hashlib.md5(json.dumps(core, sort_keys=True, ensure_ascii=False).encode("utf-8")).hexdigest()[:12]
    js = "window.WB_DATA = " + json.dumps(payload, ensure_ascii=False, indent=2) + ";\n"
    (BASE / cfg.get("output_js", "data.js")).write_text(js, encoding="utf-8")
    log("同步完成 | 底表 %d 个 | %s | 输出 %s" % (len(files), " ".join(summary), cfg.get("output_js", "data.js")))
    if verbose:
        print("  提示：刷新工作台页面即可看到最新数据")
    return True


def cmd_init(cfg):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    d = BASE / cfg.get("input_dir", "data")
    d.mkdir(exist_ok=True)
    target = d / "示例底表.xlsx"
    sheets = {
        "总控台KPI": [["指标", "当期值", "目标值"],
                   ["AIPS人群(万)", 1280, 1500], ["TI人群(万)", 860, 1000],
                   ["搜索量(万)", 45.6, 60], ["小红星UV(万)", 12.4, 15],
                   ["小红星GMV(万元)", 386, 500], ["小红星ROI", 2.8, 3]],
        "趋势": [["月份", "搜索量", "小红星GMV"],
               ["7月", 32, 260], ["8月", 38, 305], ["9月", 41, 342], ["10月", 45.6, 386]],
        "聚光子账户": [["账户名称", "花费", "曝光量", "点击量", "互动量", "加购", "成交金额"],
                  ["账户A-Retail", 85000, 1250000, 38500, 9200, 1150, 238000],
                  ["账户B-Outlet", 62000, 980000, 31200, 6800, 820, 148000],
                  ["账户C-品牌词", 28000, 420000, 18600, 3900, 560, 96500]],
        "财务预算": [["项目", "预算", "已消耗"],
                 ["常规种草-Retail", 300000, 185000],
                 ["常规种草-Outlet", 200000, 92000],
                 ["【当月总预算】", 500000, 277000]],
        "待办": [["流程", "事项", "计划时间点", "负责人", "状态", "备注"],
               ["财务充值", "制作排期（确认充值时间点）", "每月25日前", "投放", "进行中", ""],
               ["财务充值", "内部下单", "每月28日前", "投放", "未开始", ""],
               ["财务充值", "PO申请提交", "次月3日前", "财务", "未开始", ""],
               ["财务充值", "客户审计填写实际消耗点", "次月5日前", "客户侧", "未开始", ""]],
    }
    wb = Workbook()
    fill = PatternFill("solid", fgColor="FFE9EAF0")
    for i, (name, rows) in enumerate(sheets.items()):
        ws = wb.active if i == 0 else wb.create_sheet()
        ws.title = name
        for r, row in enumerate(rows, 1):
            for c, v in enumerate(row, 1):
                cell = ws.cell(row=r, column=c, value=v)
                if r == 1:
                    cell.fill, cell.font = fill, Font(bold=True)
                    cell.alignment = Alignment(horizontal="center")
        for c in range(1, len(rows[0]) + 1):
            ws.column_dimensions[chr(64 + c) if c <= 26 else "A"].width = 24
        ws.freeze_panes = "A2"
    wb.save(target)
    print("已生成示例底表: %s" % target)
    print("注意：示例底表刻意使用了与你真实导出接近的列名（账户名称/花费/曝光量...），用于验证列映射能力。")
    print("下一步：python3.11 sync.py --once")


def cmd_inspect(cfg):
    files = source_files(cfg)
    if not files:
        print("data/ 目录下没有底表。先执行 --init 生成示例，或把你自己的 Excel 放进去。")
        return
    for p in files:
        print("\n文件: %s" % p.name)
        for sname, headers, body, raw in read_sheets(p):
            print("  sheet「%s」%d 行 | 表头: %s" % (sname, len(body), " / ".join(h for h in raw if h)))
    print("\n把上表里的列名填到 config.json 对应板块的 columns 别名数组中即可完成字段对齐。")


def snapshot(files):
    return {str(p): (p.stat().st_mtime_ns, p.stat().st_size) for p in files}


def cmd_watch(cfg, interval):
    d = BASE / cfg.get("input_dir", "data")
    d.mkdir(exist_ok=True)
    log("watch 启动 | 目录 %s | 每 %s 秒检查一次 | Ctrl+C 退出" % (d, interval))
    sync_once(cfg)
    last = snapshot(source_files(cfg))
    try:
        while True:
            time.sleep(interval)
            cur = snapshot(source_files(cfg))
            if cur != last:
                log("检测到底表变化，触发同步")
                sync_once(cfg, verbose=False)
                last = cur
    except KeyboardInterrupt:
        log("watch 已停止")


def main():
    ap = argparse.ArgumentParser(description="小红书投流工作台 · 本地数据同步")
    ap.add_argument("--init", action="store_true", help="生成示例底表")
    ap.add_argument("--inspect", action="store_true", help="查看底表列名，用于字段对齐")
    ap.add_argument("--once", action="store_true", help="单次同步")
    ap.add_argument("--watch", type=int, metavar="秒", help="常驻轮询目录，文件变化即同步")
    args = ap.parse_args()
    cfg = load_config()
    if args.init:
        cmd_init(cfg)
    elif args.inspect:
        cmd_inspect(cfg)
    elif args.watch:
        cmd_watch(cfg, max(5, args.watch))
    else:
        sync_once(cfg)


if __name__ == "__main__":
    main()
