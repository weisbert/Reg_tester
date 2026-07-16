#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
probe_current_data.py — 电流数据探查脚本（在测试机上运行，单文件，仅依赖 openpyxl）

对数据根目录（如 D:\\Excel）做一次结构+数据快照，输出一个 JSON：
  - 目录树（文件名/大小/修改时间）
  - 每个 xlsx 的 tab 列表
  - 仿真工作簿（Current_all_mode*.xlsx）的 Current_data 页全量数据
  - 各模式文件夹下 Result*.xlsx 的测量列（NO./Mode/Current/Temperature，按表头名定位），
    外加前 2 行数据的全部非空单元格（用于发现其他有用列）

用法：
  python probe_current_data.py D:\\Excel
  # 生成 D:\\Excel\\probe_dump.json，拷回开发机后用 mirror_from_probe.py 重建本地镜像

不改动任何原文件，只读。
"""
import argparse
import datetime
import glob
import json
import os
import re
import sys

import openpyxl

MAX_ROWS = 2000  # 每个 sheet 最多 dump 的数据行数，防意外超大表


def norm(v):
    return str(v).strip().lower() if v is not None else ""


def cell(row, idx):
    if idx is None or idx >= len(row):
        return None
    return row[idx]


def jval(v):
    if isinstance(v, (datetime.datetime, datetime.date, datetime.time)):
        return str(v)
    return v


def sheetnames_of(path):
    try:
        wb = openpyxl.load_workbook(path, read_only=True)
        names = list(wb.sheetnames)
        wb.close()
        return names, None
    except Exception as e:
        return None, f"{type(e).__name__}: {e}"


def dump_sim_sheet(path, sheet_hint="current_data"):
    """整页 dump（header + 全行，密集数组）。返回 dict 或 None。"""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        target = None
        for sn in wb.sheetnames:
            if norm(sn) == sheet_hint:
                target = sn
                break
        if target is None:
            return None
        ws = wb[target]
        rows = []
        for r in ws.iter_rows(values_only=True):
            if all(v is None for v in r):
                continue
            rows.append([jval(v) for v in r])
            if len(rows) > MAX_ROWS:
                break
        return {"file": os.path.basename(path), "sheet": target, "rows": rows,
                "truncated": len(rows) > MAX_ROWS}
    finally:
        wb.close()


def find_result_sheet(wb):
    """定位含 NO./Current 表头的 sheet。返回 (sheet名, 表头行号, 列映射dict) 或 (None,None,None)。"""
    for sn in wb.sheetnames:
        ws = wb[sn]
        for i, row in enumerate(ws.iter_rows(min_row=1, max_row=30, values_only=True), 1):
            no_col = mode_col = cur_col = temp_col = None
            for j, c in enumerate(row):
                n = norm(c)
                if n in ("no.", "no", "no．"):
                    no_col = j
                elif n == "mode" and mode_col is None:
                    mode_col = j
                elif n.startswith("current") and cur_col is None:
                    cur_col = j
                elif "temp" in n and temp_col is None:
                    temp_col = j
            if no_col is not None and cur_col is not None:
                if mode_col is None:
                    mode_col = no_col + 1
                return sn, i, {"no": no_col, "mode": mode_col, "cur": cur_col, "temp": temp_col}, row
        # 只在每个 sheet 的前 30 行找；找不到就试下一个 sheet
    return None, None, None, None


def dump_result(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        sn, hdr_idx, cols, hdr_row = find_result_sheet(wb)
        if sn is None:
            return {"file": os.path.basename(path), "error": "找不到含 NO./Current 表头的 sheet",
                    "sheets": list(wb.sheetnames)}
        ws = wb[sn]
        headers = [[j + 1, str(v)] for j, v in enumerate(hdr_row) if v not in (None, "")]
        rows, samples = [], []
        for i, row in enumerate(ws.iter_rows(min_row=hdr_idx + 1, values_only=True), hdr_idx + 1):
            no_raw = cell(row, cols["no"])
            label = cell(row, cols["mode"])
            cur = cell(row, cols["cur"])
            temp = cell(row, cols["temp"]) if cols["temp"] is not None else None
            if no_raw is None and label is None and cur is None:
                continue
            rows.append([i, jval(no_raw), jval(label), jval(cur), jval(temp)])
            if len(samples) < 2:
                samples.append([[j + 1,
                                 str(cell(hdr_row, j)) if cell(hdr_row, j) is not None else "",
                                 jval(v)] for j, v in enumerate(row) if v is not None])
            if len(rows) >= MAX_ROWS:
                break
        return {"file": os.path.basename(path), "sheet": sn, "header_row": hdr_idx,
                "key_cols_1based": {k: (v + 1 if v is not None else None) for k, v in cols.items()},
                "headers": headers,
                "row_fields": ["row_idx", "no_raw", "mode_label", "current", "temp"],
                "rows": rows, "sample_rows_full": samples}
    except Exception as e:
        return {"file": os.path.basename(path), "error": f"{type(e).__name__}: {e}"}
    finally:
        wb.close()


def main():
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    ap = argparse.ArgumentParser(description="电流数据探查：对数据根目录做 JSON 快照")
    ap.add_argument("root", nargs="?", default=r"D:\Excel", help="数据根目录（默认 D:\\Excel）")
    ap.add_argument("-o", "--out", help="输出 JSON 路径（默认 <root>\\probe_dump.json）")
    ap.add_argument("--sim-glob", default="Current_all_mode*.xlsx", help="仿真工作簿通配符")
    ap.add_argument("--result-glob", default="Result*.xlsx", help="实测文件通配符")
    args = ap.parse_args()

    root = os.path.abspath(args.root)
    if not os.path.isdir(root):
        raise SystemExit(f"[错误] 目录不存在: {root}")
    out_path = args.out or os.path.join(root, "probe_dump.json")

    dump = {"probe_version": 1,
            "probed_at": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "root": root, "tree": [], "workbooks": {}, "sim": [], "results": []}

    # 目录树 + 所有 xlsx 的 tab 列表
    for dirpath, dirnames, filenames in os.walk(root):
        rel_dir = os.path.relpath(dirpath, root)
        for f in filenames:
            full = os.path.join(dirpath, f)
            rel = os.path.normpath(os.path.join(rel_dir, f)) if rel_dir != "." else f
            try:
                st = os.stat(full)
                dump["tree"].append({"path": rel, "size": st.st_size,
                                     "mtime": datetime.datetime.fromtimestamp(st.st_mtime)
                                     .strftime("%Y-%m-%d %H:%M:%S")})
            except OSError:
                continue
            if f.lower().endswith((".xlsx", ".xlsm")) and not f.startswith("~$"):
                names, err = sheetnames_of(full)
                dump["workbooks"][rel] = names if names is not None else {"error": err}

    # 仿真工作簿
    for p in sorted(glob.glob(os.path.join(root, args.sim_glob))):
        if os.path.basename(p).startswith("~$"):
            continue
        try:
            d = dump_sim_sheet(p)
            if d:
                dump["sim"].append(d)
                print(f"[仿真] {os.path.basename(p)} -> {len(d['rows'])} 行")
            else:
                print(f"[仿真] {os.path.basename(p)} 无 Current_data 页，跳过")
        except Exception as e:
            dump["sim"].append({"file": os.path.basename(p), "error": f"{type(e).__name__}: {e}"})
            print(f"[仿真] {os.path.basename(p)} 读取失败: {e}")

    # 各子文件夹的 Result 文件
    for d in sorted(os.listdir(root)):
        sub = os.path.join(root, d)
        if not os.path.isdir(sub):
            continue
        for f in sorted(glob.glob(os.path.join(sub, args.result_glob))):
            if os.path.basename(f).startswith("~$"):
                continue
            r = dump_result(f)
            r["folder"] = d
            dump["results"].append(r)
            if "error" in r:
                print(f"[实测] {d}/{r['file']} -> 出错: {r['error']}")
            else:
                print(f"[实测] {d}/{r['file']} -> sheet={r['sheet']} {len(r['rows'])} 行 "
                      f"关键列={r['key_cols_1based']}")

    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(dump, f, ensure_ascii=False, indent=1, default=str)
    size_kb = os.path.getsize(out_path) / 1024
    print(f"\n[完成] 快照已写入 {out_path}（{size_kb:.0f} KB）")
    print("把这个 JSON 拷回开发机，用 mirror_from_probe.py 重建本地镜像。")


if __name__ == "__main__":
    main()
