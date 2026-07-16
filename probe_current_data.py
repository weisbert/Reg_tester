#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
probe_current_data.py — 电流数据探查脚本 v2（在数据机上运行，单文件，仅依赖 openpyxl）

对数据根目录做一次结构+数据快照，输出一个 JSON：
  - 目录树（文件名/大小/修改时间），全递归，与层级无关
  - 每个 xlsx 的 tab 列表
  - 仿真长表：不认文件名，扫描所有工作簿的所有 tab，表头含 ID+Mode+Current
    （且有 simulation/Unit/Tier 之一）即整页 dump
  - 实测：任意层级下匹配 Result*.xlsx 的文件，按表头定位 NO./Mode/Current/Temperature 列
    并 dump；其他工作簿里长得像实测表的 tab 只打标记不 dump

用法：
  python probe_current_data.py <数据根目录>
  # 生成 <根目录>\probe_dump.json，拷回开发机后用 mirror_from_probe.py 重建本地镜像

不改动任何原文件，只读。
"""
import argparse
import datetime
import fnmatch
import json
import os
import sys

import openpyxl

MAX_ROWS = 2000    # 每个 sheet 最多 dump 的数据行数，防意外超大表
SCAN_ROWS = 30     # 找表头时扫描的行数


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


def match_sim_header(row):
    """仿真长表表头：ID + Mode + Current*，且有 simulation/Unit/Tier 佐证。"""
    names = [norm(c) for c in row]
    return ("id" in names and "mode" in names
            and any(n.startswith("current") for n in names)
            and (any(n.startswith("simulation") for n in names)
                 or "unit" in names or "tier" in names))


def match_result_header(row):
    """实测表表头：NO. + Current*。返回列映射 dict 或 None。"""
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
    if no_col is None or cur_col is None:
        return None
    if mode_col is None:
        mode_col = no_col + 1
    return {"no": no_col, "mode": mode_col, "cur": cur_col, "temp": temp_col}


def scan_workbook(path):
    """扫一个工作簿：tab 列表 + 每个 tab 是否命中 仿真/实测 表头。"""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        info = {"sheets": list(wb.sheetnames), "sim_sheets": [], "result_sheets": {}}
        for sn in wb.sheetnames:
            ws = wb[sn]
            for i, row in enumerate(ws.iter_rows(min_row=1, max_row=SCAN_ROWS,
                                                 values_only=True), 1):
                if match_sim_header(row):
                    info["sim_sheets"].append({"sheet": sn, "header_row": i})
                    break
                cols = match_result_header(row)
                if cols is not None:
                    info["result_sheets"][sn] = {
                        "header_row": i,
                        "key_cols_1based": {k: (v + 1 if v is not None else None)
                                            for k, v in cols.items()}}
                    break
        return info, None
    except Exception as e:
        return None, f"{type(e).__name__}: {e}"
    finally:
        try:
            wb.close()
        except Exception:
            pass


def dump_sheet_rows(path, sheet):
    """整页 dump（跳过全空行，密集数组）。"""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb[sheet]
        rows = []
        for r in ws.iter_rows(values_only=True):
            if all(v is None for v in r):
                continue
            rows.append([jval(v) for v in r])
            if len(rows) > MAX_ROWS:
                break
        return rows
    finally:
        wb.close()


def dump_result(path, sheet, hdr_idx, cols_1b):
    """按已定位的表头 dump 实测关键列 + 前 2 行全列样本。"""
    cols = {k: (v - 1 if v is not None else None) for k, v in cols_1b.items()}
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb[sheet]
        hdr_row = None
        rows, samples = [], []
        for i, row in enumerate(ws.iter_rows(values_only=True), 1):
            if i == hdr_idx:
                hdr_row = row
                continue
            if i < hdr_idx:
                continue
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
        headers = [[j + 1, str(v)] for j, v in enumerate(hdr_row or []) if v not in (None, "")]
        return {"sheet": sheet, "header_row": hdr_idx,
                "key_cols_1based": cols_1b, "headers": headers,
                "row_fields": ["row_idx", "no_raw", "mode_label", "current", "temp"],
                "rows": rows, "sample_rows_full": samples}
    finally:
        wb.close()


def main():
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    ap = argparse.ArgumentParser(description="电流数据探查 v2：对数据根目录做 JSON 快照（全递归）")
    ap.add_argument("root", help="数据根目录")
    ap.add_argument("-o", "--out", help="输出 JSON 路径（默认 <root>\\probe_dump.json）")
    ap.add_argument("--result-glob", default="Result*.xlsx", help="实测文件名通配符")
    args = ap.parse_args()

    root = os.path.abspath(args.root)
    if not os.path.isdir(root):
        raise SystemExit(f"[错误] 目录不存在: {root}")
    out_path = args.out or os.path.join(root, "probe_dump.json")

    dump = {"probe_version": 2,
            "probed_at": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "root": root, "tree": [], "workbooks": {}, "detected": {},
            "sim": [], "results": []}

    xlsx_files = []
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
                xlsx_files.append((rel, full,
                                   os.path.basename(dirpath) if rel_dir != "." else ""))

    for rel, full, folder in xlsx_files:
        info, err = scan_workbook(full)
        if info is None:
            dump["workbooks"][rel] = {"error": err}
            print(f"[跳过] {rel}: {err}")
            continue
        dump["workbooks"][rel] = info["sheets"]
        is_result_file = fnmatch.fnmatch(os.path.basename(full), args.result_glob)

        for s in info["sim_sheets"]:
            try:
                rows = dump_sheet_rows(full, s["sheet"])
                dump["sim"].append({"file": rel, "sheet": s["sheet"], "rows": rows,
                                    "truncated": len(rows) > MAX_ROWS})
                print(f"[仿真] {rel} / {s['sheet']} -> {len(rows)} 行")
            except Exception as e:
                dump["sim"].append({"file": rel, "sheet": s["sheet"],
                                    "error": f"{type(e).__name__}: {e}"})

        if is_result_file and info["result_sheets"]:
            for sn, meta in info["result_sheets"].items():
                try:
                    r = dump_result(full, sn, meta["header_row"], meta["key_cols_1based"])
                    r["file"] = os.path.basename(rel)
                    r["folder"] = os.path.dirname(rel)
                    r["folder_name"] = folder
                    dump["results"].append(r)
                    print(f"[实测] {rel} / {sn} -> {len(r['rows'])} 行 "
                          f"关键列={r['key_cols_1based']}")
                except Exception as e:
                    dump["results"].append({"file": os.path.basename(rel),
                                            "folder": os.path.dirname(rel),
                                            "error": f"{type(e).__name__}: {e}"})
        elif is_result_file:
            dump["results"].append({"file": os.path.basename(rel),
                                    "folder": os.path.dirname(rel),
                                    "error": "找不到含 NO./Current 表头的 sheet",
                                    "sheets": info["sheets"]})
            print(f"[实测] {rel} -> 没找到表头！tab={info['sheets']}")
        elif info["result_sheets"]:
            dump["detected"][rel] = {"result_like_sheets": info["result_sheets"]}
            print(f"[标记] {rel} 里有疑似实测表（未 dump）: {list(info['result_sheets'])}")

    if not dump["sim"]:
        print("\n[警告] 没有发现仿真长表（表头需含 ID/Mode/Current + simulation|Unit|Tier）！")
        print("       请确认仿真数据工作簿是否在这个目录里。")

    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(dump, f, ensure_ascii=False, indent=1, default=str)
    size_kb = os.path.getsize(out_path) / 1024
    print(f"\n[完成] 快照已写入 {out_path}（{size_kb:.0f} KB）")
    print(f"       仿真表 {len(dump['sim'])} 张 / 实测文件 {len(dump['results'])} 个")
    print("把这个 JSON 拷回开发机，用 mirror_from_probe.py 重建本地镜像。")


if __name__ == "__main__":
    main()
