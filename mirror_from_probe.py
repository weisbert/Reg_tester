#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
mirror_from_probe.py — 用 probe_current_data.py 的 JSON 快照在本地重建数据镜像

重建内容（够 current_db.py 跑通即可，不还原 Result 里的寄存器列）：
  - 仿真工作簿（Current_data 页全量）
  - 各模式文件夹下的 Result 文件（NO./Mode/Current/Temperature 按原列号原行号写回）

用法：
  python mirror_from_probe.py probe_dump.json [-o private\\current_mirror]
  python current_db.py build --root private\\current_mirror --chip C1
"""
import argparse
import json
import os
import sys

import openpyxl


def main():
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    ap = argparse.ArgumentParser(description="从 probe_dump.json 重建本地数据镜像")
    ap.add_argument("dump", help="probe_current_data.py 生成的 JSON")
    ap.add_argument("-o", "--out", default=os.path.join("private", "current_mirror"),
                    help="镜像输出根目录（默认 private/current_mirror，已 gitignore）")
    args = ap.parse_args()

    with open(args.dump, "r", encoding="utf-8") as f:
        dump = json.load(f)
    out_root = os.path.abspath(args.out)
    os.makedirs(out_root, exist_ok=True)

    # 仿真工作簿
    for sim in dump.get("sim", []):
        if "error" in sim or not sim.get("rows"):
            print(f"[仿真] {sim.get('file')} 跳过（{sim.get('error', '无数据')}）")
            continue
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = sim.get("sheet", "Current_data")
        for row in sim["rows"]:
            ws.append(row)
        path = os.path.join(out_root, sim["file"])
        wb.save(path)
        print(f"[仿真] {sim['file']} -> {len(sim['rows'])} 行")

    # Result 文件
    for res in dump.get("results", []):
        if "error" in res:
            print(f"[实测] {res.get('folder')}/{res.get('file')} 跳过（{res['error']}）")
            continue
        folder = os.path.join(out_root, res["folder"])
        os.makedirs(folder, exist_ok=True)
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = res.get("sheet", "Sheet1")
        hdr_row = res["header_row"]
        for col_1b, name in res["headers"]:
            ws.cell(row=hdr_row, column=col_1b, value=name)
        cols = res["key_cols_1based"]
        for row_idx, no_raw, label, cur, temp in res["rows"]:
            if cols.get("no"):
                ws.cell(row=row_idx, column=cols["no"], value=no_raw)
            if cols.get("mode"):
                ws.cell(row=row_idx, column=cols["mode"], value=label)
            if cols.get("cur") and cur is not None:
                ws.cell(row=row_idx, column=cols["cur"], value=cur)
            if cols.get("temp") and temp is not None:
                ws.cell(row=row_idx, column=cols["temp"], value=temp)
        path = os.path.join(folder, res["file"])
        wb.save(path)
        print(f"[实测] {res['folder']}/{res['file']} -> {len(res['rows'])} 行")

    print(f"\n[完成] 镜像根目录: {out_root}")
    print(f"下一步: python current_db.py build --root \"{out_root}\" --chip C1")


if __name__ == "__main__":
    main()
