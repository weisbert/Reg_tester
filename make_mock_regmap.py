#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
make_mock_regmap.py — 用抓回的 REG_SHEET 行 + 控制信号 list，
                      (1) 解析寄存器模型、把每个控制信号解析到 寄存器/地址/bit/默认值，
                      (2) 本地生成一个结构一模一样的寄存器 .xlsx（nManager 布局）供开发。

输入(都在 private/，含 IP，不入库)：
    --rows      pll_rows.json         explore_excel --rowdump 抓的 REG_SHEET 行区间
    --signals   control_signals.json  控制信号 list(取 reg_net)
    --aliases   aliases.json          reg_net -> 实际字段名 的变体映射(可选)
    --schema    REG_SHEET.schema.json   取 sheet 头几行元信息(Base Address 等)重建结构(可选)

输出：
    --out-xlsx  REG_SHEET_mock.xlsx     本地复刻的寄存器表(结构一致)
    --out-map   signal_reg_map.json   控制信号 -> 寄存器/地址/bit/默认/off值 解析结果

本脚本本身不含任何真实信号名，只读 private/ 输入，故可入库。
"""
import argparse
import json
import os
import re
import sys

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

# REG_SHEET (nManager) 列布局（0 基）
A, B, D, E, H, J, K, L, M, O = 0, 1, 3, 4, 7, 9, 10, 11, 12, 14


def parse_registers(rows, base):
    regs, cur = [], None
    for r in rows:
        r = list(r) + [None] * (15 - len(r))
        name, off = r[A], r[D]
        if name and off:                      # 寄存器头行
            try:
                addr = base + int(str(off), 16)
            except ValueError:
                addr = None
            cur = {"reg_name": name, "regtype": r[B], "offset": str(off),
                   "addr": (hex(addr) if addr is not None else None),
                   "reset": r[H], "width": r[E], "fields": []}
            regs.append(cur)
        if r[J]:                              # 字段行(头行也带首字段)
            f = {"name": r[J], "bit": r[K], "attr": r[L], "default": r[M], "comment": r[O]}
            if cur is None:
                cur = {"reg_name": "(above_range)", "offset": None, "addr": None,
                       "reset": None, "width": None, "fields": []}
                regs.append(cur)
            cur["fields"].append(f)
    return regs


def collect_reg_nets(sig_obj):
    nets = []

    def walk(o):
        if isinstance(o, dict):
            for k, v in o.items():
                if k == "reg_net" and isinstance(v, str):
                    nets.append((v, o))
                else:
                    walk(v)
        elif isinstance(o, list):
            for x in o:
                walk(x)
    walk(sig_obj)
    return nets


def infer_off(field):
    """据注释里的 高有效/低有效 猜 1-bit enable 的关断值。返回 (active_high, off_value) 或 (None,None)。"""
    c = str(field.get("comment") or "")
    bit = str(field.get("bit") or "")
    if ":" in bit:
        return (None, None)                   # 多 bit, 不猜
    if "高有效" in c or "high active" in c.lower() or "1: en" in c.lower():
        return (True, 0)
    if "低有效" in c or "low active" in c.lower():
        return (False, 1)
    return (None, None)


def main():
    ap = argparse.ArgumentParser(description="解析控制信号->寄存器 + 生成复刻 Excel")
    ap.add_argument("--rows", required=True)
    ap.add_argument("--signals", required=True)
    ap.add_argument("--aliases", default=None)
    ap.add_argument("--schema", default=None)
    ap.add_argument("--base", default="0xBASE_ADDR", help="块基址(默认 REG_SHEET 0xBASE_ADDR)")
    ap.add_argument("--out-xlsx", default=None)
    ap.add_argument("--out-map", default=None)
    args = ap.parse_args()

    base = int(args.base, 16)
    rows_doc = json.load(open(args.rows, encoding="utf-8"))
    rows = rows_doc["rows"]
    regs = parse_registers(rows, base)

    # 字段索引（精确 + 小写）
    fidx, ci = {}, {}
    for reg in regs:
        for f in reg["fields"]:
            fidx.setdefault(f["name"], (reg, f))
            ci.setdefault(f["name"].lower(), (reg, f))

    alias, logic_set = {}, set()
    if args.aliases and os.path.isfile(args.aliases):
        ad = json.load(open(args.aliases, encoding="utf-8"))
        alias = ad.get("alias", {})
        logic_set = set(ad.get("logic_derived", []))

    def resolve(net):
        if net in fidx:
            return fidx[net], "exact"
        if net in alias and alias[net] in fidx:
            return fidx[alias[net]], "alias"
        if net.lower() in ci:
            return ci[net.lower()], "case"
        if net in logic_set:
            return None, "logic-derived"
        return None, "unresolved"

    sig_obj = json.load(open(args.signals, encoding="utf-8"))
    nets = collect_reg_nets(sig_obj)

    entries, counts = [], {}
    for net, meta in nets:
        res, how = resolve(net)
        counts[how] = counts.get(how, 0) + 1
        e = {"reg_net": net, "match": how,
             "category": meta.get("category"), "drives": meta.get("drives")}
        if res:
            reg, f = res
            ah, offv = infer_off(f)
            e.update({"field_name": f["name"], "reg_name": reg["reg_name"],
                      "offset": reg["offset"], "addr": reg["addr"],
                      "bit": f["bit"], "attr": f["attr"], "default": f["default"],
                      "comment": f["comment"], "active_high": ah, "off_value": offv})
        entries.append(e)

    print(f"寄存器 {len(regs)} 个; 控制信号 {len(nets)} 个")
    print("解析: " + ", ".join(f"{k}={v}" for k, v in sorted(counts.items())))
    un = [e["reg_net"] for e in entries if e["match"] == "unresolved"]
    if un:
        print("未解析: " + ", ".join(un))
    print()
    print(f"{'控制信号(reg_net)':40} {'match':7} {'addr':11} {'bit':6} {'dflt':6} {'off'}")
    for e in entries:
        if e["match"] in ("unresolved", "logic-derived"):
            print(f"{e['reg_net']:40} {e['match']}")
        else:
            print(f"{e['reg_net']:40} {e['match']:7} {str(e.get('addr')):11} "
                  f"{str(e.get('bit')):6} {str(e.get('default')):6} {e.get('off_value')}")

    if args.out_map:
        json.dump({"base": args.base, "n_regs": len(regs), "counts": counts,
                   "registers": regs, "signals": entries},
                  open(args.out_map, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
        print(f"\n解析映射已写: {args.out_map}")

    if args.out_xlsx:
        import openpyxl
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "REG_SHEET"
        # 元信息头几行：优先用 schema 的 sample_rows(前6行)，否则最简重建
        meta_rows = None
        if args.schema and os.path.isfile(args.schema):
            sd = json.load(open(args.schema, encoding="utf-8"))
            meta_rows = sd.get("sheet", {}).get("sample_rows")
        if not meta_rows:
            meta_rows = [
                ["Module Name", "REG_SHEET"], ["CPU Data Width", 16],
                ["Base Address", args.base], ["Memory Name", "Memory"],
                ["Table/Register Information(Item has * mean must)"],
                ["*Reg Name", "Description", "C_Reserved", "*Offset Addr", "*Width",
                 "Table Length", "Attribute", "Default Value", "I_Reserved",
                 "*Field Name", "*Field Range", "*Field Attribute",
                 "Field Default Value", "N_Reserved", "Field Comments"],
            ]
        for row in meta_rows:
            ws.append(list(row))
        for row in rows:
            ws.append(list(row))
        os.makedirs(os.path.dirname(args.out_xlsx) or ".", exist_ok=True)
        wb.save(args.out_xlsx)
        print(f"复刻 Excel 已写: {args.out_xlsx}  ({ws.max_row} 行 x {ws.max_column} 列)")


if __name__ == "__main__":
    main()
