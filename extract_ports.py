#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
extract_ports.py — 从结构化 Verilog(-A) netlist 里，只抽取指定模块的端口(I/O)信息

背景：
    netlist 由 vh_extract 生成，是 flatten 的层次结构（非 ANSI 端口风格）：
        module X ( 端口名, 端口名, ... );
            input  a;
            input  [7:0] b;
            output c;
            inout  AVDD;
            wire   ...;          // 内部网络，跳过
            ...
        endmodule
    我们只关心某几个 sub-top block 的 输入/输出端口（名字/方向/位宽），
    用来定位需要哪些控制信号。

用法：
    python extract_ports.py <netlist 文件>                 # 默认抽三个目标模块
    python extract_ports.py <netlist> --modules A,B,C       # 指定模块（逗号分隔）
    python extract_ports.py <netlist> --json ports.json     # 同时导出 JSON
    python extract_ports.py <netlist> --list                # 只列出文件里所有模块名

默认目标模块：
    CHAIN_TOP_A
    CHAIN_TOP_B
    CLK_MUX
"""
import argparse
import json
import os
import re
import sys

try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

DEFAULT_MODULES = [
    "CHAIN_TOP_A",
    "CHAIN_TOP_B",
    "CLK_MUX",
]

# 电源/地端口（不是控制信号）：AVDD*/AVSS*/VDD/VSS/SUB/GND/VBAT 等
POWER_RE = re.compile(r'^(A?V(DD|SS)|D?V(DD|SS)|SUB|GND|VBAT|VCC|VREF)', re.I)

DECL_RE = re.compile(r'\b(input|output|inout)\b\s*(\[[^\]]*\])?\s*([^;]+?);')
KEYWORDS = {"input", "output", "inout", "wire", "reg", "signed", "supply0", "supply1", "tri"}


def strip_comments(t):
    t = re.sub(r'/\*.*?\*/', ' ', t, flags=re.DOTALL)
    t = re.sub(r'//[^\n]*', ' ', t)
    return t


def parse_range(r):
    """'[7:0]' -> (7, 0, 8)；None/无 -> (0,0,1)"""
    if not r:
        return (0, 0, 1)
    m = re.search(r'\[\s*(\d+)\s*:\s*(\d+)\s*\]', r)
    if not m:
        return (0, 0, 1)
    a, b = int(m.group(1)), int(m.group(2))
    return (a, b, abs(a - b) + 1)


def ident(p):
    """从一段端口文本里取标识符名（去掉范围与方向关键字，取最后一个词）。"""
    p = re.sub(r'\[[^\]]*\]', ' ', p)
    toks = [w for w in re.findall(r'\w+', p) if w not in KEYWORDS]
    return toks[-1] if toks else None


def width_str(p):
    if p["bits"] == 1:
        return "1"
    return f"[{p['msb']}:{p['lsb']}] ({p['bits']})"


def parse_module(ports_hdr, body):
    # 方向/位宽：从模块体（也兼容 ANSI 头）里的 input/output/inout 声明抽取
    dirs = {}
    for txt in (ports_hdr, body):
        for m in DECL_RE.finditer(txt):
            d, rng, names = m.group(1), m.group(2), m.group(3)
            msb, lsb, bits = parse_range(rng)
            for nm in names.split(','):
                nm = ident(nm)
                if nm:
                    dirs[nm] = {"dir": d, "msb": msb, "lsb": lsb, "bits": bits}
    # 顺序：以模块头端口列表为准
    order, seen = [], set()
    for p in ports_hdr.split(','):
        nm = ident(p)
        if nm and nm not in seen:
            order.append(nm)
            seen.add(nm)
    ports = []
    for nm in order:
        info = dirs.get(nm, {"dir": "?", "msb": 0, "lsb": 0, "bits": 1})
        ports.append({"name": nm, **info})
    # 声明了但没出现在头里的（防御）
    for nm, info in dirs.items():
        if nm not in seen:
            ports.append({"name": nm, **info})
    return ports


def find_modules(text):
    hdr_re = re.compile(r'\bmodule\s+(\w+)\s*\((.*?)\)\s*;', re.DOTALL)
    end_re = re.compile(r'\bendmodule\b')
    mods = {}
    order = []
    for m in hdr_re.finditer(text):
        name, hdr, start = m.group(1), m.group(2), m.end()
        e = end_re.search(text, start)
        body = text[start:e.start()] if e else text[start:]
        mods[name] = (hdr, body)
        order.append(name)
    return mods, order


def summarize(name, ports):
    counts = {"input": 0, "output": 0, "inout": 0, "?": 0}
    for p in ports:
        counts[p["dir"]] = counts.get(p["dir"], 0) + 1
    supply = [p for p in ports if POWER_RE.match(p["name"])]
    sig = [p for p in ports if not POWER_RE.match(p["name"])]
    return {
        "name": name,
        "n_ports": len(ports),
        "counts": counts,
        "ports": ports,
        "control_inputs": [p["name"] for p in sig if p["dir"] == "input"],
        "outputs": [p["name"] for p in sig if p["dir"] == "output"],
        "inout_signals": [p["name"] for p in sig if p["dir"] == "inout"],
        "supply": [p["name"] for p in supply],
    }


def print_module(s):
    print("=" * 90)
    c = s["counts"]
    print(f"[MODULE] {s['name']}   端口 {s['n_ports']}  "
          f"(input {c['input']}, output {c['output']}, inout {c['inout']}"
          + (f", ?{c['?']}" if c.get('?') else "") + ")")

    def block(title, names, show_width=True):
        if not names:
            return
        print(f"\n  {title} ({len(names)}):")
        by = {p["name"]: p for p in s["ports"]}
        for nm in names:
            p = by.get(nm)
            w = f"  {width_str(p)}" if (p and show_width and p['bits'] > 1) else ""
            print(f"    {nm}{w}")

    block("控制输入 input (非电源)", s["control_inputs"])
    block("输出 output", s["outputs"])
    block("模拟 inout (非电源, 如 LO/tank/buf)", s["inout_signals"])
    if s["supply"]:
        print(f"\n  电源/地 ({len(s['supply'])}): " + ", ".join(s["supply"]))


def main():
    ap = argparse.ArgumentParser(description="抽取指定模块的端口 I/O")
    ap.add_argument("path", help="netlist 文件 (.vh/.v/.va 等)")
    ap.add_argument("--modules", default=None, help="逗号分隔模块名；默认三个目标模块")
    ap.add_argument("--json", default=None, help="导出 JSON 到该文件")
    ap.add_argument("--list", action="store_true", help="只列出文件里所有模块名")
    args = ap.parse_args()

    if not os.path.isfile(args.path):
        sys.exit(f"找不到文件: {args.path}")
    with open(args.path, encoding="utf-8", errors="replace") as f:
        text = strip_comments(f.read())

    mods, order = find_modules(text)
    print(f"文件: {args.path}")
    print(f"共找到 {len(order)} 个 module\n")

    if args.list:
        for nm in order:
            print("  " + nm)
        return

    targets = ([m.strip() for m in args.modules.split(",") if m.strip()]
               if args.modules else DEFAULT_MODULES)

    results, not_found = [], []
    for nm in targets:
        if nm not in mods:
            not_found.append(nm)
            continue
        hdr, body = mods[nm]
        s = summarize(nm, parse_module(hdr, body))
        results.append(s)
        print_module(s)

    if not_found:
        print("\n" + "!" * 90)
        print("以下模块名在文件里没找到（确认拼写，或用 --list 查看全部）：")
        for nm in not_found:
            print("  " + nm)

    if args.json:
        with open(args.json, "w", encoding="utf-8") as f:
            json.dump({"file": os.path.basename(args.path),
                       "modules": results,
                       "not_found": not_found}, f, ensure_ascii=False, indent=2)
        print(f"\n已导出 JSON: {args.json}")


if __name__ == "__main__":
    main()
