#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""probe_mode_chain.py — 只读探针：一个模式的关断链，原始文件 vs 汇总簿，当面对一遍。

用来回答同一个问题的两半：
  ① 原始文件里，这个模式在每个温度**到底有几条关断链**、基线取的是哪一行、
     每一步的格子有没有值 —— 走 current_db 自己的解析函数，所见即出簿所见。
  ② 汇总簿里，这一颗芯片这个模式的那几列**写的是什么** —— 直接读格子，不重算。

**不建库、不写文件、不改任何东西**；输出刻意压到几十行，方便隔机器贴回来。

用法（两个都给就两半都打；只给一个就只打那一半）：
  python probe_mode_chain.py --mode <模式名> --chip <芯片号> \
      --raw  "<那颗芯片的原始 all_mode 文件.xlsx>" \
      --book "<跨芯片汇总簿.xlsx>"

  --mode 用原始文件里段标签的写法或仿真表 Mode 的写法都行（大小写/下划线无关）。
  --raw 的路径漏了 .xlsx 会自动补。
"""
import argparse
import os
import re
import sys

import openpyxl

import current_db as C


def _p(*a):
    print(*a)


def _f(v, n=4):
    return "—" if v is None else f"{v:.{n}f}"


# ---------------------------------------------------------------- ① 原始文件

def probe_raw(path, mode_arg, config):
    want = C.canon_mode(mode_arg)
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        ws, hdr, cols = C.find_result_sheet(wb, config.get("result_sheet"))
        if ws is None:
            _p("[错误] 找不到含 NO./Current 表头的 tab")
            return
        raw = C.read_raw_rows(ws, hdr, cols)
    finally:
        wb.close()
    factor = C.UNIT_TO_UA.get(cols["unit"], 1000.0) / 1000.0
    _p(f"== 原始 {os.path.basename(path)}")
    _p(f"   页={ws.title} 表头行={hdr} 单位={cols['unit'] or '(空,按mA)'} 原始行={len(raw)}")

    segs = [s for s in C.split_allmode(raw) if C.canon_mode(s["mode"]) == want]
    if not segs:
        got = sorted({s["mode"] for s in C.split_allmode(raw)})
        _p(f"   [!] 没有 {mode_arg} 的段；文件里的段标签有: {'、'.join(got)}")
        return
    _p(f"   {mode_arg} 共 {len(segs)} 段")
    for s in segs:
        rows, temp0 = C.classify_raw(s["raw"], factor)
        temp = s["temp"] if s.get("temp") is not None else temp0
        steps, absorbed, sw = C.build_groups(rows, config)
        r0, r1 = rows[0]["row_idx"], rows[-1]["row_idx"]
        _p(f"\n-- {s['mode']} @ {temp}℃   行 {r0}~{r1}")
        for seq in sorted({r["seq"] for r in rows if r["seq"] >= 1}):
            g = [r for r in rows if r["seq"] == seq]
            init = next((r["row_idx"] for r in g if r["kind"] == "init"), None)
            lock = [r for r in g if r["kind"] == "lock"]
            off = [r for r in g if r["kind"] == "off"]
            off_v = [r for r in off if r["cur_ma"] is not None]
            base = next((r for r in reversed(lock) if r["cur_ma"] is not None), None)
            end = off_v[-1] if off_v else None
            tag = "  <= 出簿用的就是这一条" if seq == 1 else "  <= 被当成锁定复验，整条不用"
            _p(f"   seq{seq}: Init行={init}  lock行 {len(lock)} 个"
               f"  基线={_f(base['cur_ma']) if base else '—'}"
               f"(行{base['row_idx'] if base else '-'})"
               f"  末态={_f(end['cur_ma']) if end else '—'}"
               f"  OFF行 {len(off)} 个/带值 {len(off_v)} 个{tag}")
            blank = [r for r in off if r["cur_ma"] is None]
            if blank:
                _p("           空格子的 OFF 行: "
                   + "、".join(f"行{r['row_idx']}[{r['no_raw']}]" for r in blank))
            if seq >= 2 and off_v:
                _p("           这一条的链: "
                   + " ".join(f"{r['no_raw']}={_f(r['cur_ma'])}" for r in off_v))
        _p("   出簿的模块组: "
           + "  ".join(f"{st['disp']}={st['delta_ua']:.1f}" for st in steps))
        for line in C.parse_warns(sw, prefix="   "):
            _p(line)

    # 全文件一行收口：别的模式有没有同样的毛病，一眼看完（不逐段展开）
    bad2, bad0 = [], []
    for s in C.split_allmode(raw):
        rows, _t0 = C.classify_raw(s["raw"], factor)
        t = s["temp"]
        if any(r["seq"] >= 2 and r["kind"] == "off" and r["cur_ma"] is not None for r in rows):
            bad2.append(f"{s['mode']}@{t}")
        if any(r["seq"] == 1 and r["kind"] == "off" and r["cur_ma"] is None for r in rows):
            bad0.append(f"{s['mode']}@{t}")
    _p(f"\n[全文件] 第二条带值关断链: {('、'.join(bad2)) if bad2 else '无'}")
    _p(f"[全文件] seq1 里有空格子的关断步: {('、'.join(bad0)) if bad0 else '无'}")


# ---------------------------------------------------------------- ② 汇总簿

SKIP_HEAD = {"编号", "模块 (OFF 步)", "单位", "仿测对比", "片间一致性", "备注"}


def _merged_span(ws, row, col):
    for m in ws.merged_cells.ranges:
        if m.min_row <= row <= m.max_row and m.min_col <= col <= m.max_col:
            return m.min_col, m.max_col
    return col, col


def probe_book(path, mode_arg, chip_arg):
    want = C.canon_mode(mode_arg)
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = next((w for w in wb.worksheets if w.sheet_state == "visible"), wb.worksheets[0])
    _p(f"\n== 汇总簿 {os.path.basename(path)} / 页 {ws.title}"
       f"  ({ws.max_row}行 × {ws.max_column}列)")

    # 芯片竖条：第 1 行里不是固定表头的那些合并区
    groups = []
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=1, column=c).value
        if v is None or str(v).strip() in SKIP_HEAD:
            continue
        c0, c1 = _merged_span(ws, 1, c)
        if not groups or groups[-1][1] != c0:
            groups.append((str(v).strip(), c0, c1))
    if not groups:
        _p("   [!] 第 1 行里认不出芯片竖条")
        return
    temps = [str(ws.cell(row=2, column=c).value or "").strip()
             for c in range(groups[0][1], groups[0][2] + 1)]
    _p(f"   芯片竖条 {len(groups)} 个: "
       + "、".join(f"{n}(列{c0}~{c1})" for n, c0, c1 in groups))
    _p(f"   温度轴: {temps}")

    # 模式 band 行：A 列横跨「编号+模块+单位」三列的那些行（模块行的 A 列只占一格，
    # 里面是 "25,24,23" 这种编号——按"A 列有字"去找会把它们全认成 band）
    bands = []
    for r in range(3, ws.max_row + 1):
        v = ws.cell(row=r, column=1).value
        if not (isinstance(v, str) and v.strip()):
            continue
        c0, c1 = _merged_span(ws, r, 1)
        if c0 == 1 and c1 >= 3:
            bands.append((r, str(v).strip()))
    hit = [(r, n) for r, n in bands if C.canon_mode(n) == want]
    if not hit:
        _p(f"   [!] 找不到模式 {mode_arg}；A 列出现过的 band: "
           + "、".join(n for _r, n in bands[:20]))
        return
    r_band, band_name = hit[0]
    nxt = next((r for r, _n in bands if r > r_band), ws.max_row + 1)
    _p(f"   模式 {band_name} 在第 {r_band} 行，行区 {r_band + 1}~{nxt - 1}")

    # 1) 锁定后总电流：所有芯片一行打完 —— 7.131 到底长在哪一颗身上，看这一行
    for r in range(r_band + 1, nxt):
        name = str(ws.cell(row=r, column=2).value or "").strip()
        if name in ("锁定后总电流", "全关残留电流"):
            cells = []
            for n, c0, _c1 in groups:
                vs = [ws.cell(row=r, column=c0 + i).value for i in range(len(temps))]
                cells.append(n + "=" + "/".join(
                    ("—" if v is None else f"{float(v):g}") for v in vs))
            _p(f"   {name}(mA) 各片[{'/'.join(temps)}]: " + "  ".join(cells))

    # 2) 指定芯片那一竖条的全部行
    g = next((x for x in groups if x[0] == chip_arg), None)
    if g is None:
        _p(f"   [!] 簿子里没有芯片 {chip_arg}")
        return
    _, c0, c1 = g
    _p(f"\n   —— {chip_arg} 这一竖条 ——")
    _p(f"   {'编号':<10}{'模块':<28}{'单位':<5}" + "".join(f"{t:>11}" for t in temps))
    for r in range(r_band + 1, nxt):
        no = str(ws.cell(row=r, column=1).value or "").strip()
        name = str(ws.cell(row=r, column=2).value or "").strip()
        unit = str(ws.cell(row=r, column=3).value or "").strip()
        vals = [ws.cell(row=r, column=c).value for c in range(c0, c1 + 1)]
        if not name and not any(v is not None for v in vals):
            continue
        nd = 3 if unit.lower() == "ma" else 1     # mA 行给 3 位：5.761 不能印成 5.8
        line = f"   {no:<10}{name:<28}{unit:<5}"
        for v in vals:
            line += f"{'—':>11}" if v is None else f"{float(v):>11,.{nd}f}"
        _p(line)


def main():
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    ap = argparse.ArgumentParser(description="只读探针：一个模式的关断链，原始 vs 汇总簿")
    ap.add_argument("--mode", required=True, help="模式名（段标签或仿真 Mode 写法都行）")
    ap.add_argument("--chip", default=None, help="汇总簿里要看的那一颗芯片")
    ap.add_argument("--raw", default=None, help="原始 Result/all_mode 文件")
    ap.add_argument("--book", default=None, help="汇总簿 xlsx")
    ap.add_argument("--config", default=None, help="current_config.json（默认取 --raw 同级往上找）")
    args = ap.parse_args()

    config = dict(C.DEFAULT_CONFIG)
    cfg_path = args.config
    if cfg_path is None and args.raw:
        d = os.path.dirname(os.path.abspath(args.raw))
        for _ in range(3):
            cand = os.path.join(d, "current_config.json")
            if os.path.exists(cand):
                cfg_path = cand
                break
            d = os.path.dirname(d)
    if cfg_path and os.path.exists(cfg_path):
        import json
        with open(cfg_path, "r", encoding="utf-8-sig") as f:
            config.update(json.load(f))
        print(f"[配置] {cfg_path}")
    else:
        print("[配置] 没找到 current_config.json，用内置默认（ldo_reparent 等可能与出簿不一致）")

    if args.raw:
        p = args.raw
        if not os.path.exists(p) and os.path.exists(p + ".xlsx"):
            p += ".xlsx"
        if not os.path.exists(p):
            raise SystemExit(f"[错误] 找不到 {p}")
        probe_raw(p, args.mode, config)
    if args.book:
        if not os.path.exists(args.book):
            raise SystemExit(f"[错误] 找不到 {args.book}")
        probe_book(args.book, args.mode, args.chip or "")
    if not args.raw and not args.book:
        raise SystemExit("[错误] --raw / --book 至少给一个")


if __name__ == "__main__":
    main()
