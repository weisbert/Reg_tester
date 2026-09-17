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


# ---------------------------------------------------------------- ①' 两份原始文件逐模式对拍

def collect_all(path, config):
    """一份原始文件 -> {(canon模式, 温度): dict(mode, init, base, groups)}。

    ★ 要的是**锁前**那一格。簿子的条件行只有「锁定后总电流」和「全关残留电流」，
      锁前整个丢掉了——而"异常在写完初始寄存器时就已经存在"这件事，只有锁前能告诉你。
    """
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        ws, hdr, cols = C.find_result_sheet(wb, config.get("result_sheet"))
        if ws is None:
            raise SystemExit(f"[错误] {os.path.basename(path)} 找不到实测表头")
        raw = C.read_raw_rows(ws, hdr, cols)
    finally:
        wb.close()
    factor = C.UNIT_TO_UA.get(cols["unit"], 1000.0) / 1000.0
    out = {}
    for s in C.split_allmode(raw):
        rows, temp0 = C.classify_raw(s["raw"], factor)
        temp = s["temp"] if s.get("temp") is not None else temp0
        steps, _ab, _sw = C.build_groups(rows, config)
        g1 = [r for r in rows if r["seq"] == 1]
        ir = next((r for r in g1 if r["kind"] == "init"), None)
        base = next((r for r in reversed(g1)
                     if r["kind"] == "lock" and r["cur_ma"] is not None), None)
        out[(C.canon_mode(s["mode"]), temp)] = dict(
            mode=s["mode"],
            init=ir["cur_ma"] if ir else None,
            base=base["cur_ma"] if base else None,
            groups={st["disp"]: st["delta_ua"] for st in steps})
    return out


def compare_all_modes(path_a, path_b, config, group):
    """两份原始文件 **逐模式逐温度** 对拍锁前/锁定后，外加指定那一组的电流。

    一张表回答"这个块是全局 init 就开着，还是只有某些模式开着"——
    这两种结论要修的地方完全不同，而它只要读两份已经躺在盘上的文件。
    """
    A, B = collect_all(path_a, config), collect_all(path_b, config)
    _p(f"\n== 逐模式对拍   A={os.path.basename(path_a)[:34]}")
    _p(f"                B={os.path.basename(path_b)[:34]}")
    _p(f"   {'模式':<20}{'温度':>7}{'A锁前':>9}{'B锁前':>9}{'Δ锁前':>9}"
       f"{'A锁后':>9}{'B锁后':>9}{'Δ锁后':>9}   {group}: A / B")
    for k in sorted(set(A) & set(B), key=lambda k: (k[0], k[1] if k[1] is not None else 0)):
        a, b = A[k], B[k]
        d_i = (a["init"] - b["init"]) if (a["init"] is not None and b["init"] is not None) else None
        d_b = (a["base"] - b["base"]) if (a["base"] is not None and b["base"] is not None) else None
        ga, gb = a["groups"].get(group), b["groups"].get(group)
        gtxt = ("—" if ga is None else f"{ga:,.1f}") + " / " + ("—" if gb is None else f"{gb:,.1f}")
        d_g = (ga - gb) if (ga is not None and gb is not None) else None
        mark = "  ⚠" if (d_i is not None and abs(d_i) > 0.2) \
            or (d_b is not None and abs(d_b) > 0.2) \
            or (d_g is not None and abs(d_g) > 200) else ""
        _p(f"   {a['mode']:<20}{_t(k[1]):>7}{_f(a['init'],3):>9}{_f(b['init'],3):>9}"
           f"{('—' if d_i is None else f'{d_i:+.3f}'):>9}"
           f"{_f(a['base'],3):>9}{_f(b['base'],3):>9}"
           f"{('—' if d_b is None else f'{d_b:+.3f}'):>9}   {gtxt}{mark}")
    only = (set(A) ^ set(B))
    if only:
        _p("   [!] 只有一边有的 (模式,温度): "
           + "、".join(f"{m}@{t}" for m, t in sorted(only, key=str)))


def _t(v):
    return "?" if v is None else (f"{v:g}℃")


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
            ir = next((r for r in g if r["kind"] == "init"), None)
            init = ir["row_idx"] if ir else None
            lock = [r for r in g if r["kind"] == "lock"]
            off = [r for r in g if r["kind"] == "off"]
            off_v = [r for r in off if r["cur_ma"] is not None]
            base = next((r for r in reversed(lock) if r["cur_ma"] is not None), None)
            end = off_v[-1] if off_v else None
            tag = "  <= 出簿用的就是这一条" if seq == 1 else "  <= 被当成锁定复验，整条不用"
            # ★ Init（锁定前）那一格要打出来：异常在**锁定前就有**还是**锁定后才有**，
            #   是"静态偏置问题"和"跟着时钟走的问题"的分水岭，而它一直躺在原始表里没人看
            _p(f"   seq{seq}: Init行={init} 锁前={_f(ir['cur_ma']) if ir else '—'}"
               f"  lock {len(lock)} 行: "
               + "/".join(_f(r["cur_ma"]) for r in lock)
               + f"  基线={_f(base['cur_ma']) if base else '—'}"
               f"(行{base['row_idx'] if base else '-'})"
               f"  末态={_f(end['cur_ma']) if end else '—'}"
               f"  OFF行 {len(off)}/带值 {len(off_v)}{tag}")
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


def _book_layout(ws):
    """(芯片竖条, 温度轴, 模式 band 行) —— probe_book 与 scan_book 共用同一套定位。"""
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
    return groups, temps, bands


def _band_rows(ws, bands, r_band):
    """band 行的行区 [r_band+1, 下一个 band)。"""
    return range(r_band + 1, next((r for r, _n in bands if r > r_band), ws.max_row + 1))


def _open_book(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = next((w for w in wb.worksheets if w.sheet_state == "visible"), wb.worksheets[0])
    _p(f"\n== 汇总簿 {os.path.basename(path)} / 页 {ws.title}"
       f"  ({ws.max_row}行 × {ws.max_column}列)")
    return ws


def probe_book(path, mode_arg, chip_arg):
    want = C.canon_mode(mode_arg)
    ws = _open_book(path)
    lay = _book_layout(ws)
    if lay is None:
        return
    groups, temps, bands = lay
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


def scan_book(path, chip_arg):
    """把这一颗片和其余片**逐格**比一遍，只打"和别人不一样"的格子。

    ★ 簿子里本来就有「片间极差」两列，但出簿时自己说了"未设标色判据"——所以这个信息
      一直躺在那儿没人看得见。这里先用一个临时判据把它翻出来（双阈值，跟仿测偏差
      那套一个路子：光看百分比会把小电流模块放大成假红）：
          |偏离其余片中位数| > 200µA（mA 行 0.2mA）  **且**  偏离 > ±50%
      判据定了再挪进出簿脚本，别急着写死。
    """
    import statistics
    ws = _open_book(path)
    lay = _book_layout(ws)
    if lay is None:
        return
    groups, temps, bands = lay
    if not any(g[0] == chip_arg for g in groups):
        _p(f"   [!] 簿子里没有芯片 {chip_arg}")
        return
    _p(f"\n   —— {chip_arg} 逐格对其余 {len(groups) - 1} 片（只打异常）——")
    n_hit = 0
    for r_band, band_name in bands:
        for r in _band_rows(ws, bands, r_band):
            no = str(ws.cell(row=r, column=1).value or "").strip()
            name = str(ws.cell(row=r, column=2).value or "").strip()
            unit = str(ws.cell(row=r, column=3).value or "").strip()
            if not name or name.startswith("Σ"):
                continue
            # ★ 相对门限要**分行类**，不能一个 50% 走天下：
            #   模块行（µA）是单个块的电流，真异常通常是几倍，50% 挡得住噪声；
            #   条件行（mA）是**整段的总电流**，同样一笔 1.5mA 的异常摊到 5.7mA 的
            #   总电流上只有 26% —— 用 50% 去卡，等于把"总电流也高了 1.5mA"这条
            #   旁证整条漏报掉（2026-09-17 就漏了：模块行报了，总电流行一声没吭）。
            #   八片 −40 锁定电流实测散布只有 ±1.7%，10% 远在噪声之上。
            ma = unit.lower() == "ma"
            thr, rel = (0.2, 0.10) if ma else (200.0, 0.50)
            nd = 3 if ma else 1
            for ti, t in enumerate(temps):
                vals = {}
                for gname, c0, _c1 in groups:
                    v = ws.cell(row=r, column=c0 + ti).value
                    if isinstance(v, (int, float)):
                        vals[gname] = float(v)
                tv = vals.pop(chip_arg, None)
                if tv is None or len(vals) < 3:
                    continue
                med = statistics.median(vals.values())
                dev = tv - med
                if abs(dev) <= thr or med == 0 or abs(dev) / abs(med) <= rel:
                    continue
                n_hit += 1
                _p(f"   {band_name:<18}{(no or name):<12}@{t:<6}"
                   f"{chip_arg}={tv:,.{nd}f}  其余{len(vals)}片中位={med:,.{nd}f}"
                   f"  {dev:+,.{nd}f}{unit} (×{tv / med:.1f})")
    if not n_hit:
        _p("   （没有一格越过判据）")


def main():
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    ap = argparse.ArgumentParser(description="只读探针：一个模式的关断链，原始 vs 汇总簿")
    ap.add_argument("--mode", default=None, help="模式名（段标签或仿真 Mode 写法都行）；--scan 时可省")
    ap.add_argument("--scan", action="store_true",
                    help="把 --chip 那颗片在**所有模式所有行**上对其余片比一遍，只打异常格")
    ap.add_argument("--chip", default=None, help="汇总簿里要看的那一颗芯片")
    ap.add_argument("--raw", default=None, help="原始 Result/all_mode 文件")
    ap.add_argument("--raw2", default=None,
                    help="第二份原始文件：与 --raw **逐模式逐温度**对拍锁前/锁定后")
    ap.add_argument("--group", default="25,24,23",
                    help="对拍时额外列出的那一组（默认 25,24,23）")
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

    def _fix(q):
        if q and not os.path.exists(q) and os.path.exists(q + ".xlsx"):
            q += ".xlsx"
        if q and not os.path.exists(q):
            raise SystemExit(f"[错误] 找不到 {q}")
        return q

    if args.raw and args.raw2:
        compare_all_modes(_fix(args.raw), _fix(args.raw2), config, args.group)
        if not args.mode:
            return
    if args.raw:
        if not args.mode:
            raise SystemExit("[错误] --raw 要配 --mode（或配 --raw2 做逐模式对拍）")
        p = args.raw
        if not os.path.exists(p) and os.path.exists(p + ".xlsx"):
            p += ".xlsx"
        if not os.path.exists(p):
            raise SystemExit(f"[错误] 找不到 {p}")
        probe_raw(p, args.mode, config)
    if args.book:
        if not os.path.exists(args.book):
            raise SystemExit(f"[错误] 找不到 {args.book}")
        if args.scan:
            if not args.chip:
                raise SystemExit("[错误] --scan 要配 --chip")
            scan_book(args.book, args.chip)
        else:
            if not args.mode:
                raise SystemExit("[错误] --book 不带 --scan 时要配 --mode")
            probe_book(args.book, args.mode, args.chip or "")
    if not args.raw and not args.book:
        raise SystemExit("[错误] --raw / --book 至少给一个")


if __name__ == "__main__":
    main()
