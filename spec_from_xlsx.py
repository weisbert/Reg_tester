#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
spec_from_xlsx.py — 手填好 Spec 的汇总簿 → JSON（并进 chips.json，下次生成自动带上）

汇总页上 Limit / Spec(Min/Typ/Max) / 仿真(Min/Typ/Max) 那七列是留给人填的。
填完是有用的（判定列自动出 PASS/FAIL），可**一加芯片就要重出簿子**，
手填的全没了。这个脚本把填好的那份读回成数据，存进配置里，
以后 `summarize_chips.py <根目录>` 出的每一份簿子都自带 spec。

用法
    # 先看看读出来什么（不写任何文件）
    python spec_from_xlsx.py <填好的簿子.xlsx>

    # 并进配置（默认就找簿子旁边/上级的 chips.json；会先备份成 chips.json.bak）
    python spec_from_xlsx.py <填好的簿子.xlsx> --merge <数据根目录>\\chips.json

    # 只要一份独立的 spec JSON（给 summarize_chips.py --spec 用）
    python spec_from_xlsx.py <填好的簿子.xlsx> -o spec.json

合并语义
    **这本簿子覆盖到的表，以簿子为准**——表里被清空的格子就此消失（填错了能删掉）；
    簿子里没有的表原样留着（拿 --no-vco 出的簿子回填，不会把 VCO 的 spec 抹掉）。
    要整块换掉加 --replace。

读的是哪些格子
    每张表按表头自己认列（col1＝"测试项"那行是表头，行内找 Spec / 仿真 / Limit /
    判定 / 备注）——不写死列号。只收**结论行**（判定列有公式那些行）里
    真填了东西的行；分组带、条件行上的字不收（它们本来也不参与判定）。
"""

import argparse
import json
import os
import shutil
import sys

from spec_book import merge_into, read_specs

try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass


def find_config(book, given):
    """--merge 不给路径时，在簿子旁边和上一级找 chips.json。"""
    if given:
        return given
    d = os.path.dirname(os.path.abspath(book))
    for p in (os.path.join(d, "chips.json"),
              os.path.join(os.path.dirname(d), "chips.json")):
        if os.path.isfile(p):
            return p
    return None


def main():
    ap = argparse.ArgumentParser(
        description="填好 Spec 的汇总簿 → JSON（并进 chips.json，下次生成自动带上）")
    ap.add_argument("book", help="填好的汇总簿 .xlsx")
    ap.add_argument("-o", "--out", default=None,
                    help="写一份独立的 spec JSON（给 summarize_chips.py --spec 用）")
    ap.add_argument("--merge", nargs="?", const="", default=None,
                    metavar="chips.json",
                    help="并进配置文件。不给路径就在簿子旁边/上一级找 chips.json。"
                         "写之前先备份成 <名字>.bak")
    ap.add_argument("--replace", action="store_true",
                    help="整块换掉配置里原有的 spec（默认只覆盖这本簿子里有的那几张表，"
                         "簿子里没有的表原样留着）")
    ap.add_argument("--print", dest="show", action="store_true",
                    help="把 JSON 打到屏幕上（黄区回贴用）")
    args = ap.parse_args()

    if not os.path.isfile(args.book):
        sys.exit(f"找不到文件: {args.book}")
    try:
        import openpyxl  # noqa: F401
    except ImportError:
        sys.exit("缺少 openpyxl，请先: pip install openpyxl")

    data, titles, stats, warns = read_specs(args.book)
    print(f"读: {os.path.abspath(args.book)}")
    print()
    n_res = n_sp = n_si = 0
    for sheet, tkey, nr, ns, ni in stats:
        n_res += nr
        n_sp += ns
        n_si += ni
        print(f"  {sheet:<16} {tkey:<28} 结论行 {nr:>3}"
              f"   填了 Spec {ns:>3}   填了仿真 {ni:>3}")
    if not stats:
        print("  一张认得出的表都没有——这份簿子是 summarize_chips.py 出的吗？"
              "（判据：某一行第一格写着「测试项」）")
    print()
    print(f"合计: {n_res} 个结论行，其中 {n_sp} 行填了 Spec、{n_si} 行填了仿真")
    for w in warns:
        print(f"  ⚠ {w}")
    if not data:
        print()
        print("没读到任何 spec，不写文件。")
        return

    blob = {"spec": data, "spec_meta": {"titles": titles,
                                        "from": os.path.basename(args.book)}}
    if args.show:
        print()
        print(json.dumps(blob, ensure_ascii=False, indent=2))

    if args.out:
        with open(args.out, "w", encoding="utf-8") as f:
            json.dump(blob, f, ensure_ascii=False, indent=2)
            f.write("\n")
        print()
        print(f"已写出: {os.path.abspath(args.out)}")
        print(f"  用法: python summarize_chips.py <根目录> --spec "
              f"{os.path.abspath(args.out)}")

    if args.merge is None:
        if not args.out:
            print()
            print("（只是看看，没写文件。要存进配置加 --merge）")
        return

    cfg_path = find_config(args.book, args.merge)
    if not cfg_path:
        sys.exit("--merge 没给路径，簿子旁边和上一级也没有 chips.json——"
                 "把路径写全: --merge <数据根目录>\\chips.json")
    cfg = {}
    if os.path.isfile(cfg_path):
        with open(cfg_path, encoding="utf-8") as f:
            cfg = json.load(f)
        if not isinstance(cfg, dict):
            sys.exit(f"配置文件最外层要是一个对象: {cfg_path}")
        # ★ 手维护的文件，覆盖之前先留一份。折算倍数、电流分组都在里面，
        #   这个脚本写坏了就得重填一遍。
        shutil.copy2(cfg_path, cfg_path + ".bak")
    hit, kept = merge_into(cfg, data, titles, replace=args.replace)
    with open(cfg_path, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)
        f.write("\n")

    print()
    print(f"已并进: {os.path.abspath(cfg_path)}"
          + ("（原件备份在 .bak）" if os.path.isfile(cfg_path + ".bak") else ""))
    for sheet, tkey in hit:
        n = sum(len(v) for v in data[sheet][tkey].values())
        print(f"  ← {sheet} / {tkey}: {n} 行")
    for sheet, tkey in kept:
        print(f"  · {sheet} / {tkey}: 这本簿子里没有这张表，配置里原来的 spec "
              f"原样留着（要清掉加 --replace）")
    print()
    print("下次直接 `python summarize_chips.py <根目录>`，Spec 列自动填好。")


if __name__ == "__main__":
    main()
