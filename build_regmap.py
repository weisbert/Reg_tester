#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""build_regmap.py —— 寄存器侧 adapter：signal_reg_map.json -> 规范 regmap.json

把阶段一解析出的 signal_reg_map.json 规范化成中间层 `regmap.json`（带版本号），
供 GUI(inspector) 与序列生成器(RMW) 消费。核心增量：给每个控制信号补 **WL / WLT
平行寄存器字段**（BT/WL 两套并行寄存器组都要测，见 PLAN 阶段二决策 5）。

匹配策略（抗命名漂移）：寄存器按“基名”分组——去掉各 reg_group 前缀后同基名的寄存器互为孪生
（主组 / 并行组 / 第三份）。信号字段 → 孪生寄存器里**同 bit 位**的字段 = 平行字段。
bit 位是主键，信号名前缀变换只做校验。项目专属的真实前缀规则放 gitignore 的
private/tool_config/build_regmap.json（启动自动加载）；代码里的 DEFAULT_RULES 只留通用占位。

只读 private/ 输入、写 private/ 输出；脚本本身不含真实信号名/地址。stdlib only。

用法:
  python build_regmap.py                                   # 默认 in/out（见下）
  python build_regmap.py --in signal_reg_map.json --out regmap.json
  python build_regmap.py --config regmap_rules.json        # 覆盖默认分组规则
  python build_regmap.py --print                           # 额外打印一张紧凑核对表
"""
import argparse
import io
import json
import os
import re
import sys

SCHEMA_VERSION = "regmap/1"

# ---- 默认规则（可被 --config 覆盖；换项目改这里而非改代码）--------------------
DEFAULT_RULES = {
    # 寄存器名前缀 -> reg_group 标签。顺序=匹配优先级（长前缀 WL1 必须排在 WL 前）。
    "reg_group_prefixes": [
        ["WL1_", "WLT"],
        ["WL_", "WL"],
        ["BT_", "BT"],
    ],
    "primary_group": "BT",       # 无前缀 / 主寄存器所属组
    # 视为“单副本、对所有组通用”的 reg_group（无孪生的寄存器都归到它）
    "common_group": "COMMON",
    # 名字变换校验：primary 字段名 -> 各组期望前缀替换（仅告警，不作为匹配依据）。
    # 项目专属的真实前缀放 gitignore 的 private/tool_config/build_regmap.json（启动自动加载）；
    # 代码里默认空 = 跳过名字校验，不影响 bit 位主键匹配与输出。
    "name_xform": {},
}


def load_json(path):
    with io.open(path, encoding="utf-8") as f:
        return json.load(f)


def load_project(proj_dir):
    """读工程包 project.json（schema/2）。None -> 走传统 private/adpll + tool_config 默认。"""
    if not proj_dir:
        return None
    p = os.path.join(proj_dir, "project.json")
    if not os.path.exists(p):
        sys.exit("工程包缺 project.json: %s" % p)
    return load_json(p)


def base_name(reg_name, prefixes):
    """去掉 reg_group 前缀返回 (group_label, base)。无前缀 -> (None, reg_name)。"""
    for pref, label in prefixes:
        if reg_name.startswith(pref):
            return label, reg_name[len(pref):]
    return None, reg_name


def field_at_bit(reg, bit):
    for f in reg.get("fields", []):
        if f.get("bit") == bit:
            return f
    return None


def variant_from(reg, field, group):
    """从寄存器 + 字段构造一个 variant 记录。"""
    return {
        "reg_group": group,
        "reg_name": reg["reg_name"],
        "offset": reg.get("offset"),
        "addr": reg.get("addr"),
        "reset": reg.get("reset"),          # 整字复位值（算 baseline 用）
        "width": reg.get("width"),
        "field_name": field["name"],
        "bit": field["bit"],
        "attr": field.get("attr"),
        "default": field.get("default"),    # 该 bit 字段默认值
        "comment": field.get("comment"),
    }


def name_matches(primary_field, cand_field, group, rules):
    """名字变换校验：primary 字段名按规则变换后是否 == 候选名。返回 True/False/None(无规则)。"""
    xforms = rules.get("name_xform", {}).get(group)
    if not xforms:
        return None
    pn = primary_field
    for a, b in xforms:
        if pn.startswith(a):
            return pn.replace(a, b, 1) == cand_field
    return None


def load_annotations(path):
    """从 control_signals.json 抽 reg_net -> {warn, desc, shared}（连接侧标注，非寄存器侧）。"""
    ann = {}
    if not path or not os.path.exists(path):
        return ann
    cs = load_json(path)
    for key in ("primary_current_related", "config_secondary"):
        for e in cs.get(key, []):
            rn = e.get("reg_net")
            if not rn:
                continue
            ann[rn] = {"warn": e.get("warn"), "desc": e.get("desc"), "shared": e.get("shared", False)}
    return ann


def build(sigmap, rules, annotations=None):
    annotations = annotations or {}
    prefixes = rules["reg_group_prefixes"]
    primary = rules["primary_group"]
    common = rules["common_group"]

    regs = sigmap.get("registers", [])

    # 1) 寄存器按基名分组，标注每个寄存器的 reg_group
    groups = {}          # base -> {group_label: reg}
    reg_group_of = {}    # reg_name -> group_label
    canon_registers = []
    for r in regs:
        label, base = base_name(r["reg_name"], prefixes)
        g = label or primary
        # 基名带前缀的才算孪生候选；无前缀寄存器每个自成一组（不参与孪生）
        key = base if label else ("#" + r["reg_name"])
        groups.setdefault(key, {})[g] = r
        reg_group_of[r["reg_name"]] = g
        cr = dict(r)
        cr["reg_group"] = g
        cr["base_name"] = base
        canon_registers.append(cr)

    # 一个基名是否真“孪生”（含 >=2 个 group，或含 WL/WLT）
    def twin_groups(key):
        g = groups.get(key, {})
        return {k: v for k, v in g.items() if k != primary}

    warnings = []          # 真异常：名字变换不符
    notes = []             # 预期内信息：孪生寄存器某 bit 是 reserved（如 WLT 稀疏）
    out_signals = []
    n_wl = n_wlt = n_common = n_resolved = 0
    # 哪些组允许“稀疏”（缺 bit 属正常，不告警）——WLT(WL1) 只放少数关键位
    sparse_groups = set(rules.get("sparse_groups", ["WLT"]))

    for s in sigmap.get("signals", []):
        reg_net = s["reg_net"]
        resolved = "reg_name" in s and s.get("reg_name") is not None
        ann = annotations.get(reg_net, {})
        sig = {
            "id": reg_net,
            "reg_net": reg_net,
            "match": s.get("match"),
            "resolved": resolved,
            "category": s.get("category"),
            "drives": s.get("drives", []),
            "shared": ann.get("shared") or (len(s.get("drives", [])) > 1),
            "warn": ann.get("warn"),
            "desc": ann.get("desc"),
            "active_high": s.get("active_high"),
            "off_value": s.get("off_value"),
            "comment": s.get("comment"),
        }
        if not resolved:
            # 未解析（频率码 ct/mt/ft 等）/ logic-derived：无寄存器位，保留占位
            sig["single_copy"] = None
            sig["variants"] = {}
            out_signals.append(sig)
            continue

        n_resolved += 1
        rn = s["reg_name"]
        bit = s["bit"]
        pgroup = reg_group_of.get(rn, primary)
        label, base = base_name(rn, prefixes)
        key = base if label else ("#" + rn)
        twins = twin_groups(key)

        primary_reg = groups[key].get(pgroup) or next(iter(groups[key].values()))
        pfield = field_at_bit(primary_reg, bit) or {
            "name": s.get("field_name"), "bit": bit, "attr": s.get("attr"),
            "default": s.get("default"), "comment": s.get("comment"),
        }

        variants = {}
        if not twins:
            # 无孪生 -> 单副本，对所有 reg_group 通用
            sig["single_copy"] = True
            variants[common] = variant_from(primary_reg, pfield, common)
            n_common += 1
        else:
            sig["single_copy"] = False
            # 主组
            variants[pgroup] = variant_from(primary_reg, pfield, pgroup)
            # 平行组：同 bit 位取字段
            for g, greg in twins.items():
                cf = field_at_bit(greg, bit)
                if cf is None:
                    msg = "%s: %s 组在 %s 无 bit=%s 的字段（reserved）" % (reg_net, g, greg["reg_name"], bit)
                    (notes if g in sparse_groups else warnings).append(msg)
                    continue
                # 名字变换校验（仅告警）
                ok = name_matches(pfield.get("name", ""), cf["name"], g, rules)
                if ok is False:
                    warnings.append("%s: %s 组名字变换不符 (%s -> 期望≠%s)" %
                                    (reg_net, g, pfield.get("name"), cf["name"]))
                variants[g] = variant_from(greg, cf, g)
                if g == "WL":
                    n_wl += 1
                elif g == "WLT":
                    n_wlt += 1

        sig["variants"] = variants
        out_signals.append(sig)

    # reg_groups 自然顺序：主组在前，其余按 BT,WL,WLT 习惯序
    order = {"BT": 0, "WL": 1, "WLT": 2}
    other = sorted((lbl for _, lbl in prefixes if lbl != primary), key=lambda x: order.get(x, 99))
    reg_groups = [primary] + other
    regmap = {
        "schema_version": SCHEMA_VERSION,
        "generated_from": "signal_reg_map.json",
        "base_addr": sigmap.get("base"),
        "reg_groups": reg_groups,          # ["BT","WLT","WL"] 顺序化
        "primary_group": primary,
        "common_group": common,
        "note": ("单副本信号 variants 只有 COMMON 键（对所有组通用）；孪生信号有 "
                 "BT/WL[/WLT] 键。生成器取 variants.get(mode.reg_group) or variants[COMMON]。"),
        "registers": canon_registers,
        "signals": out_signals,
        "stats": {
            "signals_total": len(out_signals),
            "signals_resolved": n_resolved,
            "with_wl_variant": n_wl,
            "with_wlt_variant": n_wlt,
            "single_copy": n_common,
            "warnings": warnings,
            "notes": notes,
        },
    }
    return regmap


def print_table(regmap):
    print("reg_net".ljust(38), "grp", "BT/WL/WLT addr@bit")
    for s in regmap["signals"]:
        if not s["resolved"]:
            print(s["reg_net"].ljust(38), "-  ", "(unresolved:%s)" % s["match"])
            continue
        v = s["variants"]
        if s["single_copy"]:
            c = v["COMMON"]
            desc = "COMMON %s@%s" % (c["addr"], c["bit"])
        else:
            parts = []
            for g in ("BT", "WL", "WLT"):
                if g in v:
                    parts.append("%s %s@%s" % (g, v[g]["addr"], v[g]["bit"]))
            desc = " | ".join(parts)
        shared = "S" if len(s.get("drives", [])) > 1 else " "
        print(s["reg_net"].ljust(38), shared, " ", desc)


def main(argv=None):
    ap = argparse.ArgumentParser(description="signal_reg_map.json -> 规范 regmap.json（补 WL/WLT 平行字段）")
    here = os.path.dirname(os.path.abspath(__file__))
    pdir = os.path.join(here, "private", "adpll")
    ap.add_argument("--project", help="工程包目录（project.json schema/2）：规则读 regbook 段，"
                                       "IO 按 artifacts 从包内解析。取代 private/tool_config + private/adpll 默认。")
    ap.add_argument("--in", dest="inp", default=None, help="输入 signal_reg_map.json")
    ap.add_argument("--out", dest="out", default=None, help="输出 regmap.json")
    ap.add_argument("--signals", default=None, help="control_signals.json（并 warn/desc 标注）")
    ap.add_argument("--config", help="规则覆盖 JSON（合并进 DEFAULT_RULES）")
    ap.add_argument("--print", dest="do_print", action="store_true", help="额外打印核对表")
    args = ap.parse_args(argv)

    proj = load_project(args.project)
    art = proj.get("artifacts", {}) if proj else {}

    def ppath(name, default):
        return os.path.join(args.project, name) if args.project else os.path.join(pdir, default)

    inp = args.inp or ppath(art.get("signal_reg_map", "signal_reg_map.json"), "signal_reg_map.json")
    out = args.out or ppath(art.get("regmap", "regmap.json"), "regmap.json")
    signals_path = args.signals or ppath(art.get("control_signals", "control_signals.json"), "control_signals.json")

    rules = json.loads(json.dumps(DEFAULT_RULES))
    if proj is not None:
        rb = proj.get("regbook", {})
        for k in ("reg_group_prefixes", "primary_group", "common_group", "name_xform"):
            if k in rb:
                rules[k] = rb[k]
    else:
        local_cfg = os.path.join(here, "private", "tool_config", "build_regmap.json")
        if os.path.exists(local_cfg):
            rules.update(load_json(local_cfg))
    if args.config:
        rules.update(load_json(args.config))

    if not os.path.exists(inp):
        print("找不到输入:", inp, file=sys.stderr)
        return 2
    sigmap = load_json(inp)
    annotations = load_annotations(signals_path)
    regmap = build(sigmap, rules, annotations)

    os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
    with io.open(out, "w", encoding="utf-8") as f:
        json.dump(regmap, f, ensure_ascii=False, indent=1)
    st = regmap["stats"]
    print("写出 %s (%d bytes)" % (out, os.path.getsize(out)))
    print("signals: total=%d resolved=%d  variants: WL=%d WLT=%d single-copy=%d" %
          (st["signals_total"], st["signals_resolved"], st["with_wl_variant"],
           st["with_wlt_variant"], st["single_copy"]))
    if st["warnings"]:
        print("warnings(%d):" % len(st["warnings"]))
        for w in st["warnings"]:
            print("  -", w)
    else:
        print("warnings: 0 (无名字变换异常)")
    if st.get("notes"):
        print("notes(%d): WLT 稀疏寄存器缺位属正常，示例: %s" % (len(st["notes"]), st["notes"][0]))
    if args.do_print:
        print()
        print_table(regmap)
    return 0


if __name__ == "__main__":
    sys.exit(main())
