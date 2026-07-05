#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""build_flowgraph.py —— 网表侧 adapter：conn.json -> 规范 flowgraph.json

把 extract_ports.py --connections 抽出的 sub-top 连接（conn.json）转成中间层
`flowgraph.json`（带版本号），供 GUI(信号流图 + inspector) 与序列生成器(RMW) 消费。

设计要点:
  * 节点分层：module 分组框 -> composite 黑盒(buffer bank) -> **推断子节点**(channel synth)。
    不透明 divider/route primitive（内部无数据，不合成）；logic = 控制域(默认折叠)。
  * 控制脚挂寄存器：控制脚的**驱动网**(level-shift ls_ 网)经 logic 追回 sub-top 原始端口 -> 经 drives
    反查信号。**只信连接不信名字**：寄存器位由驱动网决定，不由引脚名决定。
    例：sig_bufA_en <= ls_sigB（level-shift 后的原始网）。
  * off_controls：类别属“电流门”(master/buf/div/clk/mux/ckdiv/adc/bias_en...)的 enable 脚，
    active_high 缺失时按“高有效/关=0”兜底并标 polarity_inferred，供序列生成器逐级关。
  * 差分合并：仅当 p/n 两相**同一驱动节点**才合成一条边（真实反相器、两个 buf 各出必须各自保留）。
  * signals 表内嵌 flowgraph（引用式 signal_ref + banks）：inspector 单文件一跳到位。
  * 规则全配置化：换项目=改配置(型名正则/前缀/后缀/跨模块边)不改代码。
    **项目专属真实值放 gitignore 的 private/tool_config/build_flowgraph.json（启动自动加载）**，
    代码里的 DEFAULT_RULES 只留通用占位——保证本脚本零真实模块名/网名/信号名。

只读 private/ 输入、写 private|projects 输出；脚本本身不含真实信号名/地址。stdlib only。

用法:
  python build_flowgraph.py                         # 默认 in/out（本地 config 若存在则自动加载）
  python build_flowgraph.py --conn conn.json --regmap regmap.json --out flowgraph.json
  python build_flowgraph.py --config rules.json     # 额外规则覆盖
  python build_flowgraph.py --print                 # 打印节点/边核对清单
"""
import argparse
import io
import json
import os
import re
import sys
from collections import defaultdict, Counter

SCHEMA_VERSION = "flowgraph/1"

DEFAULT_RULES = {
    "device_blacklist": ["cdm_ggmos_wo_tiel", "CDM_CROSS_MOS_CORE_DNW", "cfmom_3t", "pinductor"],
    "power_net_regex": r"^(AVDD|AVSS|VDD|VSS|VPP|VBB|SUB)(_.*)?$",
    "internal_net_regex": r"^net\d+$",
    # 实例分类（first-match-wins）。role 决定处理方式：
    # bufbank(合成子节点) / dco / logic(控制域) / route(直通) / opaquediv(不透明,带控制脚) / prim
    # 通用占位分类规则；项目专属的 composite/route/opaquediv/dco 型名由本地 config 覆盖（见文件头说明）。
    "type_rules": [
        [r"(?i)bufbank|buf_group",       "composite", "blackbox", "bufbank"],
        [r"(?i)_logic_|_ctrl",           "primitive", "logic",    "logic"],
        [r"(?i)trace|route",             "primitive", "route",    "route"],
        [r"(?i)div_chain|divchain",      "primitive", "div",      "opaquediv"],
        [r"(?i)osc|_dco",                "primitive", "dco",      "dco"],
        [r"(?i)mux",                     "primitive", "mux",      "prim"],
        [r"(?i)^inv",                    "primitive", "inv",      "prim"],
        [r"(?i)buf|_buf",                "primitive", "buf",      "prim"],
        [r"(?i)div",                     "primitive", "div",      "prim"],
    ],
    "default_prim_device": "unknown",
    "symbols": {"dco": "oscillator", "buf": "triangle", "inv": "triangle_bubble",
                "mux": "trapezoid", "div": "box_divN", "route": "pass", "logic": "ctrl_block",
                "blackbox": "group", "unknown": "box", "group": "group"},
    "module_bands": {},                 # tag -> band 标签（项目专属，本地 config 覆盖）
    "module_reg_group_default": {},     # tag -> 默认 reg_group（项目专属，本地 config 覆盖）

    # channel synthesis
    "strip_prefix": [],   # 实例内命名域前缀（项目专属，本地 config 覆盖；不含 ls_，ls_ 只在网上）
    "ctrl_suffixes": [["_N_ictrl", "ictrl_n"], ["_P_ictrl", "ictrl_p"], ["_ictrl", "ictrl"],
                      ["_enout", "enout"], ["_enb", "enable_b"], ["_en", "enable"],
                      ["_sel", "sel"], ["_rstn", "reset"]],
    "out_suffixes": ["_outp", "_outn", "_outp_in", "_outn_in", "_out", "_lc"],
    "in_suffixes": ["_inn_in", "_inp_in", "_inn", "_inp"],
    # 标准单元/黑盒无引脚方向：按引脚名判驱动脚（输出）。ADC 的 out/ZN/clk_ref_buf 属此。
    "output_pin_regex": r"^(out|ZN|ZO|Z|Q|QN|clk_ref_buf|clk_ref|clkout)$",
    # route 节点：*_loout 是送出(输出)，*_out/*_outp 等是收进(输入)——名字带 out 但方向是入。
    "route_output_regex": r"loout",
    "lane_suffixes": ["_core", "_IB", "_QB", "_I", "_Q"],
    "diff_pairs": [["_outp", "_outn"], ["_p_lp", "_n_lp"], ["_I_lc", "_IB_lc"], ["_Q_lc", "_QB_lc"]],
    "diff_bare_pn": [["tankp", "tankn"], ["gmp", "gmn"], ["bufp", "bufn"]],
    "device_of": {"endswith": {"buf": "buf", "mux": "mux"},
                  "contains": {"div2": "div", "div": "div"}},
    "off_gate_categories": ["master_en", "buf_en", "div_en", "clk_en", "mux_en",
                            "ckdiv_en", "adc_en", "mixed_en", "bias_en"],
    # 网解析出信号但引脚名无 ctrl 后缀时（如 ADC EN、DCO_EN），按信号类别定角色
    "category_role": {
        "master_en": "enable", "buf_en": "enable", "div_en": "enable", "clk_en": "enable",
        "mux_en": "enable", "ckdiv_en": "enable", "adc_en": "enable", "mixed_en": "enable",
        "bias_en": "enable", "mux_sel": "sel", "ckdiv_cfg": "sel", "mode": "sel", "reset": "reset",
        "current_tune": "tune", "current_trim": "tune", "bias_sel": "tune", "tail_current": "tune",
        "buf_ictrl": "ictrl", "reserve": "tune",
    },
    "ctrl_roles": ["enable", "enable_b", "enout", "ictrl_n", "ictrl_p", "ictrl", "sel", "reset"],
    "levelshift_prefix": "ls_",
    # 确认的跨模块边（netlist 顶层才连得上）——项目专属实锤，从本地 config 覆盖。
    # 结构：[{"from_tag","from_net","to_tag","to_net","polarity"}]。
    "known_cross_edges": [],
    # 同名网自动连跨模块边：默认关（同名多为各 sub-top 各一份，会造假边）。
    # 换项目若同名确可信可开；开时同名边标 provenance=name + warn，且跳过 blocklist。
    "crossmodule_by_netname": False,
    "crossmodule_net_blocklist": [],    # 项目专属，本地 config 覆盖
    # 已知的 unresolved / logic-derived 集合大小，用于回归自检
    "expected_unresolved_signals": 10,
    "expected_logic_derived_signals": 2,
}


# ---------- 工具 ----------
def load_json(path):
    with io.open(path, encoding="utf-8") as f:
        return json.load(f)


def parse_expr(expr):
    """连接表达式 -> 涉及的 base 网列表（去重保序）。处理 concat {a,b,c}、slice a[3:0]。"""
    if expr is None:
        return []
    s = expr.strip()
    parts = [p.strip() for p in s[1:-1].split(",")] if (s.startswith("{") and s.endswith("}")) else [s]
    out = []
    for p in parts:
        m = re.match(r"^([A-Za-z_]\w*)(\[[\d:]+\])?$", p)
        base = m.group(1) if m else p
        if base and base not in out:
            out.append(base)
    return out


def first_net(expr):
    n = parse_expr(expr)
    return n[0] if n else None


class Rules:
    def __init__(self, d):
        self.d = d
        self.power_re = re.compile(d["power_net_regex"])
        self.internal_re = re.compile(d["internal_net_regex"])
        self.type_rules = [(re.compile(r), k, dev, role) for r, k, dev, role in d["type_rules"]]
        self.ls = d["levelshift_prefix"]
        self.out_pin_re = re.compile(d["output_pin_regex"])
        self.route_out_re = re.compile(d["route_output_regex"])

    def pin_dir(self, pin, role, route=False):
        """数据脚方向：out 后缀/输出脚名 -> output；否则 input（做 sink）。
        route 节点：只有匹配 route_output_regex(loout) 的才是输出，其余名字带 out 也是入。"""
        if route:
            return "output" if self.route_out_re.search(pin) else "input"
        if role == "out" or self.out_pin_re.match(pin):
            return "output"
        return "input"

    def primary_net(self, expr):
        """连接表达式的主网锚：优先返回首个**非电源**成员（DCO_FT_SDM 的 concat 里应取 ds 而非 AVSS）。"""
        nets = parse_expr(expr)
        for n in nets:
            if not self.is_power(n):
                return n
        return nets[0] if nets else None

    def is_power(self, net):
        # 只按电源/地网名过滤。net\d+ 类哑网不在此列：它们的非隐藏端点多 <2（另一端是被隐藏的
        # pinductor/ESD），自然不成边；而 ADC 的 net22/net017 是真实时钟网，必须放行。
        return bool(net) and bool(self.power_re.match(net))

    def classify(self, t):
        for rx, kind, dev, role in self.type_rules:
            if rx.search(t):
                return kind, dev, role
        return "primitive", self.d["default_prim_device"], "prim"

    def symbol(self, device):
        return self.d["symbols"].get(device, "box")

    def device_of(self, token):
        dv = self.d["device_of"]
        for suf, dev in dv.get("endswith", {}).items():
            if token.endswith(suf):
                return dev
        for sub, dev in dv.get("contains", {}).items():
            if sub in token:
                return dev
        return "buf"


# ---------- 端口->信号 反查 + tag 自动推导 + Logic 别名 ----------
def build_signal_index(regmap, modules):
    port_to_signal = {}
    tag_votes = defaultdict(Counter)
    mod_ports = {m["name"]: set(p[0] for p in m["ports"]) for m in modules}
    sig_by_id = {}
    for s in regmap.get("signals", []):
        sig_by_id[s["id"]] = s
        for dref in s.get("drives", []):
            if "." not in dref:
                continue
            tag, port = dref.split(".", 1)
            port_to_signal[(tag, port)] = s["id"]
            for mname, ports in mod_ports.items():
                if port in ports:
                    tag_votes[tag][mname] += 1
    tag_to_module = {tag: c.most_common(1)[0][0] for tag, c in tag_votes.items()}
    module_to_tag = {v: k for k, v in tag_to_module.items()}
    return port_to_signal, tag_to_module, module_to_tag, sig_by_id


def build_ls_alias(module, rules):
    """Logic 实例学出 ls_ 网 -> 上游原始网。成对 pin X / ls_X -> alias[ls网]=X的expr网。"""
    alias = {}
    for inst in module["instances"]:
        _k, _d, role = rules.classify(inst["t"])
        if role != "logic":
            continue
        pin_net = {pin: first_net(expr) for pin, expr in inst["c"]}
        for pin, net in pin_net.items():
            if pin.startswith(rules.ls):
                base = pin[len(rules.ls):]
                if base in pin_net and net:
                    alias[net] = pin_net[base]
    return alias


def resolve_signal(expr, tag, port_to_signal, ls_alias, ls):
    """驱动网 -> 原始端口 -> signal id。只信网不信引脚名。返回 (sid|None, src_port|None)。"""
    for net in parse_expr(expr):
        cands = [net]
        if net in ls_alias:
            cands.insert(0, ls_alias[net])
        if net.startswith(ls):
            cands.append(net[len(ls):])
        for c in cands:
            if (tag, c) in port_to_signal:
                return port_to_signal[(tag, c)], c
    return None, None


# ---------- channel 分词 ----------
def tokenize(pin, rules):
    """引脚名 -> (token, role, lane)。role ∈ ctrl角色/out/in/None。"""
    for suf, r in rules.d["ctrl_suffixes"]:
        if pin.endswith(suf):
            stem = pin[:-len(suf)]
            return _stem_token(stem, rules), r, _lane(stem, rules)
    for suf in rules.d["out_suffixes"]:
        if pin.endswith(suf):
            stem = pin[:-len(suf)]
            return _stem_token(stem, rules), "out", _lane(stem, rules)
    for suf in rules.d["in_suffixes"]:
        if pin.endswith(suf):
            stem = pin[:-len(suf)]
            return _stem_token(stem, rules), "in", _lane(stem, rules)
    return None, None, None


def _strip_prefix(s, rules):
    for p in rules.d["strip_prefix"]:
        if s.startswith(p):
            return s[len(p):]
    return s


def _lane(stem, rules):
    s = _strip_prefix(stem, rules)
    for l in rules.d["lane_suffixes"]:
        if s.endswith(l):
            return l.strip("_")
    return None


def _stem_token(stem, rules):
    s = _strip_prefix(stem, rules)
    for l in rules.d["lane_suffixes"]:
        if s.endswith(l):
            s = s[:-len(l)]
            break
    return s


def _stems(token):
    """token 的别名：去掉尾部器件词(mux/buf/div)得到语义 stem，便于输出脚按 stem 归并
    (如 <stem>_mux <-> <stem>_out)。返回一个 set（可能为空）。"""
    out = set()
    m = re.sub(r"_?(mux|buf|div)$", "", token)
    if m and m != token:
        out.add(m)
    return out


# ---------- 差分 base ----------
def diff_base(net, rules):
    """(base, polarity) 或 (net, None)。"""
    for pos, neg in rules.d["diff_pairs"]:
        if net.endswith(pos):
            return net[:-len(pos)], "p"
        if net.endswith(neg):
            return net[:-len(neg)], "n"
    for pos, neg in rules.d["diff_bare_pn"]:
        if net == pos:
            return pos[:-1], "p"
        if net == neg:
            return neg[:-1], "n"
    return net, None


# ---------- signals 表内嵌 ----------
class SignalTable:
    def __init__(self, sig_by_id, off_cats):
        self.sig_by_id = sig_by_id
        self.off_cats = set(off_cats)
        self.out = {}

    def ref(self, sid, pin_id):
        if sid is None:
            return None
        if sid not in self.out:
            s = self.sig_by_id[sid]
            self.out[sid] = {
                "reg_net": s["reg_net"], "match": s.get("match"), "category": s.get("category"),
                "resolved": s.get("resolved"), "shared": s.get("shared"), "warn": s.get("warn"),
                "desc": s.get("desc"), "active_high": s.get("active_high"), "off_value": s.get("off_value"),
                "single_copy": s.get("single_copy"), "banks": s.get("variants", {}),
                "drives": s.get("drives", []), "bound_pins": [],
            }
        if pin_id:
            self.out[sid]["bound_pins"].append(pin_id)
        return sid

    def is_off_gate(self, sid, role):
        if sid is None or role not in ("enable", "enout"):
            return False
        return self.sig_by_id[sid].get("category") in self.off_cats

    def off_control(self, sid, pin, lane):
        s = self.sig_by_id[sid]
        ah, ov, inferred = s.get("active_high"), s.get("off_value"), False
        if ah is None:
            ah, inferred = True, True
        if ov is None:
            ov, inferred = (0 if ah else 1), True
        return {"pin": pin, "signal_ref": sid, "off_value": ov, "active_high": ah,
                "polarity_inferred": inferred, "lane": lane}


# ---------- 主构建 ----------
def build(conn, regmap, rules):
    modules = conn["modules"]
    port_to_signal, tag_to_module, module_to_tag, sig_by_id = build_signal_index(regmap, modules)
    ls = rules.ls
    hide = set(rules.d["device_blacklist"])
    sigtab = SignalTable(sig_by_id, rules.d["off_gate_categories"])

    nodes = []
    diag = {"unresolved_control_pins": [], "unbound_outputs": [], "hidden_counts": {},
            "tag_map": tag_to_module, "polarity_inferred_gates": 0}
    net_ep = defaultdict(list)   # (tag, base_net) -> [(node_id, pin, dir)]

    def reg_touch(node):
        rt = []
        for c in node["controls"]:
            sid = c.get("signal_ref")
            if sid and sid in sigtab.out:
                for bank in sigtab.out[sid]["banks"].values():
                    rn = bank.get("reg_name")
                    if rn and rn not in rt:
                        rt.append(rn)
        return rt

    for m in modules:
        mname = m["name"]
        tag = module_to_tag.get(mname, mname)
        ls_alias = build_ls_alias(m, rules)
        gid = tag
        nodes.append({
            "id": gid, "kind": "module", "device": "group", "symbol": "group", "name": mname,
            "module": tag, "parent": None, "inst_type": None, "inst_name": None,
            "band": rules.d["module_bands"].get(tag), "reg_group_default": rules.d["module_reg_group_default"].get(tag, "BT"),
            "inferred": False, "hidden_default": False, "opaque_blackbox": False, "expandable": True,
            "control_domain": False, "children": [], "pins": [], "controls": [], "off_controls": [],
            "reg_touch": [], "warn": None,
        })
        hidden_c = Counter()
        for inst in m["instances"]:
            t = inst["t"]
            if t in hide:
                hidden_c[t] += 1
                continue
            kind, device, role = rules.classify(t)
            iid = "%s/%s" % (gid, inst["n"])
            node = {
                "id": iid, "kind": kind, "device": device, "symbol": rules.symbol(device),
                "name": inst["n"], "module": tag, "parent": gid, "inst_type": t, "inst_name": inst["n"],
                "inferred": False, "hidden_default": (role == "logic"),
                "opaque_blackbox": role in ("opaquediv", "route"),
                "expandable": (role == "bufbank"), "control_domain": (role == "logic"),
                "children": [], "pins": [], "controls": [], "off_controls": [],
                "reg_touch": [], "warn": None,
            }
            nodes[[n["id"] for n in nodes].index(gid)]["children"].append(iid)

            children = []
            if role == "bufbank":
                children = synth_channels(inst, iid, tag, rules, port_to_signal, ls_alias,
                                          sig_by_id, sigtab, net_ep, diag)
                node["children"] = [c["id"] for c in children]
                # bank 顶层只登记信号 IO 脚（供连边），控制脚归子节点
                annotate_io_only(node, inst, tag, rules, net_ep)
            elif role == "logic":
                pass   # 控制域：不挂控制脚、不进 net_index（纯 plumbing）
            elif role == "route":
                annotate_io_only(node, inst, tag, rules, net_ep)
            else:
                # dco / opaquediv / prim：挂控制脚 + 信号 IO
                annotate_full(node, inst, tag, rules, port_to_signal, ls_alias, sig_by_id, sigtab, net_ep, diag)
            node["reg_touch"] = reg_touch(node)
            nodes.append(node)
            for c in children:
                c["reg_touch"] = reg_touch(c)
            nodes.extend(children)
        diag["hidden_counts"][tag] = dict(hidden_c)

    # 后处理：enb 继承互补 en 的 signal_ref；off_controls 按信号去重；未覆盖门诊断
    fixup_enb_and_off(nodes, sigtab, sig_by_id)
    diag["uncovered_off_gates"] = compute_uncovered_gates(nodes, sig_by_id, rules)

    # ADC 同类型实例配对 diff_partner
    tag_diff_partner(nodes)
    # 每模块：input 端口=外部/父层驱动(连内部会造假边)；output 端口=可连到边界
    mod_input = {module_to_tag.get(m["name"], m["name"]):
                 set(p[0] for p in m["ports"] if p[1] == "input") for m in modules}
    mod_output = {module_to_tag.get(m["name"], m["name"]):
                  set(p[0] for p in m["ports"] if p[1] == "output") for m in modules}
    edges = build_edges(net_ep, rules, diag, mod_input, mod_output)
    build_cross_edges(net_ep, rules, edges)

    diag["polarity_inferred_gates"] = sum(
        1 for n in nodes for oc in n["off_controls"] if oc.get("polarity_inferred"))
    stats = {
        "modules": len(modules), "nodes": len(nodes),
        "nodes_inferred": sum(1 for n in nodes if n["inferred"]),
        "edges": len(edges), "edges_cross_module": sum(1 for e in edges if e["cross_module"]),
        "signals_referenced": len(sigtab.out),
        "controls_total": sum(len(n["controls"]) for n in nodes),
        "controls_resolved": sum(1 for n in nodes for c in n["controls"] if c.get("signal_ref")),
        "off_controls_total": sum(len(n["off_controls"]) for n in nodes),
        "unresolved_control_pins": len(diag["unresolved_control_pins"]),
    }
    return {
        "schema_version": SCHEMA_VERSION, "generated_from": "conn.json",
        "reg_base": regmap.get("base_addr"), "reg_groups": regmap.get("reg_groups", []),
        "module_tags": tag_to_module, "module_bands": rules.d["module_bands"],
        "nodes": nodes, "edges": edges, "signals": sigtab.out,
        "stats": stats, "diagnostics": diag,
    }


def _mk_control_pin(node, pin, expr, tag, rules, port_to_signal, ls_alias, sig_by_id, sigtab, diag, role, lane,
                    presid="__none__", presrc=None):
    sid, src = (presid, presrc) if presid != "__none__" else resolve_signal(expr, tag, port_to_signal, ls_alias, rules.ls)
    pin_id = "%s.%s" % (node["id"], pin)
    p = {"id": pin_id, "name": pin, "dir": "input", "net": rules.primary_net(expr), "role": role, "lane": lane}
    if role == "enable_b":
        p["complement_of"] = pin.replace("_enb", "_en") if "_enb" in pin else pin
    if sid:
        sigtab.ref(sid, pin_id)
        s = sig_by_id[sid]
        p.update({"signal_ref": sid, "shared": s.get("shared"), "warn": s.get("warn"), "resolved": s.get("resolved")})
        node["controls"].append({"pin": pin, "signal_ref": sid, "role": role, "shared": s.get("shared"), "lane": lane})
        if sigtab.is_off_gate(sid, role):
            node["off_controls"].append(sigtab.off_control(sid, pin, lane))
    else:
        p["signal_ref"] = None
        p["resolved"] = False
        if role != "enable_b":
            diag["unresolved_control_pins"].append({"node": node["id"], "pin": pin, "net": first_net(expr), "role": role})
    node["pins"].append(p)


def annotate_full(node, inst, tag, rules, port_to_signal, ls_alias, sig_by_id, sigtab, net_ep, diag):
    ctrl_roles = set(rules.d["ctrl_roles"])
    cat_role = rules.d["category_role"]
    for pin, expr in inst["c"]:
        if rules.is_power(pin):
            continue
        token, role, lane = tokenize(pin, rules)
        # 控制脚判定：引脚名后缀是 ctrl 角色，或**驱动网解析出信号**(只信连接不信名字：
        # DCO_EN / ADC EN 等大写脚名无后缀，但网能追到寄存器信号)
        name_ctrl = role in ctrl_roles
        sid, src = resolve_signal(expr, tag, port_to_signal, ls_alias, rules.ls)
        if name_ctrl or sid:
            if not name_ctrl:
                cat = sig_by_id[sid].get("category") if sid else None
                role = cat_role.get(cat, "enable")
            _mk_control_pin(node, pin, expr, tag, rules, port_to_signal, ls_alias, sig_by_id, sigtab, diag,
                            role, lane, presid=sid, presrc=src)
            continue
        base = first_net(expr)
        if not base or rules.is_power(base):
            continue
        pdir = "input" if role == "in" else rules.pin_dir(pin, role)
        node["pins"].append({"id": "%s.%s" % (node["id"], pin), "name": pin, "dir": pdir, "net": base, "role": "data_" + ("out" if pdir == "output" else "in")})
        net_ep[(tag, base)].append((node["id"], pin, pdir))


def annotate_io_only(node, inst, tag, rules, net_ep):
    route = node.get("device") == "route"
    for pin, expr in inst["c"]:
        if rules.is_power(pin) or rules.is_power(first_net(expr) or ""):
            continue
        token, role, lane = tokenize(pin, rules)
        if role in ("enable", "enable_b", "enout", "ictrl_n", "ictrl_p", "ictrl", "sel", "reset"):
            continue  # 控制脚不在此处登记
        base = rules.primary_net(expr)
        pdir = "input" if role == "in" else rules.pin_dir(pin, role, route=route)
        node["pins"].append({"id": "%s.%s" % (node["id"], pin), "name": pin, "dir": pdir, "net": base, "role": "data_" + ("out" if pdir == "output" else "in")})
        if base:
            net_ep[(tag, base)].append((node["id"], pin, pdir))


def synth_channels(inst, parent_id, tag, rules, port_to_signal, ls_alias, sig_by_id, sigtab, net_ep, diag):
    """buffer bank(composite) -> 推断子节点。lane 合并(如某 divider 的 I/Q/core -> 一个三 lane 脚)。"""
    chans = {}   # token -> {ctrl:[], out:[], in:[], tokens:set}
    for pin, expr in inst["c"]:
        if rules.is_power(pin) or rules.is_power(first_net(expr) or ""):
            continue
        token, role, lane = tokenize(pin, rules)
        if token is None:
            continue
        ch = chans.setdefault(token, {"ctrl": [], "out": [], "in": [], "tokens": set([token]) | _stems(token)})
        # 网 token（去 ls_ 后按同规则分词）也纳入 tokens，抗名字漂移（引脚名与其驱动网名可能不同）
        net = first_net(expr)
        if net:
            nb = net[len(rules.ls):] if net.startswith(rules.ls) else net
            ntok = tokenize(nb, rules)[0] or _stem_token(nb, rules)
            if ntok:
                ch["tokens"].add(ntok)
                ch["tokens"] |= _stems(ntok)
        if role in ("enable", "enable_b", "enout", "ictrl_n", "ictrl_p", "ictrl", "sel", "reset"):
            ch["ctrl"].append((pin, expr, role, lane))
        elif role == "out":
            ch["out"].append((pin, expr, lane))
        elif role == "in":
            ch["in"].append((pin, expr, lane))

    # 只有控制脚的 token 才成推断节点；纯 out/in token 暂存，稍后按 token 归并
    ctrl_tokens = {t: c for t, c in chans.items() if c["ctrl"]}
    out_only = {t: c for t, c in chans.items() if not c["ctrl"] and (c["out"] or c["in"])}

    children = []
    tokindex = {}   # 某 token -> node（含其 tokens 别名）
    for token in sorted(ctrl_tokens):
        c = ctrl_tokens[token]
        cid = "%s::%s" % (parent_id, token)
        device = rules.device_of(token)
        div_ratio = 2 if device == "div" else None
        node = {
            "id": cid, "kind": "inferred", "device": device, "symbol": rules.symbol(device),
            "div_ratio": div_ratio, "name": token, "module": tag, "parent": parent_id,
            "inst_type": None, "inst_name": None, "tokens": sorted(c["tokens"]),
            "inferred": True, "provisional": True, "hidden_default": False,
            "opaque_blackbox": False, "expandable": False, "control_domain": False,
            "children": [], "pins": [], "controls": [], "off_controls": [], "reg_touch": [], "warn": None,
        }
        for pin, expr, role, lane in c["ctrl"]:
            _mk_control_pin(node, pin, expr, tag, rules, port_to_signal, ls_alias, sig_by_id, sigtab, diag, role, lane)
        for pin, expr, lane in c["out"]:
            base = first_net(expr)
            node["pins"].append({"id": "%s.%s" % (cid, pin), "name": pin, "dir": "output", "net": base, "role": "data_out", "lane": lane})
            if base:
                net_ep[(tag, base)].append((cid, pin, "output"))
        for pin, expr, lane in c["in"]:
            base = first_net(expr)
            node["pins"].append({"id": "%s.%s" % (cid, pin), "name": pin, "dir": "input", "net": base, "role": "data_in", "lane": lane})
            if base:
                net_ep[(tag, base)].append((cid, pin, "input"))
        node["off_controls"] = _dedup_off(node["off_controls"])
        for tk in c["tokens"]:
            tokindex.setdefault(tk, node)
        children.append(node)

    # 纯输出 token：尽力按 token / stem 归并到某控制通道，否则挂 bank + 诊断
    for token in sorted(out_only):
        c = out_only[token]
        target = None
        for tk in list(c["tokens"]) + list(_stems(token)):
            if tk in tokindex:
                target = tokindex[tk]
                break
        for pin, expr, lane in c["out"] + c["in"]:
            base = first_net(expr)
            direction = "output" if (pin, expr, lane) in c["out"] else "input"
            if target:
                target["pins"].append({"id": "%s.%s" % (target["id"], pin), "name": pin, "dir": direction, "net": base, "role": "data_" + ("out" if direction == "output" else "in"), "bound": "token"})
                owner = target["id"]
            else:
                diag["unbound_outputs"].append({"composite": parent_id, "pin": pin, "net": base})
                owner = parent_id
            if base:
                net_ep[(tag, base)].append((owner, pin, direction))
    return children


def _dedup_off(offs):
    """按 (signal_ref, lane) 去重：同一寄存器位在同节点只留一个门（en/enout 指同 bit 不重复计）。"""
    seen, out = set(), []
    for o in offs:
        k = (o["signal_ref"], o.get("lane"))
        if k not in seen:
            seen.add(k)
            out.append(o)
    return out


def fixup_enb_and_off(nodes, sigtab, sig_by_id):
    """enb 互补脚继承 en 脚的 signal_ref（inspector 点 enb 也能看寄存器）；off_controls 全局去重。"""
    for n in nodes:
        by_name = {p["name"]: p for p in n["pins"]}
        for p in n["pins"]:
            if p.get("role") == "enable_b" and not p.get("signal_ref"):
                twin = by_name.get(p.get("complement_of"))
                if twin and twin.get("signal_ref"):
                    sid = twin["signal_ref"]
                    p["signal_ref"] = sid
                    p["resolved"] = twin.get("resolved")
                    p["shared"] = twin.get("shared")
                    sigtab.ref(sid, p["id"])
        n["off_controls"] = _dedup_off(n["off_controls"])


def compute_uncovered_gates(nodes, sig_by_id, rules):
    """已解析、类别属电流门、却没被任何节点 off_controls 覆盖的信号（例：某端口进了 logic 控制域
    而非 buffer bank，无可挂节点）。列出来供人工核对——避免序列生成器漏关。"""
    off_cats = set(rules.d["off_gate_categories"])
    covered = set(o["signal_ref"] for n in nodes for o in n["off_controls"])
    out = []
    for sid, s in sig_by_id.items():
        if s.get("resolved") and s.get("category") in off_cats and sid not in covered:
            bt = s.get("variants", {}).get("BT") or s.get("variants", {}).get("COMMON") or {}
            out.append({"signal": sid, "category": s.get("category"), "drives": s.get("drives", []),
                        "addr": bt.get("addr"), "bit": bt.get("bit"),
                        "note": "已解析电流门，但驱动端口不在任何建模节点上（数据缺口）——序列不会关它"})
    return out


def tag_diff_partner(nodes):
    """同 parent、同 inst_type 的两个 primitive 互标 diff_partner（ADC 的 P/N 对）。"""
    groups = defaultdict(list)
    for n in nodes:
        if n["kind"] == "primitive" and n["inst_type"]:
            groups[(n["parent"], n["inst_type"])].append(n)
    for g in groups.values():
        if len(g) == 2:
            g[0]["diff_partner"] = g[1]["id"]
            g[1]["diff_partner"] = g[0]["id"]


def build_edges(net_ep, rules, diag, mod_input, mod_output):
    """模块内连边：按 net 配对；差分仅同一驱动节点(或同一对节点)才合并。fan-out。"""
    edges = []
    # 先算每个 net 的端点
    net_points = {(tag, net): eps for (tag, net), eps in net_ep.items()}
    consumed = set()
    # 差分合并：找 p/n 同 base 且驱动同节点
    for (tag, net), eps in sorted(net_points.items()):
        if (tag, net) in consumed:
            continue
        base, pol = diff_base(net, rules)
        partner = None
        if pol:
            other = net.replace("_outp", "_outn").replace("_outn", "_outp") if "_out" in net else None
        # 找对偶网
        cand = None
        for (t2, n2) in net_points:
            if t2 != tag or n2 == net:
                continue
            b2, p2 = diff_base(n2, rules)
            if b2 == base and p2 and pol and p2 != pol:
                cand = (t2, n2)
                break
        drivers = [e for e in eps if e[2] == "output"]
        if cand:
            eps2 = net_points[cand]
            drv2 = [e for e in eps2 if e[2] == "output"]
            nodes1 = set(e[0] for e in eps)
            nodes2 = set(e[0] for e in eps2)
            # 合并成一条差分边：两相驱动同一节点，或两相都无驱动但连的是同一对节点
            same_driver = drivers and drv2 and drivers[0][0] == drv2[0][0]
            same_pair = not drivers and not drv2 and nodes1 == nodes2 and len(nodes1) == 2
            if same_driver or same_pair:
                consumed.add((tag, net))
                consumed.add(cand)
                p_net, n_net = (net, cand[1]) if pol == "p" else (cand[1], net)
                p_eps = eps if pol == "p" else eps2
                n_eps = eps2 if pol == "p" else eps
                _emit_edge(edges, tag, base, [p_net, n_net], p_eps, n_eps, True, rules, mod_input, mod_output)
                continue
        # 非差分或不合并：单网边
        consumed.add((tag, net))
        _emit_edge(edges, tag, net, [net], eps, None, False, rules, mod_input, mod_output)
    return edges


def _emit_edge(edges, tag, base, nets, p_eps, n_eps, differential, rules, mod_input, mod_output):
    if rules.is_power(base):
        return
    drivers = [e for e in p_eps if e[2] == "output"]
    direction = "forward"
    if drivers:
        src = drivers[0]
        dsts = [e for e in p_eps if e is not drivers[0]]
        # 单驱动、无内部 sink，但 net 是模块输出端口 -> 连到模块边界节点（如 ADC_CLK_OUT_P/N）
        if not dsts and any(n in mod_output.get(tag, ()) for n in nets):
            dsts = [(tag, base, "boundary")]
    else:
        # 无内部驱动脚：net 是模块 input 端口 -> 外部/父层驱动，各端点互不相连（不造假边）
        if any(n in mod_input.get(tag, ()) for n in nets):
            return
        if len(p_eps) != 2:
            return   # 多端点无驱动：方向不明，交给 GUI 手工
        src, dsts, direction = p_eps[0], [p_eps[1]], "unknown"
    dsts = [d for d in dsts if d[0] != src[0]]
    if not dsts:
        return
    e = {
        "id": "%s:%s" % (tag, base), "scope": tag, "kind": _edge_kind(base),
        "differential": differential, "direction": direction, "net_base": base, "nets": nets,
        "from": {"node": src[0], "pin": src[1]},
        "to": [{"node": d[0], "pin": d[1]} for d in dsts if d[0] != src[0]],
        "provenance": "net", "cross_module": False, "warn": None,
    }
    if differential and n_eps:
        ndrv = [x for x in n_eps if x[2] == "output"]
        nsrc = ndrv[0] if ndrv else n_eps[0]
        nsnk = [x for x in n_eps if x[2] != "output"] or n_eps[1:]
        e["pair"] = {"from_pin": nsrc[1], "to_pins": [x[1] for x in nsnk if x[0] != nsrc[0]]}
    if e["to"]:
        edges.append(e)


def _edge_kind(base):
    if "clk" in base or "82M" in base or "ckgating" in base or "sdm" in base:
        return "clk"
    if "rxsync" in base:
        return "data"
    return "lo"


def build_cross_edges(net_ep, rules, edges):
    if rules.d.get("crossmodule_by_netname"):
        blocklist = set(rules.d.get("crossmodule_net_blocklist", []))
        by_net = defaultdict(list)
        for (tag, net), eps in net_ep.items():
            for (nid, pin, d) in eps:
                by_net[net].append((tag, nid, pin, d))
        for net, eps in by_net.items():
            tags = set(e[0] for e in eps)
            if len(tags) < 2 or rules.is_power(net) or net in blocklist:
                continue
            reps = {}
            for tag, nid, pin, d in eps:
                reps.setdefault(tag, (tag, nid, pin, d))
            tl = list(reps.values())
            for i in range(len(tl)):
                for j in range(i + 1, len(tl)):
                    a, b = tl[i], tl[j]
                    edges.append({
                        "id": "X:%s" % net, "scope": "X", "kind": _edge_kind(net),
                        "differential": False, "direction": "unknown", "net_base": net, "nets": [net],
                        "from": {"node": a[1], "pin": a[2]}, "to": [{"node": b[1], "pin": b[2]}],
                        "provenance": "name", "cross_module": True,
                        "warn": "同名跨模块推断，待核（可能各 sub-top 各一份）"})
    for ce in rules.d["known_cross_edges"]:
        fa = net_ep.get((ce["from_tag"], ce["from_net"]))
        tb = net_ep.get((ce["to_tag"], ce["to_net"]))
        if fa and tb:
            edges.append({
                "id": "X:%s->%s" % (ce["from_net"], ce["to_net"]), "scope": "X", "kind": "lo",
                "differential": False, "direction": "forward",
                "net_base": "%s->%s" % (ce["from_net"], ce["to_net"]),
                "nets": [ce["from_net"], ce["to_net"]],
                "from": {"node": fa[0][0], "pin": fa[0][1]}, "to": [{"node": tb[0][0], "pin": tb[0][1]}],
                "provenance": "asserted", "cross_module": True, "warn": "netlist 顶层才连得上，配置桥接", "polarity": ce.get("polarity")})


# ---------- 打印 ----------
def print_summary(fg):
    print("schema:", fg["schema_version"], " tags:", fg["module_tags"])
    st = fg["stats"]
    print("nodes=%d(inferred=%d) edges=%d(cross=%d) signals=%d  ctrl=%d resolved=%d off=%d unresolved=%d polarity_inferred=%d" % (
        st["nodes"], st["nodes_inferred"], st["edges"], st["edges_cross_module"], st["signals_referenced"],
        st["controls_total"], st["controls_resolved"], st["off_controls_total"], st["unresolved_control_pins"],
        fg["diagnostics"]["polarity_inferred_gates"]))
    by_parent = defaultdict(list)
    for n in fg["nodes"]:
        by_parent[n["parent"]].append(n)

    def walk(pid, depth):
        for n in by_parent.get(pid, []):
            mark = "~" if n["inferred"] else ("#" if n.get("hidden_default") else " ")
            offs = ",".join(sorted(set(o["signal_ref"] for o in n["off_controls"]))) if n["off_controls"] else ""
            print("  " * depth + "%s[%s] %s <%s>%s" % (mark, n["symbol"], n["name"], n["device"],
                                                        (" off=" + offs) if offs else ""))
            walk(n["id"], depth + 1)
    for n in fg["nodes"]:
        if n["parent"] is None:
            print("=" * 3, n["name"], "band=%s" % n.get("band"), "=" * 3)
            walk(n["id"], 1)
    d = fg["diagnostics"]
    print("\ndiag: unresolved_ctrl=%d unbound_out=%d" % (len(d["unresolved_control_pins"]), len(d["unbound_outputs"])))
    for x in d["unresolved_control_pins"][:15]:
        print("   unresolved:", x["node"], x["pin"], "<=", x["net"], "(%s)" % x["role"])


def main(argv=None):
    ap = argparse.ArgumentParser(description="conn.json -> 规范 flowgraph.json")
    here = os.path.dirname(os.path.abspath(__file__))
    pdir = os.path.join(here, "private", "adpll")
    ap.add_argument("--conn", default=os.path.join(pdir, "conn.json"))
    ap.add_argument("--regmap", default=os.path.join(pdir, "regmap.json"))
    ap.add_argument("--out", default=os.path.join(pdir, "flowgraph.json"))
    ap.add_argument("--config")
    ap.add_argument("--print", dest="do_print", action="store_true")
    args = ap.parse_args(argv)

    rd = json.loads(json.dumps(DEFAULT_RULES))
    # 项目专属真实规则（模块名/网名/前缀/跨模块边）放 gitignore 的本地 config，代码里只留通用占位。
    local_cfg = os.path.join(here, "private", "tool_config", "build_flowgraph.json")
    if os.path.exists(local_cfg):
        rd.update(load_json(local_cfg))
    if args.config:
        rd.update(load_json(args.config))
    rules = Rules(rd)
    for p in (args.conn, args.regmap):
        if not os.path.exists(p):
            print("找不到输入:", p, file=sys.stderr)
            return 2
    conn = load_json(args.conn)
    regmap = load_json(args.regmap)
    fg = build(conn, regmap, rules)

    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    with io.open(args.out, "w", encoding="utf-8") as f:
        json.dump(fg, f, ensure_ascii=False, indent=1, sort_keys=False)
    print("写出 %s (%d bytes)" % (args.out, os.path.getsize(args.out)))
    st = fg["stats"]
    print("nodes=%d edges=%d signals=%d  off_controls=%d unresolved=%d" % (
        st["nodes"], st["edges"], st["signals_referenced"], st["off_controls_total"], st["unresolved_control_pins"]))
    if args.do_print:
        print()
        print_summary(fg)
    return 0


if __name__ == "__main__":
    sys.exit(main())
