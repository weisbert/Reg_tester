#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""gen_testcase.py — 无头序列生成器 + 渲染器（阶段二 M3 核心）。

读 flowgraph/1 + regmap/1 + modes/1 → 生成 testcase/1（唯一事实来源），
再渲染 ate.txt（交付格式）/ debug.html（designer 看）。

算法权威定义见仓库 SCHEMAS.md 第 5、6 节；GUI 的 webapp/generator.js 与本文件
**逐字节一致**（M4 用 node 交叉验证）。只依赖标准库；脚本本身不含任何真实信号名/地址，
全部从 JSON 读入 → 可安全进公开仓库。

用法：
    python gen_testcase.py --project projects/adpll_demo --mode BT_2G_RX
    python gen_testcase.py --flowgraph fg.json --regmap rm.json --mode-file m.json \\
        --out-json tc.json --out-ate ate.txt --out-html debug.html
    python gen_testcase.py --project projects/adpll_demo --mode BT_2G_RX --print
"""
import argparse
import html
import json
import os
import sys

SCHEMA = "testcase/1"

# 排序兜底基座：距 DCO 源头的器件类深度（边缺失时用）。见 SCHEMAS.md Step D。
DEVICE_STAGE = {
    "dco": 0, "logic": 0, "div": 1, "mux": 2, "buf": 2, "inv": 2,
    "blackbox": 2, "route": 3, "group": 0,
}


# ------------------------------------------------------------------ helpers
def load_json(path):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def parse_bit(bit):
    """'11:10' -> (11,10); '4' -> (4,4). 返回 (hi, lo)."""
    bit = str(bit).strip()
    if ":" in bit:
        hi, lo = bit.split(":", 1)
        return int(hi), int(lo)
    n = int(bit)
    return n, n


def field_mask(bit):
    hi, lo = parse_bit(bit)
    width = hi - lo + 1
    return (((1 << width) - 1) << lo), lo, width


def set_field(word, bit, val):
    mask, lo, width = field_mask(bit)
    val = int(val) & ((1 << width) - 1)
    return (word & ~mask) | (val << lo)


def get_field(word, bit):
    mask, lo, _ = field_mask(bit)
    return (word & mask) >> lo


def hex4(v):
    return "0x%04X" % (v & 0xFFFF)


def hex_addr(a):
    """规范化地址成 0x + 大写8位。"""
    s = str(a).lower().replace("0x", "")
    return "0x%08X" % int(s, 16)


def to_int(v, default=0):
    if v is None:
        return default
    if isinstance(v, int):
        return v
    s = str(v).strip()
    try:
        return int(s, 16) if s.lower().startswith("0x") else int(s, 0)
    except ValueError:
        return default


# ------------------------------------------------------------------ core model
class RegView:
    """把 regmap/1 的 signals 按 reg_group 解析成 variant 视图。"""

    def __init__(self, regmap, group):
        self.regmap = regmap
        self.group = group
        self.common = regmap.get("common_group", "COMMON")
        self.by_id = {s["id"]: s for s in regmap.get("signals", [])}

    def variant(self, sig_id):
        s = self.by_id.get(sig_id)
        if not s:
            return None
        variants = s.get("variants", {})
        return variants.get(self.group) or variants.get(self.common)

    def signal(self, sig_id):
        return self.by_id.get(sig_id)


def collect_gates(flowgraph):
    """建 gate_nodes[signal]=set(node_id)、node_gates[node]=[gate...]。gate 带 node 归属。"""
    gate_nodes = {}
    node_gates = {}
    for n in flowgraph.get("nodes", []):
        nid = n["id"]
        gs = []
        for oc in n.get("off_controls", []):
            sig = oc.get("signal_ref")
            if not sig:
                continue
            g = {
                "node": nid,
                "signal": sig,
                "pin": oc.get("pin"),
                "off_value": oc.get("off_value", 0),
                "active_high": oc.get("active_high"),
                "polarity_inferred": oc.get("polarity_inferred", False),
                "lane": oc.get("lane"),
            }
            gs.append(g)
            gate_nodes.setdefault(sig, set()).add(nid)
        if gs:
            node_gates[nid] = gs
    return gate_nodes, node_gates


def on_value(gate):
    """enable 门的'开'值：高有效→1，低有效→0；缺失按高有效兜底。"""
    ah = gate.get("active_high")
    if ah is None:
        return 1
    return 1 if ah else 0


def build_edge_maps(flowgraph, visible):
    """建 可见层 前驱表 preds[node]=set(前驱node)。跨 composite 折叠到可见祖先。"""
    node_by_id = {n["id"]: n for n in flowgraph.get("nodes", [])}

    def visible_anc(nid):
        # 折叠到最近的可见节点（inferred 子节点本身可见；隐藏/折叠情况由 visible 集合决定）
        cur = nid
        seen = set()
        while cur is not None and cur not in seen:
            seen.add(cur)
            if cur in visible:
                return cur
            n = node_by_id.get(cur)
            cur = n.get("parent") if n else None
        return nid if nid in visible else None

    preds = {}
    for e in flowgraph.get("edges", []):
        frm = visible_anc(e["from"]["node"])
        for t in e.get("to", []):
            to = visible_anc(t["node"])
            if not frm or not to or frm == to:
                continue
            preds.setdefault(to, set()).add(frm)
    return preds, node_by_id


def shutdown_rank(node_id, node_by_id, preds, memo, stack=None):
    """距 DCO 源头深度（越大=越靠末端）。边缺失时用器件类基座。带环保护。"""
    if node_id in memo:
        return memo[node_id]
    if stack is None:
        stack = set()
    if node_id in stack:  # 环：退回基座
        n = node_by_id.get(node_id, {})
        return DEVICE_STAGE.get(n.get("device"), 1)
    stack.add(node_id)
    n = node_by_id.get(node_id, {})
    base = DEVICE_STAGE.get(n.get("device"), 1)
    best = base
    for p in preds.get(node_id, ()):  # 前驱越深，本节点越深
        best = max(best, 1 + shutdown_rank(p, node_by_id, preds, memo, stack))
    stack.discard(node_id)
    memo[node_id] = best
    return best


# ------------------------------------------------------------------ generator
def generate(flowgraph, regmap, mode):
    group = mode.get("reg_group", regmap.get("primary_group", "BT"))
    rv = RegView(regmap, group)
    enabled = set(mode.get("enabled_nodes", []))
    baseline_over = mode.get("baseline", {}) or {}

    gate_nodes, node_gates = collect_gates(flowgraph)

    # Step B: 每个门信号开/关（共用位：开压倒关）
    def signal_on(sig):
        return any(n in enabled for n in gate_nodes.get(sig, ()))

    warnings = []

    # Step C: 基线寄存器映像 -------------------------------------------------
    # touched addrs = 所有门信号 ∪ baseline 信号 ∪ enabled 节点上任何 control 信号
    touched_sigs = set(gate_nodes.keys())
    touched_sigs |= set(baseline_over.keys())
    for n in flowgraph.get("nodes", []):
        if n["id"] in enabled:
            for c in n.get("controls", []):
                if c.get("signal_ref"):
                    touched_sigs.add(c["signal_ref"])

    images = {}   # addr -> {value, reg, reset}
    fields_set = {}  # addr -> list of field dicts explicitly set by this mode

    def ensure_img(sig):
        v = rv.variant(sig)
        if not v or not v.get("addr"):
            return None
        addr = hex_addr(v["addr"])
        if addr not in images:
            images[addr] = {
                "value": to_int(v.get("reset"), 0),
                "reg": v.get("reg_name"),
                "reset": to_int(v.get("reset"), 0),
            }
            fields_set[addr] = []
        return addr, v

    unresolved = []
    # 0) 预置：所有 touched 信号（含激活节点的 tune/tail 控制）建立 reset 映像，
    #    保证基线写覆盖激活通路相关的全部寄存器（如 DCO 尾电流寄存器）。
    #    不解析（无寄存器的频率/模式信号如 ct/mt/ft）静默跳过——它们不是电流门。
    for sig in sorted(touched_sigs):
        ensure_img(sig)
    # 3) 门信号写 on/off
    for sig in sorted(gate_nodes.keys()):
        # 取该 signal 任一门（同信号 off_value 一致）
        any_gate = None
        for nid in gate_nodes[sig]:
            for g in node_gates.get(nid, []):
                if g["signal"] == sig:
                    any_gate = g
                    break
            if any_gate:
                break
        r = ensure_img(sig)
        if not r:
            unresolved.append(sig)
            continue
        addr, v = r
        val = on_value(any_gate) if signal_on(sig) else any_gate["off_value"]
        images[addr]["value"] = set_field(images[addr]["value"], v["bit"], val)
        fields_set[addr].append({
            "signal": sig, "bit": v["bit"], "value": val,
            "role": "enable", "on": bool(signal_on(sig)),
        })

    # 3.5) mux_sel：把 GUI 的 MUX 选择写进对应 sel 字段（约定 0=上/1=下；真实寄存器编码可能不同，见 SCHEMAS）。
    #      放在显式 baseline 之前，故 baseline 手改值仍可压倒它。
    mux_sel = mode.get("mux_sel", {}) or {}
    node_by_id_all = {n["id"]: n for n in flowgraph.get("nodes", [])}
    for nid in sorted(mux_sel.keys()):
        node = node_by_id_all.get(nid)
        if not node:
            continue
        sel_ctrls = [c for c in node.get("controls", [])
                     if c.get("role") == "sel" or (c.get("pin") and "sel" in c["pin"].lower())]
        if len(sel_ctrls) != 1:
            if sel_ctrls:
                warnings.append("MUX %s 有多个 sel 控制，mux_sel 未落值（请在 baseline 里明确）" % nid)
            continue
        sig = sel_ctrls[0].get("signal_ref")
        r = ensure_img(sig)
        if not r:
            continue
        addr, v = r
        val = to_int(mux_sel[nid], 0)
        images[addr]["value"] = set_field(images[addr]["value"], v["bit"], val)
        fields_set[addr].append({"signal": sig, "bit": v["bit"], "value": val, "role": "mux_sel"})

    # 4) baseline 显式覆盖（tune/ictrl/ct/mt 手改值，压倒门默认与 mux_sel）
    for sig, val in baseline_over.items():
        r = ensure_img(sig)
        if not r:
            unresolved.append(sig)
            continue
        addr, v = r
        val = to_int(val, 0)
        images[addr]["value"] = set_field(images[addr]["value"], v["bit"], val)
        fields_set[addr].append({
            "signal": sig, "bit": v["bit"], "value": val, "role": "override",
        })

    if unresolved:
        warnings.append("基线里有未解析到寄存器的信号（跳过）：%s" % ", ".join(sorted(set(unresolved))))

    # 5) baseline.writes
    baseline_writes = []
    for addr in sorted(images.keys()):
        im = images[addr]
        baseline_writes.append({
            "addr": addr, "reg": im["reg"], "value": hex4(im["value"]),
            "reset": hex4(im["reset"]),
            "fields": fields_set[addr],
        })

    # Step D: 关闭步骤 -------------------------------------------------------
    # active_gate_nodes = enabled 里有至少一个基线为开门的节点
    active_nodes = []
    for nid in enabled:
        gs = [g for g in node_gates.get(nid, []) if signal_on(g["signal"])]
        if gs:
            active_nodes.append(nid)

    # 排序
    visible = set(n["id"] for n in flowgraph.get("nodes", []))  # 全节点可见（生成层不折叠）
    preds, node_by_id = build_edge_maps(flowgraph, visible)
    order_mode = (mode.get("order", {}) or {}).get("mode", "auto")
    # 次级排序键：enabled_nodes 里的位置（设计者按 源→末端 录入）→ 越靠后越靠末端 → 越先关。
    # 这解决"同 rank 同器件类相邻级"用 id 字典序会把上下游关反的问题（composite 内缺 leaf→leaf 边所致）。
    enabled_list = mode.get("enabled_nodes", []) or []
    eidx = {}
    for i, n in enumerate(enabled_list):
        eidx[n] = i
    memo = {}

    def sort_key(n):
        return (-shutdown_rank(n, node_by_id, preds, memo), -eidx.get(n, -1), n)

    if order_mode == "manual":
        manual = list((mode.get("order", {}) or {}).get("manual", []))
        ordered = [n for n in manual if n in active_nodes]
        rest = [n for n in active_nodes if n not in ordered]
        if rest:
            warnings.append("manual 顺序未覆盖的激活节点按 auto 追加：%s" % ", ".join(sorted(rest)))
        rest.sort(key=sort_key)
        ordered += rest
    else:
        ordered = sorted(active_nodes, key=sort_key)

    # 顺序不确定性提示：相邻两级 rank 相同（拓扑没定序，靠 enabled_nodes 录入序兜底）时提醒人工确认。
    ambiguous = []
    for a, b in zip(ordered, ordered[1:]):
        if shutdown_rank(a, node_by_id, preds, memo) == shutdown_rank(b, node_by_id, preds, memo):
            ambiguous.append([a, b])
    if ambiguous:
        warnings.append("以下相邻级的先后由拓扑无法判定，按 enabled_nodes 录入序（源→末端）兜底，请人工确认：%s"
                        % "; ".join("%s→%s" % (a, b) for a, b in ambiguous))

    # 逐节点累积关闭
    reg_image = {addr: images[addr]["value"] for addr in images}
    signals_off = set()
    steps = []
    shared_collateral = []
    for idx, nid in enumerate(ordered, 1):
        node = node_by_id.get(nid, {})
        gs = [g for g in node_gates.get(nid, []) if signal_on(g["signal"])]
        writes_by_addr = {}
        step_gates = []
        step_warn = []
        for g in gs:
            sig = g["signal"]
            v = rv.variant(sig)
            if not v or not v.get("addr"):
                continue
            addr = hex_addr(v["addr"])
            # 该信号还驱动哪些别的激活节点（共用位受害者）
            others = sorted(x for x in gate_nodes.get(sig, ()) if x != nid and x in enabled)
            gate_rec = {
                "signal": sig, "pin": g.get("pin"), "off_value": g["off_value"],
                "shared": len(gate_nodes.get(sig, ())) > 1,
                "polarity_inferred": g.get("polarity_inferred", False),
                "collateral_nodes": others,
            }
            step_gates.append(gate_rec)
            if sig in signals_off:
                shared_collateral.append({"step": idx, "node": nid, "signal": sig})
                continue
            before = get_field(reg_image[addr], v["bit"])
            new = set_field(reg_image[addr], v["bit"], g["off_value"])
            if new != reg_image[addr]:
                rec = writes_by_addr.setdefault(addr, {
                    "addr": addr, "reg": v.get("reg_name"),
                    "prev": hex4(reg_image[addr]), "fields": [],
                })
                rec["fields"].append({
                    "signal": sig, "bit": v["bit"],
                    "before": before, "after": g["off_value"], "role": "enable",
                })
                reg_image[addr] = new
                rec["value"] = hex4(new)
            signals_off.add(sig)
            if others:
                step_warn.append(
                    "共用位 %s 关闭同时波及：%s（这些块的电流一并消失）"
                    % (sig, ", ".join(others)))
        writes = [writes_by_addr[a] for a in sorted(writes_by_addr.keys())]
        note = None
        if not writes:
            note = "本级门已被前面共用位提前关掉（仍是一个测量点）"
        steps.append({
            "index": idx,
            "off_node": nid,
            "off_label": node.get("inst_name") or node.get("name") or nid.split("::")[-1],
            "device": node.get("device"),
            "measure": "关此级后测总电流",
            "gates": step_gates,
            "writes": writes,
            "warnings": step_warn,
            "note": note,
        })

    # Step E: 诊断透传
    uncovered = []
    for u in flowgraph.get("diagnostics", {}).get("uncovered_off_gates", []):
        uncovered.append({"signal": u.get("signal"), "note": "真门但无可挂节点，序列不会自动关，需人工补"})

    return {
        "schema_version": SCHEMA,
        "mode": mode.get("id"),
        "mode_name": mode.get("name"),
        "reg_group": group,
        "base_addr": regmap.get("base_addr"),
        "order_mode": order_mode,
        "baseline": {"note": "建立全开起始态（激活通路开、其余门关、tune/ictrl 取基线值）",
                     "writes": baseline_writes},
        "steps": steps,
        "extra_writes": mode.get("extra_writes", []) or [],
        "warnings": warnings,
        "diagnostics": {"uncovered_off_gates": uncovered, "shared_collateral": shared_collateral},
        "stats": {"baseline_regs": len(baseline_writes), "steps": len(steps),
                  "gates_off": len(signals_off)},
    }


# ------------------------------------------------------------------ renderers
def render_ate(tc):
    L = []
    ap = L.append
    ap("# " + "=" * 66)
    ap("# Test sequence: %s  (reg_group=%s)" % (tc.get("mode"), tc.get("reg_group")))
    if tc.get("mode_name"):
        ap("# %s" % tc["mode_name"])
    ap("# Generated by Reg_tester gen_testcase (%s)" % tc["schema_version"])
    ap("# 语义：累积逐级关闭；每一步先发本段写、再测总电流；相邻步电流差=该级功耗。")
    ap("# 数据行格式：ADDR VALUE MODULE  [; 行内注释]  ——ADDR=0x+大写8位、VALUE=0x+大写4位；")
    ap("#            MODULE 后可有以 ' ; ' 起的行内注释；以 # 起头的整行为纯注释。")
    ap("# " + "=" * 66)
    ap("")
    ap("# --- baseline：建立全开起始态 ---")
    for w in tc["baseline"]["writes"]:
        seg = []
        for f in w.get("fields", []):
            if f.get("role") in ("override", "mux_sel"):
                seg.append("%s=%s" % (f["signal"], f["value"]))
            elif f.get("on"):
                seg.append("%s=on" % f["signal"])
        cmt = ("  ; " + ", ".join(seg)) if seg else ""
        ap("%s %s  baseline:%s%s" % (w["addr"], w["value"], w["reg"] or "?", cmt))
    ap("")
    for st in tc["steps"]:
        ap("# --- step %d: OFF %s (%s) → 测总电流 ---"
           % (st["index"], st["off_label"], st.get("device") or "?"))
        for wn in st.get("warnings", []):
            ap("#   ⚠ %s" % wn)
        if st.get("note"):
            ap("#   · %s" % st["note"])
        for w in st["writes"]:
            ftxt = ", ".join("%s[%s]:%d→%d" % (f["signal"], f["bit"], f["before"], f["after"])
                             for f in w.get("fields", []))
            ap("%s %s  off:%s  ; %s→%s  %s"
               % (w["addr"], w["value"], st["off_node"], w["prev"], w["value"], ftxt))
        ap("")
    if tc.get("extra_writes"):
        ap("# --- extra writes（模式级额外写）---")
        for w in tc["extra_writes"]:
            ap("%s %s  extra  ; %s" % (hex_addr(w["addr"]), hex4(to_int(w["value"])), w.get("note", "")))
        ap("")
    if tc["diagnostics"]["uncovered_off_gates"]:
        ap("# --- ⚠ 未覆盖的门（需人工补写）---")
        for u in tc["diagnostics"]["uncovered_off_gates"]:
            ap("#   %s : %s" % (u["signal"], u["note"]))
    return "\n".join(L) + "\n"


def render_debug_html(tc, flowgraph=None):
    def esc(s):
        return html.escape(str(s if s is not None else ""))

    parts = []
    parts.append("""<!doctype html><html lang="zh"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>debug — %s</title><style>
:root{--bg:#0f1216;--fg:#e6e9ef;--mut:#8b93a2;--acc:#6ea8fe;--warn:#f0b400;--off:#ff6b6b;--card:#1a1f27;--line:#2a2f3a}
@media(prefers-color-scheme:light){:root{--bg:#f7f8fa;--fg:#1a1f27;--mut:#5a6270;--acc:#2b6cb0;--card:#fff;--line:#e2e6ec}}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--fg);font:14px/1.5 ui-monospace,SFMono-Regular,Menlo,Consolas,monospace}
.wrap{max-width:1100px;margin:0 auto;padding:24px}
h1{font-size:20px;margin:0 0 4px}.sub{color:var(--mut);margin-bottom:20px}
.card{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:14px 16px;margin:12px 0}
.step h2{font-size:15px;margin:0 0 8px}.badge{display:inline-block;padding:1px 8px;border-radius:20px;font-size:12px;background:var(--acc);color:#000;margin-right:6px}
table{border-collapse:collapse;width:100%%;margin-top:8px;font-size:13px}
th,td{text-align:left;padding:4px 8px;border-bottom:1px solid var(--line)}th{color:var(--mut);font-weight:600}
.addr{color:var(--acc)}.val{font-weight:700}.off{color:var(--off)}.warn{color:var(--warn)}
.bit{color:var(--mut)}.arrow{color:var(--mut)}.mut{color:var(--mut)}
.wbox{background:var(--bg);border:1px solid var(--line);border-radius:6px;padding:2px 6px;font-size:12px}
</style></head><body><div class="wrap">""" % esc(tc.get("mode")))
    parts.append("<h1>%s <span class='mut'>· %s</span></h1>" % (esc(tc.get("mode")), esc(tc.get("mode_name") or "")))
    parts.append("<div class='sub'>reg_group=<b>%s</b> · order=<b>%s</b> · base=%s · %d baseline regs · %d steps</div>"
                 % (esc(tc.get("reg_group")), esc(tc.get("order_mode")), esc(tc.get("base_addr")),
                    tc["stats"]["baseline_regs"], tc["stats"]["steps"]))
    for w in tc.get("warnings", []):
        parts.append("<div class='card warn'>⚠ %s</div>" % esc(w))

    # baseline
    parts.append("<div class='card'><h2>Baseline — 全开起始态</h2><table><tr><th>ADDR</th><th>REG</th><th>VALUE</th><th>reset</th><th>set fields</th></tr>")
    for w in tc["baseline"]["writes"]:
        ff = "; ".join("%s[%s]=%s%s" % (esc(f["signal"]), esc(f["bit"]), esc(f["value"]),
                                        " (on)" if f.get("on") else "")
                       for f in w.get("fields", []))
        parts.append("<tr><td class='addr'>%s</td><td>%s</td><td class='val'>%s</td><td class='mut'>%s</td><td>%s</td></tr>"
                     % (esc(w["addr"]), esc(w["reg"]), esc(w["value"]), esc(w.get("reset", "")), ff))
    parts.append("</table></div>")

    # steps
    for st in tc["steps"]:
        parts.append("<div class='card step'><h2><span class='badge'>step %d</span>OFF <span class='off'>%s</span> <span class='mut'>(%s)</span></h2>"
                     % (st["index"], esc(st["off_label"]), esc(st.get("device") or "")))
        for wn in st.get("warnings", []):
            parts.append("<div class='warn'>⚠ %s</div>" % esc(wn))
        if st.get("note"):
            parts.append("<div class='mut'>· %s</div>" % esc(st["note"]))
        if st["writes"]:
            parts.append("<table><tr><th>ADDR</th><th>REG</th><th>prev→VALUE</th><th>fields</th></tr>")
            for w in st["writes"]:
                ff = "; ".join("%s <span class='bit'>[%s]</span> %d<span class='arrow'>→</span><b class='off'>%d</b>"
                               % (esc(f["signal"]), esc(f["bit"]), f["before"], f["after"])
                               for f in w.get("fields", []))
                parts.append("<tr><td class='addr'>%s</td><td>%s</td><td><span class='mut'>%s</span><span class='arrow'>→</span><b class='val'>%s</b></td><td>%s</td></tr>"
                             % (esc(w["addr"]), esc(w["reg"]), esc(w["prev"]), esc(w["value"]), ff))
            parts.append("</table>")
        parts.append("<div class='mut' style='margin-top:6px'>▸ %s</div></div>" % esc(st["measure"]))

    if tc.get("extra_writes"):
        parts.append("<div class='card'><h2>Extra writes（模式级额外写）</h2><table><tr><th>ADDR</th><th>VALUE</th><th>note</th></tr>")
        for w in tc["extra_writes"]:
            parts.append("<tr><td class='addr'>%s</td><td class='val'>%s</td><td>%s</td></tr>"
                         % (esc(hex_addr(w["addr"])), esc(hex4(to_int(w["value"]))), esc(w.get("note", ""))))
        parts.append("</table></div>")

    if tc["diagnostics"]["uncovered_off_gates"]:
        parts.append("<div class='card warn'><h2>⚠ 未覆盖的门（需人工补写）</h2><ul>")
        for u in tc["diagnostics"]["uncovered_off_gates"]:
            parts.append("<li>%s — %s</li>" % (esc(u["signal"]), esc(u["note"])))
        parts.append("</ul></div>")
    parts.append("</div></body></html>")
    return "".join(parts)


# ------------------------------------------------------------------ project I/O
def load_inputs(args):
    if args.project:
        fg = load_json(os.path.join(args.project, "flowgraph.json"))
        rm = load_json(os.path.join(args.project, "regmap.json"))
        if not args.mode:
            sys.exit("--project 需配 --mode <id>")
        mode = load_json(os.path.join(args.project, "modes", args.mode + ".json"))
    else:
        if not (args.flowgraph and args.regmap and args.mode_file):
            sys.exit("需 --project 或 (--flowgraph --regmap --mode-file)")
        fg = load_json(args.flowgraph)
        rm = load_json(args.regmap)
        mode = load_json(args.mode_file)
    return fg, rm, mode


def _harden_stdout():
    """非 UTF-8 控制台（如 Windows GBK）打印含 ⚠/中文的 warning 不该崩进程。文件输出一律 utf-8。"""
    for stream in (sys.stdout, sys.stderr):
        try:
            if hasattr(stream, "reconfigure"):
                stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


def main(argv=None):
    _harden_stdout()
    ap = argparse.ArgumentParser(description="生成'电流逐级关闭'测试序列（testcase/1）+ 渲染 ate.txt/debug.html")
    ap.add_argument("--project", help="projects/<name> 目录（含 flowgraph.json/regmap.json/modes/）")
    ap.add_argument("--mode", help="模式 id（配 --project）")
    ap.add_argument("--flowgraph")
    ap.add_argument("--regmap")
    ap.add_argument("--mode-file")
    ap.add_argument("--out-json", help="写 testcase JSON（默认 project/testcases/<mode>.json）")
    ap.add_argument("--out-ate", help="写 ate.txt")
    ap.add_argument("--out-html", help="写 debug.html")
    ap.add_argument("--print", dest="do_print", action="store_true", help="打印 ate.txt 到控制台")
    args = ap.parse_args(argv)

    fg, rm, mode = load_inputs(args)
    tc = generate(fg, rm, mode)

    out_json = args.out_json
    if not out_json and args.project:
        d = os.path.join(args.project, "testcases")
        os.makedirs(d, exist_ok=True)
        out_json = os.path.join(d, (mode.get("id") or "mode") + ".json")
    if out_json:
        with open(out_json, "w", encoding="utf-8") as f:
            json.dump(tc, f, ensure_ascii=False, indent=1)

    ate = render_ate(tc)
    if args.out_ate:
        with open(args.out_ate, "w", encoding="utf-8", newline="\n") as f:
            f.write(ate)
    elif args.project and not args.do_print:
        with open(os.path.join(args.project, "testcases", (mode.get("id") or "mode") + ".ate.txt"),
                  "w", encoding="utf-8", newline="\n") as f:
            f.write(ate)

    if args.out_html:
        with open(args.out_html, "w", encoding="utf-8", newline="\n") as f:
            f.write(render_debug_html(tc, fg))
    elif args.project:
        with open(os.path.join(args.project, "testcases", (mode.get("id") or "mode") + ".debug.html"),
                  "w", encoding="utf-8", newline="\n") as f:
            f.write(render_debug_html(tc, fg))

    if args.do_print:
        sys.stdout.write(ate)
    else:
        s = tc["stats"]
        print("[gen_testcase] mode=%s group=%s -> %d baseline regs, %d steps, %d gates off"
              % (tc["mode"], tc["reg_group"], s["baseline_regs"], s["steps"], s["gates_off"]))
        if out_json:
            print("  testcase -> %s" % out_json)
        for w in tc.get("warnings", []):
            print("  ⚠ %s" % w)


if __name__ == "__main__":
    main()
