#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
extract_ports.py — 从结构化 Verilog(-A) netlist 抽取指定模块的端口 与 内部连接

背景：
    netlist 由 vh_extract 生成，是 flatten 的层次结构（非 ANSI 端口风格）：
        module X ( 端口名, ... );
            input  a;  input [7:0] b;  output c;  inout AVDD;   // 端口方向声明
            wire   n1;  wire [1:0] n2;                          // 内部网络
            assign a = n1;                                      // 别名(可选)
            ChildType inst0 ( .EN(n1), .OUTP(c), .IN(b[0]) );   // 子模块例化
            ...
        endmodule

    两种用法：
      1) 端口模式(默认)：只抽端口 I/O（名/方向/位宽）。
      2) 连接模式(--connections)：把端口 + wire + 例化(instance)+连线 全抽出来，
         并建 net_index（每根网络接到哪些 顶层端口/实例引脚），用于追踪
         EN → buffer实例 → 输出 的真实通路。命名多轮迭代后不可信，只信连接。

用法：
    python extract_ports.py <netlist>                         # 端口模式，默认三个目标模块
    python extract_ports.py <netlist> --json ports.json
    python extract_ports.py <netlist> --list                 # 列出所有模块名
    python extract_ports.py <netlist> --modules A,B,C

    python extract_ports.py <netlist> --connections --json conn.json   # 连接模式(核心)
    python extract_ports.py <netlist> --body <MODULE> --head 80        # 打印模块体原文(核对语法)

默认目标模块：项目专属，放 gitignore 的 private/tool_config/extract_ports.json（启动自动加载），
    或用 --modules A,B,C 显式指定。一般是若干 DCO/LO 链顶层 block + 一个时钟 MUX block。

自诊断：连接模式下，模块体里凡是没被识别为 端口/wire/assign/instance 的残留
    都会进 unparsed 列表。unparsed 为空、且 instance/connection 数量合理，才算可信。
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

# 默认目标模块名属项目专属 IP：不硬编码。优先读 gitignore 的本地 config，否则空（需 --modules）。
DEFAULT_MODULES = []


def _load_default_modules():
    p = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                     "private", "tool_config", "extract_ports.json")
    if os.path.exists(p):
        try:
            with open(p, encoding="utf-8") as f:
                return json.load(f).get("modules", [])
        except Exception:
            pass
    return list(DEFAULT_MODULES)

POWER_RE = re.compile(r'^(A?V(DD|SS)|D?V(DD|SS)|SUB|GND|VBAT|VCC|VREF)', re.I)

DECL_RE = re.compile(r'\b(input|output|inout)\b\s*(\[[^\]]*\])?\s*([^;]+?);')
WIRE_RE = re.compile(r'\b(wire|tri|wand|wor|supply0|supply1|tri0|tri1)\b\s*(\[[^\]]*\])?\s*([^;]+?);')
ASSIGN_RE = re.compile(r'\bassign\b([^;]*);')
KEYWORDS = {"input", "output", "inout", "wire", "reg", "signed", "supply0", "supply1",
            "tri", "wand", "wor", "assign", "module", "endmodule"}


# ----------------------------- 通用小工具 -----------------------------

def strip_comments(t):
    t = re.sub(r'/\*.*?\*/', ' ', t, flags=re.DOTALL)
    t = re.sub(r'//[^\n]*', ' ', t)
    return t


def parse_range(r):
    """'[7:0]' -> (7,0,8)；None/无 -> (0,0,1)"""
    if not r:
        return (0, 0, 1)
    m = re.search(r'\[\s*(\d+)\s*:\s*(\d+)\s*\]', r)
    if not m:
        return (0, 0, 1)
    a, b = int(m.group(1)), int(m.group(2))
    return (a, b, abs(a - b) + 1)


def ident(p):
    """从端口文本取标识符名（去范围与关键字，取最后一个词）。"""
    p = re.sub(r'\[[^\]]*\]', ' ', p)
    toks = [w for w in re.findall(r'\\\S+|[A-Za-z_][\w$]*', p) if w not in KEYWORDS]
    return toks[-1] if toks else None


def width_str(p):
    return "1" if p["bits"] == 1 else f"[{p['msb']}:{p['lsb']}] ({p['bits']})"


def blank(text, regex):
    """把匹配到的片段替换成等长空白（保留偏移，便于后续扫描残留）。"""
    return regex.sub(lambda m: " " * len(m.group(0)), text)


# ----------------------------- 端口解析 -----------------------------

def parse_ports(ports_hdr, body):
    dirs = {}
    for txt in (ports_hdr, body):
        for m in DECL_RE.finditer(txt):
            d, rng, names = m.group(1), m.group(2), m.group(3)
            msb, lsb, bits = parse_range(rng)
            for nm in names.split(','):
                nm = ident(nm)
                if nm:
                    dirs[nm] = {"dir": d, "msb": msb, "lsb": lsb, "bits": bits}
    order, seen = [], set()
    for p in ports_hdr.split(','):
        nm = ident(p)
        if nm and nm not in seen:
            order.append(nm)
            seen.add(nm)
    ports = []
    for nm in order:
        ports.append({"name": nm, **dirs.get(nm, {"dir": "?", "msb": 0, "lsb": 0, "bits": 1})})
    for nm, info in dirs.items():
        if nm not in seen:
            ports.append({"name": nm, **info})
    return ports


# ----------------------------- 实例(instance)扫描 -----------------------------

def read_token(s, i):
    """读一个标识符 token（支持转义标识符 \\xxx ）。返回 (token 或 None, 新位置)。"""
    n = len(s)
    while i < n and s[i] in " \t\r\n":
        i += 1
    if i >= n:
        return None, i
    if s[i] == "\\":                      # 转义标识符，到空白结束
        j = i + 1
        while j < n and s[j] not in " \t\r\n":
            j += 1
        return s[i:j], j
    m = re.match(r'[A-Za-z_][\w$]*', s[i:])
    if m:
        return m.group(0), i + m.end()
    return None, i


def read_balanced(s, i, open_ch, close_ch):
    """s[i]==open_ch，返回 (内部文本, 结束后位置)。"""
    n = len(s)
    depth = 0
    start = i
    while i < n:
        c = s[i]
        if c == open_ch:
            depth += 1
        elif c == close_ch:
            depth -= 1
            if depth == 0:
                return s[start + 1:i], i + 1
        i += 1
    return s[start + 1:], n            # 不平衡：尽力而为


def split_top_commas(text):
    """按顶层逗号切分（跳过 {} [] () 内的逗号）。"""
    parts, depth, cur = [], 0, ""
    for c in text:
        if c in "{[(":
            depth += 1
        elif c in "}])":
            depth -= 1
        if c == "," and depth == 0:
            parts.append(cur)
            cur = ""
        else:
            cur += c
    if cur.strip():
        parts.append(cur)
    return parts


def extract_nets(expr):
    """从连接表达式抽出基础网络名（去位选/常量）。net[3:0] -> net；{a,b[1]} -> a,b。"""
    e = re.sub(r"\d*'[sS]?[bodhBODH][0-9a-fA-FxXzZ_?]+", " ", expr)   # sized 常量
    e = re.sub(r'\b\d+\b', ' ', e)                                    # 裸数字/索引
    out, seen = [], set()
    for tok in re.findall(r'\\\S+|[A-Za-z_][\w$]*', e):
        if tok in KEYWORDS:
            continue
        if tok not in seen:
            seen.add(tok)
            out.append(tok)
    return out


def parse_conns(conn_text):
    """解析一条例化的连接列表 -> [{pin, expr, nets}]（pin 为 None 表示位置连接）。"""
    conns = []
    for idx, part in enumerate(split_top_commas(conn_text)):
        s = part.strip()
        if s == "":
            continue
        m = re.match(r'\.\s*(\\\S+|[A-Za-z_][\w$]*)\s*\((.*)\)\s*$', s, re.DOTALL)
        if m:                                   # 具名 .pin(expr)
            pin, expr = m.group(1), m.group(2)
        else:                                   # 位置连接
            pin, expr = None, s
        conns.append({"pin": pin, "pos": idx, "expr": expr.strip(),
                      "nets": extract_nets(expr)})
    return conns


def scan_instances(work):
    """在(已抹掉端口/wire/assign的)模块体里扫描例化。返回 (instances, unparsed)。"""
    n = len(work)
    i = 0
    instances = []
    consumed = bytearray(n)               # 标记已解析区域，剩下的算 unparsed
    while i < n:
        while i < n and work[i] in " \t\r\n":
            i += 1
        if i >= n:
            break
        start = i
        typ, i1 = read_token(work, i)
        if typ is None:
            i += 1
            continue
        name, i2 = read_token(work, i1)
        if name is None:
            i = i1
            _mark(consumed, start, i1)
            continue
        j = i2
        while j < n and work[j] in " \t\r\n":
            j += 1
        array = None
        if j < n and work[j] == "[":       # 数组例化 type name [3:0] (...)
            array, j = read_balanced(work, j, "[", "]")
            array = "[" + array + "]"
            while j < n and work[j] in " \t\r\n":
                j += 1
        if j >= n or work[j] != "(":       # 不是例化 -> 跳到分号
            k = work.find(";", start)
            k = k + 1 if k >= 0 else n
            i = k
            continue
        conn_text, jafter = read_balanced(work, j, "(", ")")
        k = jafter
        while k < n and work[k] in " \t\r\n":
            k += 1
        if k < n and work[k] == ";":
            k += 1
        instances.append({
            "type": typ,
            "name": name,
            "array": array,
            "connections": parse_conns(conn_text),
        })
        _mark(consumed, start, k)
        i = k
    # 未解析残留
    unparsed = []
    buf = ""
    for idx in range(n):
        if not consumed[idx] and work[idx] not in " \t\r\n":
            buf += work[idx]
        else:
            if buf.strip():
                unparsed.append(buf.strip())
            buf = ""
    if buf.strip():
        unparsed.append(buf.strip())
    # 合并/截断
    unparsed = [u for u in unparsed if len(u) > 1][:40]
    return instances, unparsed


def _mark(consumed, a, b):
    for k in range(a, min(b, len(consumed))):
        consumed[k] = 1


# ----------------------------- 完整连接解析 -----------------------------

def parse_connectivity(name, hdr, body):
    ports = parse_ports(hdr, body)

    wires = []
    for m in WIRE_RE.finditer(body):
        rng, names = m.group(2), m.group(3)
        msb, lsb, bits = parse_range(rng)
        for nm in names.split(','):
            nm = ident(nm)
            if nm:
                wires.append({"name": nm, "msb": msb, "lsb": lsb, "bits": bits})

    assigns = []
    for m in ASSIGN_RE.finditer(body):
        s = m.group(1)
        if "=" in s:
            lhs, rhs = s.split("=", 1)
            assigns.append({"lhs": extract_nets(lhs), "rhs": extract_nets(rhs)})

    work = blank(body, DECL_RE)
    work = blank(work, WIRE_RE)
    work = blank(work, ASSIGN_RE)
    instances, unparsed = scan_instances(work)

    # net_index：网络 -> 触点列表（顶层端口 或 实例引脚）
    net_index = {}

    def add(net, tag):
        net_index.setdefault(net, []).append(tag)

    port_dir = {p["name"]: p["dir"] for p in ports}
    for p in ports:
        add(p["name"], f"TOP:{p['dir']}:{p['name']}")
    for a in assigns:
        for L in a["lhs"]:
            for R in a["rhs"]:
                add(L, f"ASSIGN<= {R}")
                add(R, f"ASSIGN=> {L}")
    n_conn = 0
    for inst in instances:
        for c in inst["connections"]:
            n_conn += 1
            pin = c["pin"] if c["pin"] is not None else f"#{c['pos']}"
            for net in c["nets"]:
                add(net, f"{inst['type']}/{inst['name']}.{pin}")

    return {
        "name": name,
        "n_ports": len(ports),
        "n_wires": len(wires),
        "n_instances": len(instances),
        "n_connections": n_conn,
        "n_unparsed": len(unparsed),
        "ports": [{"name": p["name"], "dir": p["dir"], "bits": p["bits"]} for p in ports],
        "wires": wires,
        "assigns": assigns,
        "instances": instances,
        "net_index": net_index,
        "unparsed": unparsed,
    }


# ----------------------------- 层级/向上追踪 -----------------------------

def module_instances(body):
    work = blank(body, DECL_RE)
    work = blank(work, WIRE_RE)
    work = blank(work, ASSIGN_RE)
    insts, _ = scan_instances(work)
    return insts


def build_graph(mods, order):
    """返回 uses(child_type -> [(parent, inst, conns)]), roots, externals(叶子/未定义)。"""
    uses = {}
    for pname in order:
        hdr, body = mods[pname]
        for it in module_instances(body):
            uses.setdefault(it["type"], []).append((pname, it["name"], it["connections"]))
    defined = set(mods)
    used = set(uses)
    roots = [m for m in order if m not in used]        # 定义了但没被谁例化 = 顶
    externals = sorted(used - defined)                  # 被例化但没定义 = 叶子(.va)/外部
    return uses, roots, externals


def paths_to_root(target, uses, max_paths=30):
    """target 到根的所有层级路径。每条 = [(childtype, instname) ...] 顶->下, 末尾是 target。"""
    out = []

    def rec(node, below, seen):
        parents = uses.get(node)
        if not parents:
            out.append((node, below))                   # node 是根
            return
        for (p, inst, _) in parents:
            if p in seen or len(out) >= max_paths:
                out.append((f"<cycle {p}>", below))
                continue
            rec(p, [(node, inst)] + below, seen | {p})

    rec(target, [], {target})
    return out


def ports_via_assign(module, net, assign_adj, pset):
    """在 module 内，net 通过 assign 别名能到达哪些本模块端口（BFS 传递闭包）。"""
    adj = assign_adj.get(module, {})
    ports = pset.get(module, set())
    seen, stack, hits = {net}, [net], []
    while stack:
        x = stack.pop()
        for y in adj.get(x, ()):
            if y in seen:
                continue
            seen.add(y)
            if y in ports:
                hits.append(y)
            stack.append(y)
    return hits


def get_netindex(module, mods, cache):
    """惰性计算并缓存某模块的 net_index（网络 -> 触点列表）。"""
    if module not in cache:
        hdr, body = mods[module]
        cache[module] = parse_connectivity(module, hdr, body)["net_index"]
    return cache[module]


def trace_up(module, port, uses, porder, pset, assign_adj, mods, netcache,
             depth=0, max_depth=14):
    """从 module.port 向上追：连线若是父端口就继续爬；若是内部网络则试 assign 别名到父端口；
    直到文件顶(根)或在某层被真正内部消耗。"""
    if depth > max_depth:
        return [{"note": "达到最大深度"}]
    parents = uses.get(module)
    if not parents:
        return [{"top_pin": port, "in": module}]        # 根：port 就是文件顶层引脚
    out = []
    for (p, inst, conns) in parents:
        found = None
        for c in conns:
            if c["pin"] == port:
                found = c
                break
            if c["pin"] is None:                        # 位置连接：按目标端口顺序匹配
                order = porder.get(module, [])
                if 0 <= c["pos"] < len(order) and order[c["pos"]] == port:
                    found = c
                    break
        if not found:
            out.append({"parent": p, "inst": inst, "note": "此例化未连该端口"})
            continue
        for net in found["nets"]:
            step = {"parent": p, "inst": inst, "net": net,
                    "net_is_parent_port": net in pset.get(p, set())}
            if step["net_is_parent_port"]:
                step["up"] = trace_up(p, net, uses, porder, pset, assign_adj, mods, netcache,
                                      depth + 1, max_depth)
            else:
                via = ports_via_assign(p, net, assign_adj, pset)
                if via:
                    step["via_assign_to"] = via
                    step["up"] = []
                    for pv in via:
                        step["up"] += trace_up(p, pv, uses, porder, pset, assign_adj, mods,
                                               netcache, depth + 1, max_depth)
                else:
                    # 终点内部网络：列出这根网上的兄弟触点（真正驱动/负载它的 cell 实例）
                    ni = get_netindex(p, mods, netcache)
                    own = f"{module}/{inst}."
                    step["siblings"] = [t for t in ni.get(net, []) if not t.startswith(own)][:12]
            out.append(step)
    return out


def flatten_trace(steps, prefix):
    """把 trace_up 的树压成可读的链字符串列表。"""
    lines = []
    for s in steps:
        if "top_pin" in s:
            lines.append(f"{prefix}  ==> 顶层引脚 {s['top_pin']} @ {s['in']}")
        elif "note" in s:
            lines.append(f"{prefix}  [{s['note']}]" + (f" @ {s.get('parent','')}" if s.get('parent') else ""))
        else:
            here = f"{prefix} ⟵ {s['parent']}.{s['inst']} (net {s['net']})"
            if s.get("net_is_parent_port") and s.get("up"):
                lines += flatten_trace(s["up"], here)
            elif s.get("via_assign_to"):
                here += f" [assign→ {', '.join(s['via_assign_to'])}]"
                lines += flatten_trace(s["up"], here)
            else:
                tail = "  [内部网络"
                if s.get("siblings"):
                    tail += "; 相邻: " + ", ".join(s["siblings"])
                tail += "]"
                lines.append(here + tail)
    return lines


# ----------------------------- 端口模式输出 -----------------------------

def summarize_ports(name, ports):
    counts = {"input": 0, "output": 0, "inout": 0, "?": 0}
    for p in ports:
        counts[p["dir"]] = counts.get(p["dir"], 0) + 1
    supply = [p for p in ports if POWER_RE.match(p["name"])]
    sig = [p for p in ports if not POWER_RE.match(p["name"])]
    return {
        "name": name, "n_ports": len(ports), "counts": counts, "ports": ports,
        "control_inputs": [p["name"] for p in sig if p["dir"] == "input"],
        "outputs": [p["name"] for p in sig if p["dir"] == "output"],
        "inout_signals": [p["name"] for p in sig if p["dir"] == "inout"],
        "supply": [p["name"] for p in supply],
    }


def print_ports(s):
    print("=" * 90)
    c = s["counts"]
    print(f"[MODULE] {s['name']}   端口 {s['n_ports']}  "
          f"(input {c['input']}, output {c['output']}, inout {c['inout']}"
          + (f", ?{c['?']}" if c.get('?') else "") + ")")

    by = {p["name"]: p for p in s["ports"]}

    def block(title, names):
        if not names:
            return
        print(f"\n  {title} ({len(names)}):")
        for nm in names:
            p = by.get(nm)
            w = f"  {width_str(p)}" if (p and p['bits'] > 1) else ""
            print(f"    {nm}{w}")

    block("控制输入 input (非电源)", s["control_inputs"])
    block("输出 output", s["outputs"])
    block("模拟 inout (非电源, 如 LO/tank/buf)", s["inout_signals"])
    if s["supply"]:
        print(f"\n  电源/地 ({len(s['supply'])}): " + ", ".join(s["supply"]))


def compact_conn(r):
    """把 parse_connectivity 结果压成紧凑无损结构：
    去掉 net_index(可由 instances+ports+assigns 重建)、每连接的 pos(=顺序)/nets(=从 expr 解析)。
    连接表示为 [pin, expr]，pin=null 为位置连接。expr 保留(含位选/拼接)故无损。"""
    return {
        "name": r["name"],
        "stats": {"ports": r["n_ports"], "wires": r["n_wires"],
                  "instances": r["n_instances"], "connections": r["n_connections"],
                  "unparsed": r["n_unparsed"]},
        "ports": [[p["name"], p["dir"], p["bits"]] for p in r["ports"]],
        "wires": [[w["name"], w["bits"]] for w in r["wires"]],
        "assigns": [[a["lhs"], a["rhs"]] for a in r["assigns"]],
        "instances": [{"t": i["type"], "n": i["name"], "a": i["array"],
                       "c": [[c["pin"], c["expr"]] for c in i["connections"]]}
                      for i in r["instances"]],
        "unparsed": r["unparsed"],
    }


def print_connectivity(r):
    print("=" * 90)
    print(f"[MODULE] {r['name']}   端口 {r['n_ports']}, wire {r['n_wires']}, "
          f"instance {r['n_instances']}, 连线 {r['n_connections']}"
          + (f"   ⚠ unparsed {r['n_unparsed']}" if r["n_unparsed"] else "   ✓ 无残留"))
    if r["instances"]:
        print("  例化（type  name  连线数）:")
        for inst in r["instances"][:60]:
            arr = inst["array"] or ""
            print(f"    {inst['type']:<34} {inst['name']}{arr}  ({len(inst['connections'])})")
        if len(r["instances"]) > 60:
            print(f"    ... 其余 {len(r['instances'])-60} 个见 JSON")
    if r["n_unparsed"]:
        print("  ⚠ 未识别片段(前几条，用来核对语法)：")
        for u in r["unparsed"][:8]:
            print("    " + (u[:80] + ("…" if len(u) > 80 else "")))


# ----------------------------- 模块查找 & main -----------------------------

def find_modules(text):
    hdr_re = re.compile(r'\bmodule\s+(\w+)\s*\((.*?)\)\s*;', re.DOTALL)
    end_re = re.compile(r'\bendmodule\b')
    mods, order = {}, []
    for m in hdr_re.finditer(text):
        nm, hdr, start = m.group(1), m.group(2), m.end()
        e = end_re.search(text, start)
        body = text[start:e.start()] if e else text[start:]
        mods[nm] = (hdr, body)
        order.append(nm)
    return mods, order


def main():
    ap = argparse.ArgumentParser(description="抽取指定模块的端口 I/O 与内部连接")
    ap.add_argument("path", help="netlist 文件 (.vh/.v/.va 等)")
    ap.add_argument("--modules", default=None, help="逗号分隔模块名；缺省读本地 config，无则需显式指定")
    ap.add_argument("--json", default=None, help="导出 JSON 到该文件")
    ap.add_argument("--list", action="store_true", help="只列出文件里所有模块名")
    ap.add_argument("--connections", action="store_true",
                    help="连接模式：抽 端口+wire+instance+连线 并建 net_index")
    ap.add_argument("--compact", action="store_true",
                    help="配 --connections：输出紧凑无损 JSON（去 net_index/pos/nets，紧凑排版），大幅减小体积")
    ap.add_argument("--tree", action="store_true",
                    help="层级树：目标模块怎么从根(文件顶)一层层例化下来")
    ap.add_argument("--uptrace", action="store_true",
                    help="向上追踪：目标模块每个端口连到父层哪根网络、一直追到文件顶层引脚")
    ap.add_argument("--body", default=None, help="打印指定模块体的原文（核对语法用）")
    ap.add_argument("--head", type=int, default=80, help="配 --body：打印前 N 行")
    args = ap.parse_args()

    if not os.path.isfile(args.path):
        sys.exit(f"找不到文件: {args.path}")
    with open(args.path, encoding="utf-8", errors="replace") as f:
        raw = f.read()
    text = strip_comments(raw)

    mods, order = find_modules(text)
    print(f"文件: {args.path}")
    print(f"共找到 {len(order)} 个 module\n")

    if args.list:
        for nm in order:
            print("  " + nm)
        return

    if args.body:
        if args.body not in mods:
            sys.exit(f"没有模块 {args.body}；用 --list 查看")
        hdr, body = mods[args.body]
        print(f"----- {args.body} 模块体原文（前 {args.head} 行）-----")
        for ln in body.splitlines()[:args.head]:
            if ln.strip():
                print(ln)
        return

    targets = ([m.strip() for m in args.modules.split(",") if m.strip()]
               if args.modules else _load_default_modules())
    if not targets and not args.list:
        sys.exit("未指定目标模块：用 --modules A,B,C，或放 private/tool_config/extract_ports.json")

    # ---- 层级树 / 向上追踪：需要全局例化关系 ----
    if args.tree or args.uptrace:
        not_found = []
        uses, roots, externals = build_graph(mods, order)
        porder = {nm: [p["name"] for p in parse_ports(h, b)] for nm, (h, b) in mods.items()}
        pset = {nm: set(v) for nm, v in porder.items()}
        assign_adj = {}
        for nm, (h, b) in mods.items():
            adj = {}
            for mm in ASSIGN_RE.finditer(b):
                s = mm.group(1)
                if "=" in s:
                    L, R = s.split("=", 1)
                    for a in extract_nets(L):
                        for c in extract_nets(R):
                            adj.setdefault(a, set()).add(c)
                            adj.setdefault(c, set()).add(a)
            assign_adj[nm] = adj
        netcache = {}
        print(f"文件顶层(根)模块: {', '.join(roots) if roots else '(未识别)'}")
        print(f"叶子/外部(被例化但本文件未定义)约 {len(externals)} 种\n")

        tree_out, up_out = [], []
        for nm in targets:
            if nm not in mods:
                not_found.append(nm)
                continue
            print("=" * 90)
            print(f"[目标] {nm}")
            paths = paths_to_root(nm, uses)
            tree_out.append({"module": nm, "paths": [
                {"root": root, "chain": [{"type": t, "inst": i} for (t, i) in below]}
                for (root, below) in paths]})
            if not paths:
                print("  (没找到父模块——它可能就是根)")
            for (root, below) in paths:
                print(f"  {root}")
                indent = "  "
                for (t, inst) in below:
                    indent += "  "
                    print(f"{indent}└ {inst} : {t}")

            if args.uptrace:
                hdr, body = mods[nm]
                tports = parse_ports(hdr, body)
                print(f"\n  ── 端口向上追踪（到文件顶层引脚）──")
                per_port = []
                for p in tports:
                    if POWER_RE.match(p["name"]):
                        continue
                    steps = trace_up(nm, p["name"], uses, porder, pset, assign_adj, mods, netcache)
                    lines = flatten_trace(steps, f"    {p['name']}({p['dir']})")
                    per_port.append({"port": p["name"], "dir": p["dir"], "steps": steps})
                    for ln in lines:
                        print(ln)
                up_out.append({"module": nm, "ports": per_port})

        if not_found:
            print("\n" + "!" * 90)
            print("以下模块名没找到（确认拼写，或用 --list）：")
            for nm in not_found:
                print("  " + nm)
        if args.json:
            with open(args.json, "w", encoding="utf-8") as f:
                json.dump({"file": os.path.basename(args.path),
                           "roots": roots, "n_externals": len(externals),
                           "tree": tree_out, "uptrace": up_out,
                           "not_found": not_found}, f, ensure_ascii=False, indent=2)
            print(f"\n已导出 JSON: {args.json}")
        return

    results, not_found = [], []
    for nm in targets:
        if nm not in mods:
            not_found.append(nm)
            continue
        hdr, body = mods[nm]
        if args.connections:
            r = parse_connectivity(nm, hdr, body)
            results.append(r)
            print_connectivity(r)
        else:
            s = summarize_ports(nm, parse_ports(hdr, body))
            results.append(s)
            print_ports(s)

    if not_found:
        print("\n" + "!" * 90)
        print("以下模块名没找到（确认拼写，或用 --list）：")
        for nm in not_found:
            print("  " + nm)

    if args.json:
        if args.connections and args.compact:
            payload = {"file": os.path.basename(args.path), "mode": "connections-compact",
                       "note": ("省略 net_index(可由 instances+ports+assigns 重建); "
                                "连接=[pin,expr](pin=null为位置连接); pos=数组顺序、nets=从expr解析，均可重建。无损。"),
                       "modules": [compact_conn(r) for r in results], "not_found": not_found}
            text = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
        else:
            text = json.dumps({"file": os.path.basename(args.path),
                               "mode": "connections" if args.connections else "ports",
                               "modules": results, "not_found": not_found},
                              ensure_ascii=False, indent=2)
        with open(args.json, "w", encoding="utf-8") as f:
            f.write(text)
        nb = len(text.encode("utf-8"))
        print(f"\n已导出 JSON: {args.json}  ({nb/1024:.1f} KB)")


if __name__ == "__main__":
    main()
