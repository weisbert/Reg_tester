#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""regtool.py — GUI 信号流工具的启动器（阶段二 M2）。

两种启动方式，前端同一套代码（webapp/）：
  --serve   stdlib http.server，直接读写 project 目录（主用；红区/本机）。
  --bundle  把 webapp + project 数据打进单个自包含 HTML（应急/黄区，无网络依赖）。

只依赖标准库。脚本本身不含真实信号名/地址（数据都在 project 目录里，gitignore）。

用法：
    python regtool.py --serve --project projects/adpll_demo            # 起服务（默认 :8765）
    python regtool.py --serve --project projects/adpll_demo --open     # 顺便开浏览器
    python regtool.py --bundle --project projects/adpll_demo --out dist/adpll.html
"""
import argparse
import json
import os
import re
import subprocess
import sys
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse, unquote

HERE = os.path.dirname(os.path.abspath(__file__))
WEBAPP = os.path.join(HERE, "webapp")

try:
    import gen_testcase
except Exception:  # 允许从别处运行
    sys.path.insert(0, HERE)
    import gen_testcase

MIME = {
    ".html": "text/html; charset=utf-8", ".js": "application/javascript; charset=utf-8",
    ".css": "text/css; charset=utf-8", ".json": "application/json; charset=utf-8",
    ".svg": "image/svg+xml", ".ico": "image/x-icon",
}

SAFE_ID = re.compile(r"^[A-Za-z0-9_.\-]+$")
BAD_BODY = object()          # _read_body 解析失败的哨兵（区别于"无 body"=None）
MAX_BODY = 8 * 1024 * 1024   # 请求体上限，防不受限 Content-Length


def valid_id(mid):
    return bool(SAFE_ID.match(mid)) and ".." not in mid


# ------------------------------------------------------------------ project io
class Project:
    def __init__(self, root):
        self.root = os.path.abspath(root)
        # 目录不存在则**脚手架新建**（P2.2 建库向导：一个指令即可对着空工程起 GUI）
        os.makedirs(self.root, exist_ok=True)
        self.modes_dir = os.path.join(self.root, "modes")
        self.tc_dir = os.path.join(self.root, "testcases")
        os.makedirs(self.modes_dir, exist_ok=True)
        os.makedirs(self.tc_dir, exist_ok=True)

    def _read(self, name, default=None):
        p = os.path.join(self.root, name)
        if not os.path.exists(p):
            return default
        with open(p, "r", encoding="utf-8") as f:
            return json.load(f)

    def _write(self, name, obj):
        p = os.path.join(self.root, name)
        os.makedirs(os.path.dirname(p), exist_ok=True)
        with open(p, "w", encoding="utf-8", newline="\n") as f:
            json.dump(obj, f, ensure_ascii=False, indent=1)

    def project_json(self):
        return self._read("project.json", {"name": os.path.basename(self.root)})

    def signal_reg_map(self):
        return self._read("signal_reg_map.json", {})

    def matching(self):
        return (self.project_json() or {}).get("matching", {}) or {}

    def save_matching(self, alias, logic_derived):
        """把确认的 alias / logic_derived 写回 project.json 的 matching 段（其余键保留）。"""
        proj = self.project_json() or {"name": os.path.basename(self.root)}
        mt = dict(proj.get("matching", {}) or {})
        mt["alias"] = alias
        mt["logic_derived"] = logic_derived
        proj["matching"] = mt
        self._write("project.json", proj)

    def rebuild(self):
        """按序跑 make_mock_regmap → build_regmap → build_flowgraph（--project 自身），重生派生物。
        stdlib subprocess；make_mock 不出 xlsx 故无需 openpyxl。返回 {ok, logs[, failed]}。"""
        scripts = ["make_mock_regmap.py", "build_regmap.py", "build_flowgraph.py"]
        logs = []
        for s in scripts:
            cmd = [sys.executable, os.path.join(HERE, s), "--project", self.root]
            r = subprocess.run(cmd, capture_output=True, text=True,
                               encoding="utf-8", errors="replace")
            logs.append({"script": s, "code": r.returncode,
                         "out": (r.stdout or "")[-1500:], "err": (r.stderr or "")[-1500:]})
            if r.returncode != 0:
                return {"ok": False, "failed": s, "logs": logs}
        return {"ok": True, "logs": logs}

    def flowgraph(self):
        return self._read("flowgraph.json", {})

    def regmap(self):
        return self._read("regmap.json", {})

    def layout(self):
        return self._read("layout.json",
                          {"schema_version": "layout/1", "positions": {}, "collapsed": [],
                           "hidden": [], "shown": [], "symbol_override": {}, "edge_flip": {},
                           "notes": {}, "expanded": []})

    def save_layout(self, obj):
        self._write("layout.json", obj)

    def list_modes(self):
        out = []
        for fn in sorted(os.listdir(self.modes_dir)):
            if fn.endswith(".json"):
                try:
                    m = self._read(os.path.join("modes", fn)) or {}
                    # id 以文件名为准（与 read_mode/generate/delete/bundle 键一致），name/组从内容取
                    out.append({"id": fn[:-5], "name": m.get("name", fn[:-5]),
                                "reg_group": m.get("reg_group", "BT")})
                except Exception:
                    pass
        return out

    def read_mode(self, mid):
        return self._read(os.path.join("modes", mid + ".json"))

    def save_mode(self, mid, obj):
        obj["id"] = mid
        self._write(os.path.join("modes", mid + ".json"), obj)

    def delete_mode(self, mid):
        p = os.path.join(self.modes_dir, mid + ".json")
        if os.path.exists(p):
            os.remove(p)

    def generate(self, mid):
        mode = self.read_mode(mid)
        if not mode:
            return None
        tc = gen_testcase.generate(self.flowgraph(), self.regmap(), mode)
        return {"testcase": tc,
                "ate": gen_testcase.render_ate(tc),
                "html": gen_testcase.render_debug_html(tc, self.flowgraph())}

    def export(self, mid):
        res = self.generate(mid)
        if not res:
            return None
        self._write(os.path.join("testcases", mid + ".json"), res["testcase"])
        with open(os.path.join(self.tc_dir, mid + ".ate.txt"), "w", encoding="utf-8", newline="\n") as f:
            f.write(res["ate"])
        with open(os.path.join(self.tc_dir, mid + ".debug.html"), "w", encoding="utf-8", newline="\n") as f:
            f.write(res["html"])
        return {"ok": True, "dir": self.tc_dir}

    # -------------------------------------------------- P2.2 建库向导（新建工程 / 导入 / 导出）
    ART_DEFAULTS = {
        "control_signals": "control_signals.json", "pll_rows": "pll_rows.json",
        "conn": "conn.json", "expand_conn": ["expand_conn.json"],
        "signal_reg_map": "signal_reg_map.json", "regmap": "regmap.json",
        "flowgraph": "flowgraph.json",
    }
    CONFIG_KEYS = ("name", "description", "netlist", "regbook", "flowgraph_rules")

    def needs_setup(self):
        """无 flowgraph 节点 = 还没建库，前端应引导到「工程」向导。"""
        return not (self.flowgraph().get("nodes"))

    def save_project_config(self, cfg):
        """把 GUI 里填的芯片配置合并进 project.json（只收已知段；artifacts 补默认固定名）。"""
        proj = self.project_json() or {}
        proj["schema_version"] = proj.get("schema_version", "project/2")
        for k in self.CONFIG_KEYS:
            if k in cfg and cfg[k] is not None:
                proj[k] = cfg[k]
        proj.setdefault("name", os.path.basename(self.root))
        proj.setdefault("matching", proj.get("matching", {}) or {})
        art = dict(proj.get("artifacts") or {})
        for k, v in self.ART_DEFAULTS.items():
            art.setdefault(k, v)
        proj["artifacts"] = art
        self._write("project.json", proj)
        return proj

    def control_signals(self):
        return self._read("control_signals.json", {})

    def save_control_signals(self, obj):
        self._write("control_signals.json", obj)

    def _run(self, arglist):
        r = subprocess.run([sys.executable] + arglist, capture_output=True,
                           text=True, encoding="utf-8", errors="replace")
        return {"cmd": "python " + " ".join(os.path.basename(a) if a.endswith(".py") else a for a in arglist),
                "code": r.returncode, "out": (r.stdout or "")[-1200:], "err": (r.stderr or "")[-1200:]}

    def import_netlist(self, path):
        """跑 extract_ports 抽 conn.json（+ expand_conn.json），再抽 uptrace 生成 control_signals 候选。
        不覆盖已存在的 control_signals.json——候选交前端展示、由人确认后 save。"""
        if not (path and os.path.isfile(path)):
            return {"ok": False, "error": "netlist 文件不存在：%s" % path}
        proj = self.project_json() or {}
        nl = proj.get("netlist") or {}
        targets = nl.get("target_modules") or []
        if not targets:
            return {"ok": False, "error": "netlist.target_modules 为空——先在配置里填目标模块并保存"}
        ep = os.path.join(HERE, "extract_ports.py")
        logs = []
        logs.append(self._run([ep, path, "--project", self.root, "--connections", "--compact",
                               "--json", os.path.join(self.root, "conn.json")]))
        if logs[-1]["code"] != 0:
            return {"ok": False, "error": "extract_ports conn 失败", "logs": logs}
        if nl.get("expand_submodules"):
            logs.append(self._run([ep, path, "--project", self.root, "--expand", "--connections",
                                   "--compact", "--json", os.path.join(self.root, "expand_conn.json")]))
            if logs[-1]["code"] != 0:
                return {"ok": False, "error": "extract_ports expand 失败", "logs": logs}
        candidate = {"primary_current_related": [], "config_secondary": []}
        up_tmp = os.path.join(self.root, "_uptrace.json")
        logs.append(self._run([ep, path, "--project", self.root, "--tree", "--uptrace", "--json", up_tmp]))
        if logs[-1]["code"] == 0:
            try:
                candidate = seed_control_signals(self._read("_uptrace.json", {}), proj)
            except Exception as e:
                candidate = {"note": "候选生成失败：%s" % e,
                             "primary_current_related": [], "config_secondary": []}
        try:
            os.remove(up_tmp)
        except OSError:
            pass
        conn = self._read("conn.json", {})
        return {"ok": True, "logs": logs,
                "conn_modules": len(conn.get("modules") or []),
                "candidate": candidate,
                "has_control_signals": os.path.exists(os.path.join(self.root, "control_signals.json"))}

    def import_excel(self, path, sheet, rowdump):
        """跑 explore_excel --rowdump 把寄存器 sheet 的行区间抽成 pll_rows.json。"""
        if not (path and os.path.isfile(path)):
            return {"ok": False, "error": "Excel 文件不存在：%s" % path}
        rb = (self.project_json() or {}).get("regbook") or {}
        sheet = sheet or rb.get("sheet_name")
        if not sheet:
            return {"ok": False, "error": "缺 sheet 名（表单或 regbook.sheet_name）"}
        if not rowdump or ":" not in str(rowdump):
            return {"ok": False, "error": "缺行区间 START:END（先用 explore_excel --index/--schema 看结构）"}
        ee = os.path.join(HERE, "explore_excel.py")
        log = self._run([ee, path, "--sheet", str(sheet), "--rowdump", str(rowdump),
                        "--dump", os.path.join(self.root, "pll_rows.json")])
        if log["code"] != 0:
            return {"ok": False, "error": "explore_excel 失败（openpyxl 装了吗？sheet 名/行区间对吗？）",
                    "logs": [log]}
        doc = self._read("pll_rows.json", {})
        return {"ok": True, "logs": [log], "rows": len(doc.get("rows") or []),
                "sheet": doc.get("sheet"), "cols": doc.get("n_cols")}

    def export_bundle(self):
        """工程配置 + control_signals + 全部模式，打成一段可复制文本（供 air-gap 贴回）。"""
        return {"project": self.project_json(),
                "control_signals": self.control_signals(),
                "modes": {m["id"]: self.read_mode(m["id"]) for m in self.list_modes()}}


# --------------------------------------------------------- control_signals 候选生成（best-effort）
def _module_tag(module, bands):
    for k in bands:
        if k and k.lower() in (module or "").lower():
            return k
    return (module or "").split("_")[-1] or module


def _classify_pin(port):
    """按引脚后缀/词根启发式分类：返回 (category, bucket)。bucket='prim'|'sec'|None。"""
    low = (port or "").lower()
    if low.endswith("_sel") or "_sel_" in low or "_mode" in low or "_mux" in low or "div_sel" in low:
        return "config", "sec"
    if low.endswith("_en") or "_en_" in low:
        return "en", "prim"
    if "ictrl" in low:
        return "ictrl", "prim"
    if "bias" in low:
        return "bias", "prim"
    if "tune" in low or "itune" in low:
        return "tune", "prim"
    if "isource" in low:
        return "isource", "prim"
    if low.endswith("_ctrl") or "_ctrl_" in low:
        return "current", "prim"
    return None, None


def seed_control_signals(uptrace_doc, proj):
    """从 extract_ports 的 uptrace（目标模块控制脚→顶层引脚）生成 control_signals 候选。
    只抓 dir=input & dest=top_pin 的控制脚（经内部逻辑的漏网脚需人工补）。category/desc 交人工确认。"""
    bands = list(((proj.get("flowgraph_rules") or {}).get("module_bands") or {}).keys())
    prim, sec = {}, {}
    for entry in (uptrace_doc.get("uptrace") or []):
        tag = _module_tag(entry.get("module", ""), bands)
        for p in (entry.get("ports") or []):
            if p.get("dir") != "input" or p.get("dest") != "top_pin":
                continue
            net, port = p.get("net"), p.get("port")
            if not net:
                continue
            cat, bucket = _classify_pin(port)
            if bucket is None:
                continue
            store = prim if bucket == "prim" else sec
            it = store.setdefault(net, {"reg_net": net, "category": cat, "drives": [], "desc": ""})
            d = "%s.%s" % (tag, port)
            if d not in it["drives"]:
                it["drives"].append(d)
    return {"note": "候选：从 netlist uptrace 自动抓（top_pin 输入控制脚）。category/desc 需人工确认；"
                    "经内部逻辑门控的控制脚（非 top_pin）不在此，需手工补。",
            "primary_current_related": list(prim.values()),
            "config_secondary": list(sec.values())}


# ------------------------------------------------------------------ http server
def make_handler(project):
    class Handler(BaseHTTPRequestHandler):
        server_version = "regtool/1"

        def log_message(self, fmt, *args):
            sys.stderr.write("  %s - %s\n" % (self.address_string(), fmt % args))

        # -- helpers
        def _send(self, code, body, ctype="application/json; charset=utf-8"):
            if isinstance(body, (dict, list)):
                body = json.dumps(body, ensure_ascii=False).encode("utf-8")
            elif isinstance(body, str):
                body = body.encode("utf-8")
            self.send_response(code)
            self.send_header("Content-Type", ctype)
            self.send_header("Content-Length", str(len(body)))
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            self.wfile.write(body)

        def _read_body(self):
            n = int(self.headers.get("Content-Length", 0))
            if not n:
                return None
            if n > MAX_BODY:
                return BAD_BODY
            raw = self.rfile.read(n)
            try:
                return json.loads(raw.decode("utf-8"))
            except Exception:
                return BAD_BODY

        def _static(self, path):
            if path in ("/", ""):
                path = "/index.html"
            fn = os.path.normpath(os.path.join(WEBAPP, path.lstrip("/")))
            if not (fn == WEBAPP or fn.startswith(WEBAPP + os.sep)) or not os.path.isfile(fn):
                return self._send(404, {"error": "not found"})
            ext = os.path.splitext(fn)[1]
            with open(fn, "rb") as f:
                data = f.read()
            self._send(200, data, MIME.get(ext, "application/octet-stream"))

        # -- routing
        def do_GET(self):
            u = urlparse(self.path)
            p = unquote(u.path)
            if not p.startswith("/api/"):
                return self._static(p)
            try:
                if p == "/api/bootstrap":
                    return self._send(200, {
                        "project": project.project_json(),
                        "flowgraph": project.flowgraph(),
                        "regmap": project.regmap(),
                        "layout": project.layout(),
                        "modes": project.list_modes(),
                        "needs_setup": project.needs_setup(),
                        "backend": "serve",
                    })
                if p == "/api/project/export":
                    return self._send(200, project.export_bundle())
                if p == "/api/flowgraph":
                    return self._send(200, project.flowgraph())
                if p == "/api/regmap":
                    return self._send(200, project.regmap())
                if p == "/api/layout":
                    return self._send(200, project.layout())
                if p == "/api/modes":
                    return self._send(200, project.list_modes())
                if p == "/api/matching":
                    return self._send(200, {"signal_reg_map": project.signal_reg_map(),
                                            "matching": project.matching()})
                m = re.match(r"^/api/mode/([^/]+)$", p)
                if m:
                    mid = m.group(1)
                    if not valid_id(mid):
                        return self._send(404, {"error": "bad id"})
                    mode = project.read_mode(mid)
                    return self._send(200 if mode else 404, mode or {"error": "no mode"})
                m = re.match(r"^/api/generate/([^/]+)$", p)
                if m:
                    if not valid_id(m.group(1)):
                        return self._send(404, {"error": "bad id"})
                    res = project.generate(m.group(1))
                    return self._send(200 if res else 404, res or {"error": "no mode"})
                return self._send(404, {"error": "unknown api"})
            except Exception as e:
                return self._send(500, {"error": str(e)})

        def do_PUT(self):
            u = urlparse(self.path)
            p = unquote(u.path)
            body = self._read_body()
            if body is BAD_BODY:
                return self._send(400, {"error": "malformed or oversized json body"})
            try:
                if p == "/api/layout":
                    if not isinstance(body, dict):
                        return self._send(400, {"error": "expect json object"})
                    project.save_layout(body)
                    return self._send(200, {"ok": True})
                m = re.match(r"^/api/mode/([^/]+)$", p)
                if m:
                    mid = m.group(1)
                    if not valid_id(mid):
                        return self._send(400, {"error": "bad id"})
                    if not isinstance(body, dict):
                        return self._send(400, {"error": "expect json object"})
                    project.save_mode(mid, body)
                    return self._send(200, {"ok": True, "id": mid})
                return self._send(404, {"error": "unknown api"})
            except Exception as e:
                return self._send(500, {"error": str(e)})

        def do_POST(self):
            u = urlparse(self.path)
            p = unquote(u.path)
            try:
                m = re.match(r"^/api/export/([^/]+)$", p)
                if m:
                    if not valid_id(m.group(1)):
                        return self._send(404, {"error": "bad id"})
                    res = project.export(m.group(1))
                    return self._send(200 if res else 404, res or {"error": "no mode"})
                # ---- P2.2 建库向导路由 ----
                if p == "/api/project/config":
                    body = self._read_body()
                    if body is BAD_BODY or not isinstance(body, dict):
                        return self._send(400, {"error": "malformed json body"})
                    return self._send(200, {"ok": True, "project": project.save_project_config(body)})
                if p == "/api/import/netlist":
                    body = self._read_body()
                    if body is BAD_BODY or not isinstance(body, dict):
                        return self._send(400, {"error": "malformed json body"})
                    res = project.import_netlist(body.get("path"))
                    return self._send(200 if res.get("ok") else 400, res)
                if p == "/api/import/excel":
                    body = self._read_body()
                    if body is BAD_BODY or not isinstance(body, dict):
                        return self._send(400, {"error": "malformed json body"})
                    res = project.import_excel(body.get("path"), body.get("sheet"), body.get("rowdump"))
                    return self._send(200 if res.get("ok") else 400, res)
                if p == "/api/control_signals":
                    body = self._read_body()
                    if body is BAD_BODY or not isinstance(body, dict):
                        return self._send(400, {"error": "malformed json body"})
                    project.save_control_signals(body)
                    return self._send(200, {"ok": True})
                if p == "/api/build":
                    rb = project.rebuild()
                    if not rb.get("ok"):
                        return self._send(500, {"error": "build failed at " + rb.get("failed", "?"),
                                                "logs": rb.get("logs")})
                    return self._send(200, {"ok": True, "flowgraph": project.flowgraph(),
                                            "regmap": project.regmap(),
                                            "signal_reg_map": project.signal_reg_map(),
                                            "needs_setup": project.needs_setup(), "logs": rb.get("logs")})
                if p == "/api/matching":
                    body = self._read_body()
                    if body is BAD_BODY or not isinstance(body, dict):
                        return self._send(400, {"error": "malformed json body"})
                    # 只接受 str->str 的 alias 与 str 列表的 logic_derived（防脏数据写坏 project.json）
                    alias = {str(k): str(v) for k, v in (body.get("alias") or {}).items()
                             if isinstance(k, str) and isinstance(v, str) and v}
                    logic = [str(x) for x in (body.get("logic_derived") or []) if isinstance(x, str)]
                    project.save_matching(alias, logic)
                    rb = project.rebuild()
                    if not rb.get("ok"):
                        return self._send(500, {"error": "rebuild failed at " + rb.get("failed", "?"),
                                                "logs": rb.get("logs")})
                    return self._send(200, {"ok": True, "regmap": project.regmap(),
                                            "flowgraph": project.flowgraph(),
                                            "signal_reg_map": project.signal_reg_map(),
                                            "matching": project.matching(), "logs": rb.get("logs")})
                return self._send(404, {"error": "unknown api"})
            except Exception as e:
                return self._send(500, {"error": str(e)})

        def do_DELETE(self):
            u = urlparse(self.path)
            p = unquote(u.path)
            m = re.match(r"^/api/mode/([^/]+)$", p)
            if m:
                if not valid_id(m.group(1)):
                    return self._send(404, {"error": "bad id"})
                project.delete_mode(m.group(1))
                return self._send(200, {"ok": True})
            return self._send(404, {"error": "unknown api"})

    return Handler


class _Server(ThreadingHTTPServer):
    # 关掉 SO_REUSEADDR：否则 Windows 上能绑到别的进程正占用的端口（不报错），冲突检测失效
    allow_reuse_address = False


def serve(args):
    project = Project(args.project)
    handler = make_handler(project)
    # 端口被占用就自动顺延（--port 0 = 直接让 OS 挑一个空闲端口）
    candidates = [0] if args.port == 0 else [args.port + i for i in range(30)]
    httpd, tried = None, []
    for pt in candidates:
        try:
            httpd = _Server(("127.0.0.1", pt), handler)
            break
        except OSError:
            tried.append(pt)
    if httpd is None:
        raise SystemExit("端口都被占用（试过 %s–%s）；用 --port <空闲端口> 或 --port 0（自动挑）"
                         % (tried[0], tried[-1]))
    actual = httpd.server_address[1]          # 实际绑定端口（顺延/0 时与 --port 不同）
    url = "http://127.0.0.1:%d/" % actual
    print("Reg_tester GUI  ·  project=%s" % project.root)
    if tried:
        print("  端口 %d 被占用，改用 %d" % (args.port, actual))
    print("  serving %s" % url)
    print("  Ctrl-C 退出")
    if args.open:
        try:
            webbrowser.open(url)
        except Exception:
            pass
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\n bye")
        httpd.shutdown()


# ------------------------------------------------------------------ bundler
def read_text(path):
    with open(path, "r", encoding="utf-8") as f:
        return f.read()


def bundle(args):
    project = Project(args.project)
    html = read_text(os.path.join(WEBAPP, "index.html"))

    # 内联 <link rel=stylesheet href=x.css> 与 <script src=x.js>
    def inline_css(m):
        href = m.group(1)
        try:
            return "<style>\n%s\n</style>" % read_text(os.path.join(WEBAPP, href))
        except Exception:
            return m.group(0)

    def inline_js(m):
        src = m.group(1)
        try:
            code = read_text(os.path.join(WEBAPP, src))
            return "<script>\n%s\n</script>" % code
        except Exception:
            return m.group(0)

    html = re.sub(r'<link[^>]*rel=["\']stylesheet["\'][^>]*href=["\']([^"\']+)["\'][^>]*>',
                  inline_css, html)
    html = re.sub(r'<script[^>]*src=["\']([^"\']+)["\'][^>]*>\s*</script>', inline_js, html)

    payload = {
        "project": project.project_json(),
        "flowgraph": project.flowgraph(),
        "regmap": project.regmap(),
        "layout": project.layout(),
        "modes": {m["id"]: project.read_mode(m["id"]) for m in project.list_modes()},
        "backend": "bundle",
    }
    inject = ("<script>window.__BUNDLE__ = %s;</script>"
              % json.dumps(payload, ensure_ascii=False))
    if "<!--BUNDLE_DATA-->" in html:
        html = html.replace("<!--BUNDLE_DATA-->", inject)
    else:
        html = html.replace("</head>", inject + "\n</head>")

    out = args.out or os.path.join(HERE, "dist", project.project_json().get("name", "bundle") + ".html")
    os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
    with open(out, "w", encoding="utf-8", newline="\n") as f:
        f.write(html)
    kb = len(html.encode("utf-8")) / 1024.0
    print("bundle -> %s  (%.0f KB, 自包含单文件)" % (out, kb))


# ------------------------------------------------------------------ main
def main(argv=None):
    ap = argparse.ArgumentParser(description="Reg_tester GUI 启动器（serve / bundle）")
    ap.add_argument("--serve", action="store_true", help="起 http 服务（读写 project）")
    ap.add_argument("--bundle", action="store_true", help="打包成自包含单 HTML")
    ap.add_argument("--project", required=True, help="projects/<name> 目录")
    ap.add_argument("--port", type=int, default=8765,
                    help="起始端口（默认 8765）；被占用自动顺延，0=让 OS 自动挑空闲端口")
    ap.add_argument("--open", action="store_true", help="serve 时顺便开浏览器")
    ap.add_argument("--out", help="bundle 输出 HTML 路径")
    args = ap.parse_args(argv)
    if args.bundle:
        bundle(args)
    else:
        serve(args)


if __name__ == "__main__":
    main()
