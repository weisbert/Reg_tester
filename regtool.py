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
        if not os.path.isdir(self.root):
            raise SystemExit("project 目录不存在：%s" % self.root)
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
                        "backend": "serve",
                    })
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


def serve(args):
    project = Project(args.project)
    handler = make_handler(project)
    httpd = ThreadingHTTPServer(("127.0.0.1", args.port), handler)
    url = "http://127.0.0.1:%d/" % args.port
    print("Reg_tester GUI  ·  project=%s" % project.root)
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
    ap.add_argument("--port", type=int, default=8765)
    ap.add_argument("--open", action="store_true", help="serve 时顺便开浏览器")
    ap.add_argument("--out", help="bundle 输出 HTML 路径")
    args = ap.parse_args(argv)
    if args.bundle:
        bundle(args)
    else:
        serve(args)


if __name__ == "__main__":
    main()
