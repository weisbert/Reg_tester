#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
summarize_vco_sweep.py — 开环压控扫描簿 → 带汇总页的 Excel

输入：一份仪器/脚本导出的 VCO 开环特性宽表（一行 = 一个测量点）。典型跑法：
    · 某温度下把环路打开，把调谐电压 Vtune 从低扫到高，记 频率/功率/相噪/电流；
    · 再把 Vtune 钉在中间某个值，改电容阵列码（CT，通过一个寄存器位段写入），
      看频段怎么搬；
    · 换温度重来（换温度前一般先闭环锁一次再打开）。

输出：一份新的 .xlsx——
    第 1 页   原始数据（原封不动，就是输入那一页）
    汇总      compliance 版式：逐点指标 Min/Max/Δ（Vtune 扫一张表、CT 扫一张表）
              + 压控特性派生指标（Kvco、线性度、单调性、温漂）
              Spec 两列留空给人填，填完 PASS/FAIL 由 Excel 公式自动出、超规自动标红
    Vtune明细 指标 × Vtune 矩阵（按温度分列），既能翻数也是图表的数据源
    Kvco明细  逐点斜率 ΔF/ΔV（区间中点 × 温度）
    CT明细    指标 × CT 码
    图表      各指标 vs Vtune（每温一条线）、Kvco vs Vtune、vs CT、相噪 vs offset
    锁定点    闭环/锁定那几行单列出来，跟开环曲线对照着看

为什么按「扫描组」而不是按温度汇总
    这类簿里自变量是 Vtune 和 CT，温度只是分组维度。同一个温度下会有两种扫描
    （先扫 Vtune 再扫 CT），横轴根本不是一回事，混在一起统计出来的极值没有意义。
    所以本脚本先按「(温度, 扫的是谁) 变一次就切一组」分组，组内再按横轴排。

依赖：只用 openpyxl。   pip install openpyxl

用法：
    python summarize_vco_sweep.py <扫描簿.xlsx>                 # 出 <原名>_summary.xlsx
    python summarize_vco_sweep.py <扫描簿.xlsx> --dry-run       # 只打印识别结果，不写文件
    python summarize_vco_sweep.py <扫描簿.xlsx> -o 汇总.xlsx
    python summarize_vco_sweep.py <扫描簿.xlsx> --vtune-col Vtune_V --ct-col "REG Value1"
    python summarize_vco_sweep.py <扫描簿.xlsx> --ref-temp 25   # 温漂参考温度

先跑 --dry-run 看四件事对不对：列映射 / 横轴选了哪列 / 分了几组 / 排除了哪些行。
排除的行不会被悄悄丢掉，汇总页底部会逐行列出原因。

⚠ CT 是靠某个寄存器位段扫的，那个地址是 IP：默认只打印列名不打印地址值，
   要看加 --show-addr（输出的 xlsx 里本来就有原始数据，那份别往公开处传）。
"""

import argparse
import os
import re
import sys
from collections import OrderedDict

try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

# report-forge compliance 表的配色（黄表头 / 米色分组行 / 白结果行 / 红超规）
FILL_HEADER = "FFFF00"
FILL_GROUP = "EEECE1"
FILL_RESULT = "FFFFFF"
FILL_SEP = "B8CCE4"
COLOR_FLAG = "FF0000"
COLOR_PASS = "006100"
FILL_FAIL = "FFC7CE"
FILL_PASS = "C6EFCE"

# 表里表示"没测/不适用"的占位符，一律当空值
BLANK_TOKENS = {"", "-", "--", "—", "n/a", "na", "null", "none", "#n/a"}


# ---------------------------------------------------------------- 取值

def is_blank(v):
    if v is None:
        return True
    if isinstance(v, str):
        return v.strip().lower() in BLANK_TOKENS
    return False


def num(v):
    """数值化；取不到数就返回 None。'-' 这类占位符算没测。"""
    if is_blank(v) or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(",", "")
    try:
        return float(s)
    except ValueError:
        return None


def txt(v):
    return "" if v is None else str(v).strip()


def qx(v, nd):
    """横轴取值量化。

    仪器脚本常用累加生成扫描点，于是同一个设定值在表里长这样：
    0.15000000000000002 / 0.39999999999999997 / 0.7000000000000001。
    不量化的话，同一个设定值在不同温度段会变成两个不同的字典键——
    明细页多出半空的行、温漂对不上点，而且一点报错都没有。
    """
    return None if v is None else round(v, nd)


def fmt_num(x, nd=3):
    if x is None:
        return None
    if abs(x - round(x)) < 1e-12:
        return int(round(x))
    return round(x, nd)


def fmt_hz(mhz):
    """offset 频率 MHz -> 好读的标签：0.001->1kHz, 1->1MHz。"""
    if mhz is None:
        return "?"
    if mhz < 1:
        return f"{fmt_num(mhz * 1000.0)}kHz"
    return f"{fmt_num(mhz)}MHz"


# ---------------------------------------------------------------- 列定位

class Columns:
    """按表头名定位列；重名列会报出来。

    这类原厂模板扩列时复制粘贴不改序号很常见（同名列出现两遍）。
    按名字取只会一直拿到第一份，第二份永远读不到且不报错——
    所以这里把重名的列位置全记下来，取第一份并留警告。
    """

    def __init__(self, header):
        self.pos = OrderedDict()
        for i, h in enumerate(header):
            k = txt(h)
            if k:
                self.pos.setdefault(k, []).append(i)
        self.duplicates = {k: v for k, v in self.pos.items() if len(v) > 1}

    def idx(self, name):
        v = self.pos.get(name)
        return v[0] if v else None

    def find(self, *patterns, **kw):
        """按正则找第一个匹配的表头，返回 (名字, 列下标)。"""
        exclude = kw.get("exclude", ())
        for k in self.pos:
            kl = k.lower()
            if any(re.search(p, kl) for p in patterns) and \
               not any(re.search(p, kl) for p in exclude):
                return k, self.pos[k][0]
        return None, None

    def match_all(self, pattern):
        """所有匹配的表头，按列序返回 [(名字, 列下标)]。"""
        out = [(k, v[0]) for k, v in self.pos.items() if re.search(pattern, k.lower())]
        return sorted(out, key=lambda x: x[1])


# ---------------------------------------------------------------- 指标识别

class Item:
    __slots__ = ("cat", "label", "unit", "col", "src")

    def __init__(self, cat, label, unit, col, src):
        self.cat, self.label, self.unit, self.col, self.src = cat, label, unit, col, src


SIMPLE_ITEMS = [
    # (表头, 分类, 单位)
    ("Freq_MHz", "Frequency", "MHz"),
    ("Power_dBm", "Output", "dBm"),
    ("IPN_SSB", "Phase Noise", "dBc"),
    ("IPN_Omit_SSB", "Phase Noise", "dBc"),
    ("Vtune_V", "Tuning", "V"),
    ("Vtemp_V", "Temp Sensor", "V"),
    ("Current_mA", "Current", "mA"),
]
# 成对列：<前缀>Freq<i> 给频点、<前缀>Result<i> 给结果
PAIRED_ITEMS = [
    ("SpotPN", "Phase Noise", "dBc/Hz", "SpotPN@{f}"),
    ("OtherSpur", "Spur", "dBc", "Spur@{f}"),
]


def build_items(cols, rows, skip_cols=()):
    """识别有数据的结果列。整片是空的（占位列）自动丢掉，并记在 dropped 里。

    skip_cols 里的列是横轴（Vtune / CT 码），它是自变量不是指标，不进 items。
    """
    items, dropped = [], []

    def nonempty(ci):
        return sum(1 for r in rows if ci < len(r) and not is_blank(r[ci]))

    for name, cat, unit in SIMPLE_ITEMS:
        ci = cols.idx(name)
        if ci is None or ci in skip_cols:
            continue
        if nonempty(ci) == 0:
            dropped.append((name, "整列没有数据"))
            continue
        items.append(Item(cat, name, unit, ci, name))

    for prefix, cat, unit, tpl in PAIRED_ITEMS:
        i = 1
        while True:
            fc, rc = cols.idx("%sFreq%d" % (prefix, i)), cols.idx("%sResult%d" % (prefix, i))
            if fc is None and rc is None:
                break
            i += 1
            if rc is None or rc in skip_cols:
                continue
            if nonempty(rc) == 0:
                dropped.append(("%sResult%d" % (prefix, i - 1), "整列没有数据"))
                continue
            freqs = set()
            if fc is not None:
                freqs = {num(r[fc]) for r in rows
                         if fc < len(r) and num(r[fc]) is not None}
            f = fmt_hz(sorted(freqs)[0]) if len(freqs) == 1 else (
                "/".join(fmt_hz(x) for x in sorted(freqs)[:3]) or "#%d" % (i - 1))
            items.append(Item(cat, tpl.format(f=f), unit, rc, "%sResult%d" % (prefix, i - 1)))
    return items, dropped


# ---------------------------------------------------------------- 行 / 分组

class Row:
    __slots__ = ("xl", "temp", "mode", "vt", "ct", "group", "vals", "raw")

    def __init__(self, xl, temp, mode, vt, ct, raw):
        self.xl, self.temp, self.mode, self.raw = xl, temp, mode, raw
        self.vt, self.ct = vt, ct
        self.group, self.vals = None, {}


KIND_LABEL = {"vtune": "Vtune扫", "ct": "CT扫", "point": "单点"}


class Group:
    def __init__(self, n, kind, temp):
        self.n, self.kind, self.temp = n, kind, temp
        self.rows = []
        self.tag = ""                     # 同温同类出现多次时的 #2 #3

    @property
    def x_of(self):
        return (lambda r: r.ct) if self.kind == "ct" else (lambda r: r.vt)

    @property
    def x_label(self):
        return "CT 码" if self.kind == "ct" else "Vtune (V)"

    @property
    def x_unit(self):
        return "code" if self.kind == "ct" else "V"

    @property
    def title(self):
        t = fmt_num(self.temp)
        head = "%s℃" % t if t is not None else "?℃"
        return "%s %s%s" % (head, KIND_LABEL.get(self.kind, self.kind), self.tag)

    @property
    def stage(self):
        xs = [x for x in (self.x_of(r) for r in self.rows) if x is not None]
        if not xs:
            return ""
        return "%s→%s  %d 点" % (fmt_num(min(xs)), fmt_num(max(xs)), len(self.rows))


def _changed(a, b):
    if a is None and b is None:
        return False
    if a is None or b is None:
        return True
    return abs(a - b) > 1e-12


def segment(rows):
    """切扫描组。

    规则：(温度, 模式) 一变就切；组内再看「这一步动的是谁」——
    只有 Vtune 动 = Vtune 扫，只有 CT 动 = CT 扫，两个一起动 = 换扫法了，硬切。
    组的类型取组内多数票（单行组没有票，记 point）。
    """
    runs, cur = [], []
    for r in rows:
        if cur and (_changed(cur[-1].temp, r.temp) or cur[-1].mode != r.mode):
            runs.append(cur)
            cur = []
        cur.append(r)
    if cur:
        runs.append(cur)

    groups = []
    for run in runs:
        seg, votes = [run[0]], []
        for prev, row in zip(run, run[1:]):
            dv, dc = _changed(prev.vt, row.vt), _changed(prev.ct, row.ct)
            if dv and dc:                      # 换扫法：硬切
                groups.append((seg, votes))
                seg, votes = [row], []
                continue
            seg.append(row)
            if dc:
                votes.append("ct")
            elif dv:
                votes.append("vtune")
        groups.append((seg, votes))

    out, seen = [], {}
    for seg, votes in groups:
        if not seg:
            continue
        kind = "point"
        if votes:
            kind = "ct" if votes.count("ct") > votes.count("vtune") else "vtune"
        g = Group(len(out) + 1, kind, seg[0].temp)
        g.rows = seg
        key = (g.temp, g.kind)
        seen[key] = seen.get(key, 0) + 1
        if seen[key] > 1:
            g.tag = " #%d" % seen[key]
        for r in seg:
            r.group = g
        out.append(g)
    return out


# ---------------------------------------------------------------- 统计

def group_series(g, item):
    """本组的「横轴 -> 值」去重视图：同一个横轴值有多个点时取最后一个。

    什么时候会同一横轴多点：CT 扫经常先粗扫一遍（0/64/128/192/255）再回头细扫
    （0/1/2…），0 就出现了两次。取后者。

    ★ 汇总统计和明细页**共用这一份**。否则汇总按全部行取极值、明细一格只放得下
    一个值，就会出现「汇总报的极值在明细里查无此值」——报表一旦对不上账就没人敢信。
    """
    xf = g.x_of
    out = OrderedDict()
    for r in g.rows:
        x, v = xf(r), r.vals.get(item.col)
        if x is not None and v is not None:
            out[x] = v
    return OrderedDict(sorted(out.items()))


def _extremes(pairs):
    if not pairs:
        return None
    lo = min(pairs, key=lambda x: x[0])
    hi = max(pairs, key=lambda x: x[0])
    return {"min": lo[0], "min_x": lo[1], "max": hi[0], "max_x": hi[1],
            "delta": hi[0] - lo[0], "n": len(pairs)}


def items_with_data(groups, items):
    """这一批组里真的量到了的指标。

    两种扫法记的东西不一样：Vtune 扫每点都测点相噪/杂散/电流，CT 扫只记
    频率/功率/积分相噪。不按组筛一遍，CT 那几页会多出十几列整片空白。
    """
    return [it for it in items if any(group_series(g, it) for g in groups)]


def stats(g, item):
    return _extremes([(v, x) for x, v in group_series(g, item).items()])


def stats_all(groups, item):
    return _extremes([(v, "%s@%s" % (fmt_num(x), g.title.split()[0]))
                      for g in groups for x, v in group_series(g, item).items()])


def slopes(g, freq_item):
    """逐区间斜率：[(区间中点, ΔF/Δx, 左端, 右端)]。Vtune 扫就是 Kvco。"""
    ser = list(group_series(g, freq_item).items())
    out = []
    for (x0, f0), (x1, f1) in zip(ser, ser[1:]):
        if abs(x1 - x0) < 1e-12:
            continue
        out.append(((x0 + x1) / 2.0, (f1 - f0) / (x1 - x0), x0, x1))
    return out


def monotonic(g, freq_item):
    ser = list(group_series(g, freq_item).values())
    if len(ser) < 2:
        return None
    ups = sum(1 for a, b in zip(ser, ser[1:]) if b > a)
    dns = sum(1 for a, b in zip(ser, ser[1:]) if b < a)
    if ups and dns:
        return "否"
    return "是↑" if ups else ("是↓" if dns else "平")


def drift_vs_ref(g, ref, freq_item):
    """同一横轴值上，本组频率相对参考组的差：(最小, 最大, 重合点数)。"""
    a, b = group_series(g, freq_item), group_series(ref, freq_item)
    shared = [a[x] - b[x] for x in a if x in b]
    if not shared:
        return None
    return min(shared), max(shared), len(shared)


# ---------------------------------------------------------------- 画格子

def _styles():
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    thin = Side(style="thin", color="FF000000")
    return {
        "border": Border(left=thin, right=thin, top=thin, bottom=thin),
        "center": Alignment(horizontal="center", vertical="center", wrap_text=True),
        "left": Alignment(horizontal="left", vertical="center", wrap_text=True),
        "f_head": PatternFill("solid", fgColor=FILL_HEADER, bgColor=FILL_HEADER),
        "f_group": PatternFill("solid", fgColor=FILL_GROUP, bgColor=FILL_GROUP),
        "f_res": PatternFill("solid", fgColor=FILL_RESULT, bgColor=FILL_RESULT),
        "f_sep": PatternFill("solid", fgColor=FILL_SEP, bgColor=FILL_SEP),
        "Font": Font,
    }


def put(ws, r, c, v, st, fill=None, bold=False, color=None, align="center", size=10):
    cell = ws.cell(row=r, column=c)
    cell.value = v
    cell.border = st["border"]
    cell.alignment = st["center"] if align == "center" else st["left"]
    cell.font = st["Font"](bold=bold, color=color, size=size)
    if fill is not None:
        cell.fill = fill
    return cell


def _pass_fail_cf(ws, col_letter, r0, r1, st):
    from openpyxl.formatting.rule import FormulaRule
    from openpyxl.styles import PatternFill
    if r1 < r0:
        return
    rng = "%s%d:%s%d" % (col_letter, r0, col_letter, r1)
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=['%s%d="FAIL"' % (col_letter, r0)],
        fill=PatternFill("solid", fgColor=FILL_FAIL, bgColor=FILL_FAIL),
        font=st["Font"](bold=True, color=COLOR_FLAG), stopIfTrue=False))
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=['%s%d="PASS"' % (col_letter, r0)],
        fill=PatternFill("solid", fgColor=FILL_PASS, bgColor=FILL_PASS),
        font=st["Font"](bold=True, color=COLOR_PASS), stopIfTrue=False))


# ---------------------------------------------------------------- 汇总页

def _range_table(ws, r0, groups, items, st, title):
    """逐点指标表：Category|Item|Unit|Spec Min|Spec Max| |组i Min/Max/Δ| … | |合计| |判定"""
    from openpyxl.utils import get_column_letter as L

    plan = [("cat", "Category", 16), ("item", "Item", 22), ("unit", "Unit", 9),
            ("spec_min", "Min", 10), ("spec_max", "Max", 10), ("sep", "", 2)]
    blocks = [("Spec（自己填）", "手工填写", [3, 4])]
    for g in groups:
        base = len(plan)
        plan += [("min", "Min", 10), ("max", "Max", 10), ("delta", "Δ", 9), ("sep", "", 2)]
        blocks.append((g.title, g.stage, [base, base + 1, base + 2]))
    base = len(plan)
    plan += [("min", "Min", 10), ("min_x", "@点", 12), ("max", "Max", 10),
             ("max_x", "@点", 12), ("delta", "Δ", 9), ("sep", "", 2)]
    blocks.append(("合计", "%d 组并起来" % len(groups),
                   [base, base + 1, base + 2, base + 3, base + 4]))
    col_judge = len(plan)
    plan += [("judge", "判定", 12)]

    for i, (_k, _l, w) in enumerate(plan):
        cur = ws.column_dimensions[L(i + 1)].width
        if not cur or cur < w:
            ws.column_dimensions[L(i + 1)].width = w

    put(ws, r0, 1, title, st, st["f_group"], bold=True, align="left")
    ws.merge_cells(start_row=r0, start_column=1, end_row=r0, end_column=min(14, len(plan)))
    h0 = r0 + 1
    for r in range(h0, h0 + 3):
        for c in range(1, len(plan) + 1):
            put(ws, r, c, None, st, st["f_head"])
    for c0, label in ((0, "Category"), (1, "Item"), (2, "Unit")):
        ws.merge_cells(start_row=h0, start_column=c0 + 1, end_row=h0 + 2, end_column=c0 + 1)
        put(ws, h0, c0 + 1, label, st, st["f_head"], bold=True)
    ws.merge_cells(start_row=h0, start_column=col_judge + 1, end_row=h0 + 2, end_column=col_judge + 1)
    put(ws, h0, col_judge + 1, "判定", st, st["f_head"], bold=True)
    for name, stage, cc in blocks:
        ws.merge_cells(start_row=h0, start_column=cc[0] + 1, end_row=h0, end_column=cc[-1] + 1)
        put(ws, h0, cc[0] + 1, name, st, st["f_head"], bold=True)
        ws.merge_cells(start_row=h0 + 1, start_column=cc[0] + 1, end_row=h0 + 1, end_column=cc[-1] + 1)
        put(ws, h0 + 1, cc[0] + 1, stage, st, st["f_head"], bold=True, size=9)
        for ci in cc:
            put(ws, h0 + 2, ci + 1, plan[ci][1], st, st["f_head"], bold=True, size=9)
    for i, (k, _l, _w) in enumerate(plan):
        if k == "sep":
            for r in range(h0, h0 + 3):
                put(ws, r, i + 1, None, st, st["f_sep"])

    d0 = h0 + 3
    for n, it in enumerate(items):
        r = d0 + n
        for ci in range(len(plan)):
            put(ws, r, ci + 1, None, st,
                st["f_sep"] if plan[ci][0] == "sep" else st["f_res"])
        put(ws, r, 2, it.label, st, st["f_res"], align="left")
        put(ws, r, 3, it.unit, st, st["f_res"])
        for gi, g in enumerate(groups):
            s = stats(g, it)
            cc = blocks[gi + 1][2]
            if s:
                put(ws, r, cc[0] + 1, fmt_num(s["min"]), st, st["f_res"])
                put(ws, r, cc[1] + 1, fmt_num(s["max"]), st, st["f_res"])
                put(ws, r, cc[2] + 1, fmt_num(s["delta"]), st, st["f_res"])
        s = stats_all(groups, it)
        cc = blocks[-1][2]
        if s:
            put(ws, r, cc[0] + 1, fmt_num(s["min"]), st, st["f_res"])
            put(ws, r, cc[1] + 1, s["min_x"], st, st["f_res"], size=9)
            put(ws, r, cc[2] + 1, fmt_num(s["max"]), st, st["f_res"])
            put(ws, r, cc[3] + 1, s["max_x"], st, st["f_res"], size=9)
            put(ws, r, cc[4] + 1, fmt_num(s["delta"]), st, st["f_res"])
        smin, smax = "$D%d" % r, "$E%d" % r
        amin, amax = "%s%d" % (L(cc[0] + 1), r), "%s%d" % (L(cc[2] + 1), r)
        put(ws, r, col_judge + 1,
            '=IF(AND(%s="",%s=""),"",IF(AND(OR(%s="",%s>=%s),OR(%s="",%s<=%s)),"PASS","FAIL"))'
            % (smin, smax, smin, amin, smin, smax, amax, smax),
            st, st["f_res"], bold=True)

    i = 0
    while i < len(items):
        j = i
        while j + 1 < len(items) and items[j + 1].cat == items[i].cat:
            j += 1
        if j > i:
            ws.merge_cells(start_row=d0 + i, start_column=1, end_row=d0 + j, end_column=1)
        put(ws, d0 + i, 1, items[i].cat, st, st["f_res"], bold=True)
        i = j + 1

    _pass_fail_cf(ws, L(col_judge + 1), d0, d0 + len(items) - 1, st)
    return d0 + len(items)


def _single_table(ws, r0, groups, derived, st, title):
    """派生指标表：每组一个值。Category|Item|Unit|Spec Min|Spec Max|组1..组N|判定"""
    from openpyxl.utils import get_column_letter as L

    ncol = 5 + len(groups) + 1
    col_judge = ncol
    put(ws, r0, 1, title, st, st["f_group"], bold=True, align="left")
    ws.merge_cells(start_row=r0, start_column=1, end_row=r0, end_column=min(14, ncol))
    h0 = r0 + 1
    for c in range(1, ncol + 1):
        put(ws, h0, c, None, st, st["f_head"])
        put(ws, h0 + 1, c, None, st, st["f_head"])
    for c0, label in ((1, "Category"), (2, "Item"), (3, "Unit")):
        ws.merge_cells(start_row=h0, start_column=c0, end_row=h0 + 1, end_column=c0)
        put(ws, h0, c0, label, st, st["f_head"], bold=True)
    ws.merge_cells(start_row=h0, start_column=4, end_row=h0, end_column=5)
    put(ws, h0, 4, "Spec（自己填）", st, st["f_head"], bold=True)
    put(ws, h0 + 1, 4, "Min", st, st["f_head"], bold=True, size=9)
    put(ws, h0 + 1, 5, "Max", st, st["f_head"], bold=True, size=9)
    for j, g in enumerate(groups):
        put(ws, h0, 6 + j, g.title, st, st["f_head"], bold=True, size=9)
        put(ws, h0 + 1, 6 + j, g.stage, st, st["f_head"], size=8)
    ws.merge_cells(start_row=h0, start_column=col_judge, end_row=h0 + 1, end_column=col_judge)
    put(ws, h0, col_judge, "判定", st, st["f_head"], bold=True)

    d0 = h0 + 2
    for n, (cat, label, unit, values) in enumerate(derived):
        r = d0 + n
        for c in range(1, ncol + 1):
            put(ws, r, c, None, st, st["f_res"])
        put(ws, r, 2, label, st, st["f_res"], align="left")
        put(ws, r, 3, unit, st, st["f_res"])
        numeric = True
        for j, g in enumerate(groups):
            v = values.get(id(g))
            if isinstance(v, str):
                numeric = False
            put(ws, r, 6 + j, v if isinstance(v, str) else fmt_num(v), st, st["f_res"])
        if numeric and groups:
            rng = "%s%d:%s%d" % (L(6), r, L(5 + len(groups)), r)
            put(ws, r, col_judge,
                '=IF(COUNT(%s)=0,"",IF(AND($D%d="",$E%d=""),"",'
                'IF(AND(OR($D%d="",MIN(%s)>=$D%d),OR($E%d="",MAX(%s)<=$E%d)),"PASS","FAIL")))'
                % (rng, r, r, r, rng, r, r, rng, r),
                st, st["f_res"], bold=True)
    i = 0                                     # Category 列：同类连续行纵向合并
    while i < len(derived):
        j = i
        while j + 1 < len(derived) and derived[j + 1][0] == derived[i][0]:
            j += 1
        if j > i:
            ws.merge_cells(start_row=d0 + i, start_column=1, end_row=d0 + j, end_column=1)
        put(ws, d0 + i, 1, derived[i][0], st, st["f_res"], bold=True)
        i = j + 1
    _pass_fail_cf(ws, L(col_judge), d0, d0 + len(derived) - 1, st)
    return d0 + len(derived)


def build_derived(groups, freq_item, ref_group):
    """压控/CT 特性派生指标。每个指标一行，每组一个值。"""
    rows, vals = [], {}

    def add(cat, label, unit, fn):
        v = {}
        for g in groups:
            try:
                v[id(g)] = fn(g)
            except Exception:
                v[id(g)] = None
        if any(x is not None for x in v.values()):
            rows.append((cat, label, unit, v))

    def endpoints(g):
        ser = list(group_series(g, freq_item).items())
        return (ser[0], ser[-1]) if len(ser) >= 2 else (None, None)

    def span(g):
        lo, hi = endpoints(g)
        return None if lo is None else hi[1] - lo[1]

    def avg_slope(g):
        lo, hi = endpoints(g)
        if lo is None or abs(hi[0] - lo[0]) < 1e-12:
            return None
        return (hi[1] - lo[1]) / (hi[0] - lo[0])

    kind = groups[0].kind if groups else "vtune"
    is_ct = kind == "ct"
    su = "MHz/code" if is_ct else "MHz/V"
    sname = "ΔF/ΔCT" if is_ct else "Kvco"
    cat = "CT 特性" if is_ct else "压控特性"

    add(cat, "起点 F（横轴最小处）", "MHz",
        lambda g: (endpoints(g)[0] or (None, None))[1])
    add(cat, "终点 F（横轴最大处）", "MHz",
        lambda g: (endpoints(g)[1] or (None, None))[1])
    add(cat, "覆盖范围 |ΔF|", "MHz", lambda g: abs(span(g)) if span(g) is not None else None)
    add(cat, "%s 平均（端到端）" % sname, su, avg_slope)
    add(cat, "%s 最小（逐点）" % sname, su,
        lambda g: min((s[1] for s in slopes(g, freq_item)), default=None))
    add(cat, "%s 最大（逐点）" % sname, su,
        lambda g: max((s[1] for s in slopes(g, freq_item)), default=None))

    def ratio(g):
        ss = [abs(s[1]) for s in slopes(g, freq_item) if abs(s[1]) > 1e-12]
        return max(ss) / min(ss) if ss else None

    add(cat, "%s 最大/最小（线性度）" % sname, "-", ratio)
    add(cat, "单调", "-", lambda g: monotonic(g, freq_item))
    add(cat, "有效点数", "点", lambda g: len(group_series(g, freq_item)) or None)

    if ref_group is not None and len(groups) > 1:
        rt = fmt_num(ref_group.temp)
        add("温漂", "ΔF vs %s℃ 最小（同横轴点）" % rt, "MHz",
            lambda g: (drift_vs_ref(g, ref_group, freq_item) or (None,))[0])
        add("温漂", "ΔF vs %s℃ 最大（同横轴点）" % rt, "MHz",
            lambda g: (drift_vs_ref(g, ref_group, freq_item) or (None, None))[1])
        add("温漂", "重合横轴点数", "点",
            lambda g: (drift_vs_ref(g, ref_group, freq_item) or (None, None, None))[2])
    return rows


def write_summary(wb, by_kind, items, meta, st):
    """每种扫法一张汇总页。

    Vtune 扫和 CT 扫的组数不一样、列宽也不一样，硬塞进一页会互相把列撑歪，
    所以分页：第一种扫法占「汇总」，其余的另起「汇总-CT扫」这样的页。
    页底的说明/排除行只写在第一页。
    """
    first_ws = None
    for kind, groups in by_kind:
        if not groups:
            continue
        name = "汇总" if first_ws is None else "汇总-%s" % KIND_LABEL.get(kind, kind)
        ws = wb.create_sheet(name)
        if first_ws is None:
            first_ws = ws
        r = _range_table(ws, 1, groups, items_with_data(groups, items), st,
                         "逐点指标 · %s（横轴 %s）" % (KIND_LABEL.get(kind, kind),
                                                  groups[0].x_label)) + 2
        freq_item = meta["freq_item"]
        if freq_item is not None:
            derived = build_derived(groups, freq_item, meta["ref_by_kind"].get(kind))
            if derived:
                _single_table(ws, r, groups, derived, st,
                              "派生指标 · %s" % KIND_LABEL.get(kind, kind))
        ws.freeze_panes = "D5"

    if first_ws is None:
        return None
    ws = first_ws
    r = ws.max_row + 2
    put(ws, r, 1, "怎么用", st, st["f_group"], bold=True, align="left")
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
    for line in [
        "① Spec 的 Min / Max 两列自己填（只关心单边就只填一边，另一边留空）。",
        "② 填完「判定」列自动出 PASS / FAIL 并上色，不用重跑脚本。",
        "③ 逐点指标表判定用的是「合计」的 Min / Max（各组所有横轴点里的极值），"
        "@点 列写明它出在哪一组的哪个横轴值上，能在明细页原样查到。",
        "④ 派生指标表判定用的是各组那一行的 MIN / MAX。",
        "⑤ 同一组同一个横轴值若测了两次（例如 CT 先粗扫再回头细扫，0 出现两遍），取后者。",
        "⑥ 分组规则：%s" % meta["why_groups"],
    ]:
        r += 1
        put(ws, r, 1, line, st, None, align="left")
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=12)

    r += 2
    put(ws, r, 1, "没算进汇总的行（逐行列出，不做静默丢弃）", st, st["f_group"], bold=True, align="left")
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    r += 1
    put(ws, r, 1, "原表行号", st, st["f_head"], bold=True)
    put(ws, r, 2, "原因", st, st["f_head"], bold=True)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
    for xl, why in meta["excluded"]:
        r += 1
        put(ws, r, 1, xl, st)
        put(ws, r, 2, why, st, align="left")
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
    if meta["warnings"]:
        r += 2
        put(ws, r, 1, "⚠ 探查告警", st, st["f_group"], bold=True, align="left")
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
        for w in meta["warnings"]:
            r += 1
            put(ws, r, 1, w, st, None, align="left", color=COLOR_FLAG)
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=12)
    return ws


# ---------------------------------------------------------------- 明细页

def write_detail(wb, name, groups, items, st):
    """每个指标一块：行=横轴取值（各组并集，升序），列=组。图表就吃这个。"""
    from openpyxl.utils import get_column_letter as L
    ws = wb.create_sheet(name)
    ws.column_dimensions["A"].width = 12
    for i in range(len(groups)):
        ws.column_dimensions[L(2 + i)].width = 14

    xs = sorted({g.x_of(r) for g in groups for r in g.rows if g.x_of(r) is not None})
    blocks, r = [], 1
    for it in items:
        put(ws, r, 1, "%s  [%s]" % (it.label, it.unit), st, st["f_group"], bold=True, align="left")
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=1 + len(groups))
        r += 1
        put(ws, r, 1, groups[0].x_label, st, st["f_head"], bold=True)
        for i, g in enumerate(groups):
            put(ws, r, 2 + i, g.title, st, st["f_head"], bold=True, size=9)
        head, r = r, r + 1
        first = r
        series = [group_series(g, it) for g in groups]
        for x in xs:
            put(ws, r, 1, fmt_num(x), st, st["f_res"])
            for i in range(len(groups)):
                put(ws, r, 2 + i, fmt_num(series[i].get(x)), st, st["f_res"])
            r += 1
        blocks.append((it, head, first, r - 1))
        r += 1
    ws.freeze_panes = "B1"
    return ws, blocks


def write_slope(wb, name, groups, freq_item, st, unit, xlabel):
    """逐点斜率页：行=区间中点，列=组。Vtune 扫的就是 Kvco vs Vtune。"""
    from openpyxl.utils import get_column_letter as L
    ws = wb.create_sheet(name)
    ws.column_dimensions["A"].width = 14
    for i in range(len(groups)):
        ws.column_dimensions[L(2 + i)].width = 14

    per = [dict(((round(m, 9), v) for m, v, _a, _b in slopes(g, freq_item))) for g in groups]
    mids = sorted({m for d in per for m in d})
    put(ws, 1, 1, "逐区间斜率 Δ%s = (F右-F左)/(x右-x左)，横坐标取区间中点  [%s]"
        % (freq_item.label, unit), st, st["f_group"], bold=True, align="left")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=1 + len(groups))
    put(ws, 2, 1, "%s 中点" % xlabel, st, st["f_head"], bold=True)
    for i, g in enumerate(groups):
        put(ws, 2, 2 + i, g.title, st, st["f_head"], bold=True, size=9)
    r = 3
    for m in mids:
        put(ws, r, 1, fmt_num(m), st, st["f_res"])
        for i in range(len(groups)):
            put(ws, r, 2 + i, fmt_num(per[i].get(m)), st, st["f_res"])
        r += 1
    ws.freeze_panes = "B2"
    return ws, (2, 3, r - 1)


# ---------------------------------------------------------------- 图表

def _scatter(title, xtitle, ytitle, logx=False):
    from openpyxl.chart import ScatterChart
    ch = ScatterChart()
    ch.title = title
    ch.style = 13
    ch.x_axis.title = xtitle
    ch.y_axis.title = ytitle
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ch.height, ch.width = 7.5, 13
    ch.dispBlanksAs = "gap"
    if logx:
        ch.x_axis.scaling.logBase = 10
    return ch


def _add_series(ch, ws, n_series, head, first, last):
    from openpyxl.chart import Reference, Series
    from openpyxl.chart.marker import Marker
    xref = Reference(ws, min_col=1, min_row=first, max_row=last)
    for i in range(n_series):
        yref = Reference(ws, min_col=2 + i, min_row=head, max_row=last)
        s = Series(yref, xref, title_from_data=True)
        s.marker = Marker(symbol="circle", size=5)
        ch.series.append(s)


class ChartGrid:
    """图一张挨一张往下摆：两列（A / I），每行带 16 行高。

    别用 ws.max_row 当锚点——图不算单元格内容，max_row 一直是 1，
    后加的图会全叠在第一张上面。
    """

    def __init__(self, ws):
        self.ws, self.n = ws, 0

    def add(self, ch):
        self.ws.add_chart(ch, "%s%d" % ("A" if self.n % 2 == 0 else "I",
                                        3 + (self.n // 2) * 16))
        self.n += 1


# 每个 offset 的点相噪/杂散单独出图 = 十几张一样的图，默认不画；
# 它们在明细页翻得到，趋势看「相噪 vs offset」那张。--all-charts 全画。
MINOR_CHART = re.compile(r"^(SpotPN@|Spur@)")


def write_charts(wb, panels, st, all_charts=False):
    """panels: [(sheet, blocks, groups, xlabel)]，每块一张图。"""
    ws = wb.create_sheet("图表")
    put(ws, 1, 1, "每张图：横轴=扫描变量，每组一条线（不同温度/不同扫法分开画）。",
        st, st["f_group"], bold=True, align="left")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=10)
    grid = ChartGrid(ws)
    for ws_src, blocks, groups, xlabel in panels:
        for it, head, first, last in blocks:
            if not all_charts and MINOR_CHART.match(it.label):
                continue
            ch = _scatter("%s vs %s" % (it.label, xlabel), xlabel, it.unit)
            _add_series(ch, ws_src, len(groups), head, first, last)
            grid.add(ch)
    return ws, grid


def write_slope_chart(grid, ws_slope, groups, anchor, title, xlabel, unit):
    head, first, last = anchor[0], anchor[1], anchor[2]
    if last < first:
        return
    ch = _scatter(title, xlabel, unit)
    _add_series(ch, ws_slope, len(groups), head, first, last)
    grid.add(ch)


def write_pn_chart(wb, grid, groups, pn_items, st):
    """相噪 vs offset：数据另开一小块（行=offset，列=各组的代表点），再挂散点图。

    每组取横轴的中间那个点当代表——相噪本身随 Vtune 变，取哪个点必须写清楚，
    否则图上三条线不知道是在什么条件下量的。
    """
    from openpyxl.chart import Reference, Series
    from openpyxl.chart.marker import Marker
    if not pn_items or not groups:
        return
    ws = wb.create_sheet("PN曲线数据")
    picks = []
    for g in groups:
        xs = sorted({g.x_of(r) for r in g.rows if g.x_of(r) is not None})
        if not xs:
            continue
        xm = xs[len(xs) // 2]
        row = None
        for r in g.rows:
            if g.x_of(r) is not None and abs(g.x_of(r) - xm) < 1e-12:
                row = r
        if row is not None:
            picks.append((g, xm, row))
    if not picks:
        return

    put(ws, 1, 1, "相噪 vs offset。每组取横轴中间那个点当代表（下面写明是哪个点）。",
        st, st["f_group"], bold=True, align="left")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=1 + len(picks))
    put(ws, 2, 1, "offset_MHz", st, st["f_head"], bold=True)
    for j, (g, xm, _r) in enumerate(picks):
        put(ws, 2, 2 + j, "%s @%s=%s" % (g.title, g.x_unit, fmt_num(xm)),
            st, st["f_head"], bold=True, size=9)
    for i, it in enumerate(pn_items):
        m = re.search(r"@([\d.]+)(k|M)Hz", it.label)
        off = float(m.group(1)) / (1000.0 if m and m.group(2) == "k" else 1.0) if m else None
        put(ws, 3 + i, 1, off, st, st["f_res"])
        for j, (_g, _xm, row) in enumerate(picks):
            put(ws, 3 + i, 2 + j, fmt_num(row.vals.get(it.col)), st, st["f_res"])
    ws.column_dimensions["A"].width = 12

    ch = _scatter("相噪 vs offset", "offset (MHz)", "dBc/Hz", logx=True)
    xref = Reference(ws, min_col=1, min_row=3, max_row=2 + len(pn_items))
    for j in range(len(picks)):
        yref = Reference(ws, min_col=2 + j, min_row=2, max_row=2 + len(pn_items))
        s = Series(yref, xref, title_from_data=True)
        s.marker = Marker(symbol="circle", size=5)
        ch.series.append(s)
    grid.add(ch)


# ---------------------------------------------------------------- 锁定点页

def write_locked(wb, locked, items, st, temp_name):
    """扫描序列之外、但确实量到东西的行，单独列出来。

    这类簿里每换一次温度都有一小段 闭环 → 锁定 → 开环 的切换动作，其中「锁定」
    那行是带完整测量的：它就是闭环最后落在压控曲线的哪个位置，跟开环曲线对照着看
    才有意义。还有自检 / 换模式那几行也可能带部分结果——一并列在这里，
    免得「排除了 N 行」看着像凭空丢了数据。
    """
    ws = wb.create_sheet("非扫描行")
    put(ws, 1, 1, "扫描序列之外但带测量值的行（不进扫描统计；「锁定」行可跟开环曲线对照）",
        st, st["f_group"], bold=True, align="left")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=5 + len(items))
    heads = ["原表行号", "类型", "Mode", temp_name or "温度", "Vtune"] + \
            ["%s\n[%s]" % (it.label, it.unit) for it in items]
    for j, h in enumerate(heads):
        put(ws, 2, 1 + j, h, st, st["f_head"], bold=True, size=9)
    for i, (r, kind) in enumerate(locked):
        put(ws, 3 + i, 1, r.xl, st, st["f_res"])
        put(ws, 3 + i, 2, kind, st, st["f_res"])
        put(ws, 3 + i, 3, r.mode, st, st["f_res"], align="left")
        put(ws, 3 + i, 4, fmt_num(r.temp), st, st["f_res"])
        put(ws, 3 + i, 5, fmt_num(r.vt), st, st["f_res"])
        for j, it in enumerate(items):
            put(ws, 3 + i, 6 + j, fmt_num(r.vals.get(it.col)), st, st["f_res"])
    ws.column_dimensions["C"].width = 22
    ws.freeze_panes = "C3"
    return ws


# ---------------------------------------------------------------- 横轴列识别

VTUNE_CANDIDATES = [r"^dc_start_v$", r"vtune.*set", r"^vtune_v$", r"^vtune$"]


def pick_vtune_col(cols, data):
    """挑 Vtune 横轴列：优先设定值列（DC_Start_V 之类），没有才用读回的 Vtune_V。"""
    for pat in VTUNE_CANDIDATES:
        for name, ci in cols.match_all(pat):
            vals = {num(r[ci]) for r in data if ci < len(r) and num(r[ci]) is not None}
            if len(vals) >= 3:
                return name, ci, "有 %d 个不同取值" % len(vals)
    return None, None, ""


def pick_ct_col(cols, data):
    """挑 CT 码列：REG Value<n> 里那个"只在一部分行有值、且值在变"的。

    ★ 优先"只覆盖一部分行"的列：每行都写的多半是常驻配置（整份簿都要写的那几组
    寄存器），扫描列一般只在扫描那几行才填。选错了会把分组切碎——所以 --dry-run
    一定要看一眼这里选中的是哪列，不对就 --ct-col 指定。
    """
    n_rows = sum(1 for r in data if any(not is_blank(v) for v in r))
    best = None
    for name, ci in cols.match_all(r"^reg[ _]?value\d+$"):
        good = [num(r[ci]) for r in data if ci < len(r) and num(r[ci]) is not None]
        distinct = set(good)
        if len(distinct) < 3:
            continue
        partial = 1 if len(good) < n_rows else 0
        cand = (partial, len(distinct), -len(good), name, ci)
        if best is None or cand > best:
            best = cand
    if best is None:
        return None, None, ""
    return best[3], best[4], "有 %d 个不同取值、%d 行非空%s" % (
        best[1], -best[2], "" if best[0] else "（整表都有值，确认一下是不是常驻配置列）")


# ---------------------------------------------------------------- main

def main():
    ap = argparse.ArgumentParser(description="开环压控扫描簿 → 带汇总页的 Excel")
    ap.add_argument("path", help="扫描簿 .xlsx")
    ap.add_argument("-o", "--out", default=None, help="输出路径（默认 <原名>_summary.xlsx）")
    ap.add_argument("--sheet", default=None, help="数据所在 sheet（默认第一个）")
    ap.add_argument("--header-row", type=int, default=1, help="表头行号（1 基，默认 1）")
    ap.add_argument("--mode-col", default="Mode", help="模式列（默认 Mode）")
    ap.add_argument("--lock-pattern", default=r"lock$|close.?loop",
                    help="模式列匹配这个正则的行 = 闭环/锁定行，不进开环统计"
                         "（默认 lock$|close.?loop）")
    ap.add_argument("--sweep-mode", default=None,
                    help="只统计该模式的行（默认取非锁定行里出现最多的那个值）。"
                         "挡的是夹在扫描序列里的旁路/自检行——它们也带部分结果值，"
                         "不挡就会被算进相邻那一组，把曲线和极值都带偏")
    ap.add_argument("--temp-col", default=None, help="温度列（默认自动找含 Temperature 的列）")
    ap.add_argument("--vtune-col", default=None, help="Vtune 横轴列（默认自动挑）")
    ap.add_argument("--ct-col", default=None,
                    help="CT 码横轴列（默认自动挑；写 none 关掉 CT 扫识别）")
    ap.add_argument("--keep-test-item", default=None,
                    help="只保留 Test Item 等于该值的行（默认取出现最多的那个值）")
    ap.add_argument("--ref-temp", type=float, default=25.0,
                    help="温漂参考温度（默认 25）")
    ap.add_argument("--x-round", type=int, default=6,
                    help="横轴取值四舍五入到几位小数（默认 6）。扫描点常是累加出来的，"
                         "表里会是 0.39999999999999997 这种，不量化就跟别的段对不上点")
    ap.add_argument("--show-addr", action="store_true",
                    help="打印 CT 扫用的寄存器地址（默认不打印，地址是 IP）")
    ap.add_argument("--all-charts", action="store_true",
                    help="每个指标都出图（默认跳过逐 offset 的点相噪/杂散那十几张）")
    ap.add_argument("--dry-run", action="store_true", help="只打印识别结果，不写文件")
    args = ap.parse_args()

    if not os.path.isfile(args.path):
        sys.exit("找不到文件: %s" % args.path)
    try:
        import openpyxl
    except ImportError:
        sys.exit("缺少 openpyxl，请先: pip install openpyxl")

    # 读两份：一份取缓存值用来算，一份原封不动用来存。
    # 只用 data_only=True 那份去存的话，原表里若有公式会被替换成计算结果——
    # 「第 1 页保留原始 excel」就不成立了。
    wb_val = openpyxl.load_workbook(args.path, data_only=True)
    ws_val = wb_val[args.sheet] if args.sheet else wb_val[wb_val.sheetnames[0]]
    wb = openpyxl.load_workbook(args.path, data_only=False)
    ws = wb[ws_val.title]
    all_rows = [list(r) for r in ws_val.iter_rows(values_only=True)]
    if len(all_rows) < args.header_row + 1:
        sys.exit("表里没有数据行")
    header = all_rows[args.header_row - 1]
    data = all_rows[args.header_row:]

    cols = Columns(header)
    warnings = []
    if cols.duplicates:
        from openpyxl.utils import get_column_letter as gl
        for k, v in cols.duplicates.items():
            warnings.append("重复列名 %s：出现在 %s，按名字只取到第一个（%s）"
                            % (k, ", ".join(gl(i + 1) for i in v), gl(v[0] + 1)))

    if args.temp_col:
        tname, tcol = args.temp_col, cols.idx(args.temp_col)
    else:
        tname, tcol = cols.find(r"temperature", r"^temp")
    if tcol is None:
        warnings.append("找不到温度列，全部行按同一个温度处理（--temp-col 可指定）")
    mode_i = cols.idx(args.mode_col)
    if mode_i is None:
        warnings.append("没有 %s 列，无法分辨闭环/开环行" % args.mode_col)
    ti_col = cols.idx("Test Item")

    # 横轴列
    if args.vtune_col:
        vt_name, vt_col, vt_why = args.vtune_col, cols.idx(args.vtune_col), "命令行指定"
        if vt_col is None:
            sys.exit("找不到 --vtune-col 指定的列: %s" % args.vtune_col)
    else:
        vt_name, vt_col, vt_why = pick_vtune_col(cols, data)
    if vt_col is None:
        sys.exit("没找到 Vtune 横轴列，用 --vtune-col 指定")

    ct_name = ct_col = None
    ct_why = ""
    if txt(args.ct_col).lower() != "none":
        if args.ct_col:
            ct_name, ct_col, ct_why = args.ct_col, cols.idx(args.ct_col), "命令行指定"
            if ct_col is None:
                sys.exit("找不到 --ct-col 指定的列: %s" % args.ct_col)
        else:
            ct_name, ct_col, ct_why = pick_ct_col(cols, data)
    ct_addr = None
    if ct_name:
        m = re.search(r"(\d+)$", ct_name)
        if m:
            ai = cols.idx("REG ADDR%s" % m.group(1))
            if ai is not None:
                for r in data:
                    if ai < len(r) and not is_blank(r[ai]):
                        ct_addr = txt(r[ai])
                        break

    lock_re = re.compile(args.lock_pattern, re.I)

    keep_ti = args.keep_test_item
    if ti_col is not None and keep_ti is None:
        cnt = {}
        for r in data:
            v = txt(r[ti_col]) if ti_col < len(r) else ""
            if v:
                cnt[v] = cnt.get(v, 0) + 1
        if cnt:
            keep_ti = max(cnt, key=lambda k: cnt[k])

    sweep_mode = args.sweep_mode
    if mode_i is not None and sweep_mode is None:
        cnt = {}
        for r in data:
            v = txt(r[mode_i]) if mode_i < len(r) else ""
            if v and not lock_re.search(v):
                cnt[v] = cnt.get(v, 0) + 1
        if cnt:
            sweep_mode = max(cnt, key=lambda k: cnt[k])

    excluded, rows, locked, others = [], [], [], []
    for n, raw in enumerate(data):
        xl = args.header_row + 1 + n
        if all(is_blank(v) for v in raw):
            continue
        mode = txt(raw[mode_i]) if mode_i is not None and mode_i < len(raw) else ""
        temp = num(raw[tcol]) if tcol is not None and tcol < len(raw) else None
        vt = qx(num(raw[vt_col]) if vt_col < len(raw) else None, args.x_round)
        ct = qx(num(raw[ct_col]) if ct_col is not None and ct_col < len(raw) else None,
                args.x_round)
        r = Row(xl, temp, mode, vt, ct, raw)
        if mode and lock_re.search(mode):
            locked.append(r)
            excluded.append((xl, "%s = %r，闭环/锁定行（列在「非扫描行」页）" % (args.mode_col, mode)))
            continue
        if ti_col is not None and keep_ti is not None:
            v = txt(raw[ti_col]) if ti_col < len(raw) else ""
            if v != keep_ti:
                excluded.append((xl, "Test Item = %r，不是主测试项 %r" % (v, keep_ti)))
                continue
        if sweep_mode is not None and mode != sweep_mode:
            others.append(r)
            excluded.append((xl, "%s = %r，不是扫描模式 %r" % (args.mode_col, mode, sweep_mode)))
            continue
        rows.append(r)

    skip = {vt_col} | ({ct_col} if ct_col is not None else set())
    items, dropped = build_items(cols, [r.raw for r in rows + locked + others],
                                 skip_cols=skip)
    if not items:
        sys.exit("没识别出任何有数据的结果列")
    for r in rows + locked + others:
        for it in items:
            r.vals[it.col] = num(r.raw[it.col]) if it.col < len(r.raw) else None

    # 扫描序列之外但确实量到东西的行：单列一页，别让「排除了 N 行」看着像丢了数据
    extra = [(r, "锁定") for r in locked if any(v is not None for v in r.vals.values())]
    extra += [(r, "其它") for r in others if any(v is not None for v in r.vals.values())]
    extra.sort(key=lambda x: x[0].xl)

    keep = []
    for r in rows:
        if all(v is None for v in r.vals.values()):
            excluded.append((r.xl, "所有结果列都是空的（配置/开关行，不是测量点）"))
        else:
            keep.append(r)
    rows = keep
    if not rows:
        sys.exit("过滤完没有测量点了，检查 --sweep-mode / --keep-test-item")

    groups = segment(rows)
    # 横轴值大面积重复 = 组里还藏着另一个在动的变量（最常见：CT 列没识别出来，
    # 于是 CT 扫那几行被并进 Vtune 扫，去重后只剩最后一个，把曲线上那个点悄悄换掉）
    for g in groups:
        xs = [g.x_of(r) for r in g.rows if g.x_of(r) is not None]
        dup = len(xs) - len(set(xs))
        # 少量重复是正常的（先粗扫一遍再回头细扫，几个点会重复量）；
        # 大面积重复才是「组里还藏着另一个在动的变量」。
        if dup >= 2 and dup > 0.2 * len(xs):
            warnings.append("组「%s」有 %d 个点但横轴只有 %d 个不同取值——"
                            "是不是还有一个没识别出来的扫描变量？(--ct-col 指定)"
                            % (g.title, len(xs), len(set(xs))))

    freq_item = next((it for it in items if it.src == "Freq_MHz"), None)
    if freq_item is None:
        freq_item = next((it for it in items if it.cat == "Frequency"), None)
        if freq_item is None:
            warnings.append("没有频率列，Kvco / 覆盖范围这些派生指标出不来")

    by_kind = []
    for kind in ("vtune", "ct", "point"):
        gs = [g for g in groups if g.kind == kind]
        if gs:
            by_kind.append((kind, gs))
    ref_by_kind = {}
    for kind, gs in by_kind:
        withtemp = [g for g in gs if g.temp is not None]
        if withtemp:
            ref_by_kind[kind] = min(withtemp, key=lambda g: abs(g.temp - args.ref_temp))

    excluded.sort(key=lambda x: x[0])
    meta = {
        "excluded": excluded, "warnings": warnings, "freq_item": freq_item,
        "ref_by_kind": ref_by_kind,
        "why_groups": ("(温度, 模式) 一变就切；组内再看这一步动的是 Vtune 还是 CT，"
                       "两个一起动就是换扫法了，也切。横轴不是一回事的点混在一起统计没有意义。"),
    }

    # ---- 打印识别结果 ----
    # ws.max_row/max_column 是「声明尺寸」——模板预设过格式的空区域也算进去，
    # 常见到 10000 行 × 205 列这种数字。真正有内容的行才是要报的。
    n_filled = sum(1 for r in data if any(not is_blank(v) for v in r))
    print("源文件 : %s   sheet=%s" % (os.path.basename(args.path), ws.title))
    print("规模   : 有内容的数据行 %d（表头第 %d 行，表头列 %d 个；"
          "工作表声明尺寸 %d 行 × %d 列，其余是模板预留的空区域）"
          % (n_filled, args.header_row, sum(1 for h in header if txt(h)),
             ws.max_row, ws.max_column))
    print("温度列 : %s" % (tname or "(没有)"))
    print("Vtune  : %s   %s" % (vt_name, vt_why))
    if ct_name:
        addr = ("  地址=%s" % ct_addr) if (ct_addr and args.show_addr) else \
               ("  (地址见原表，--show-addr 才打印)" if ct_addr else "")
        print("CT 码  : %s   %s%s" % (ct_name, ct_why, addr))
    else:
        print("CT 码  : 没识别到（--ct-col 可指定）")
    if keep_ti is not None:
        print("主测试项: %r（其余行排除）" % keep_ti)
    if sweep_mode is not None:
        print("扫描模式: %r（闭环/锁定行另列；其余行排除）" % sweep_mode)
    print("识别指标 %d 个:" % len(items))
    for it in items:
        print("    [%-12s] %-18s %-7s <- 列 %s" % (it.cat, it.label, it.unit, it.src))
    if dropped:
        print("  跳过的空列 %d 个: %s" % (len(dropped), ", ".join(k for k, _ in dropped)))
    print("分组 %d 组:" % len(groups))
    for g in groups:
        print("    %-16s %-22s 行 %d~%d" % (g.title, g.stage, g.rows[0].xl, g.rows[-1].xl))
    if extra:
        print("非扫描行里带测量值的 %d 行（列在「非扫描行」页）: %s"
              % (len(extra), ", ".join("%d/%s" % (r.xl, k) for r, k in extra[:12])))
    if excluded:
        print("排除 %d 行:" % len(excluded))
        for xl, why in excluded:
            print("    行%d: %s" % (xl, why))
    for w in warnings:
        print("  ⚠ %s" % w)
    if args.dry_run:
        print("\n--dry-run：没有写文件。")
        return

    # ---- 写出 ----
    st = _styles()
    write_summary(wb, by_kind, items, meta, st)

    panels, slope_jobs = [], []
    for kind, gs in by_kind:
        if kind == "point":
            continue
        nm = "Vtune明细" if kind == "vtune" else "CT明细"
        ws_d, blocks = write_detail(wb, nm, gs, items_with_data(gs, items), st)
        panels.append((ws_d, blocks, gs, gs[0].x_label))
        if freq_item is not None and any(len(group_series(g, freq_item)) >= 2 for g in gs):
            sn = "Kvco明细" if kind == "vtune" else "CT斜率明细"
            unit = "MHz/V" if kind == "vtune" else "MHz/code"
            ws_s, anchor = write_slope(wb, sn, gs, freq_item, st, unit, gs[0].x_label)
            slope_jobs.append((ws_s, gs, anchor,
                               "Kvco vs Vtune" if kind == "vtune" else "ΔF/ΔCT vs CT",
                               gs[0].x_label, unit))

    _ws_chart, grid = write_charts(wb, panels, st, all_charts=args.all_charts)
    for ws_s, gs, anchor, title, xlabel, unit in slope_jobs:
        write_slope_chart(grid, ws_s, gs, anchor, title, xlabel, unit)
    pn_items = [it for it in items if it.label.startswith("SpotPN@")]
    vt_groups = next((gs for k, gs in by_kind if k == "vtune"), [])
    write_pn_chart(wb, grid, vt_groups or [g for _k, gs in by_kind for g in gs],
                   pn_items, st)
    if extra:
        write_locked(wb, extra, items, st, tname)

    out = args.out or os.path.splitext(args.path)[0] + "_summary.xlsx"
    wb.save(out)
    print("\n已写出: %s" % os.path.abspath(out))
    print("  第 1 页「%s」= 原始数据原样保留；新增 汇总 / 明细 / 斜率 / 图表 / PN曲线数据 / 锁定点"
          % ws.title)
    print("  汇总页的 Spec Min/Max 两列留空，填进去 PASS/FAIL 自动出、超规自动标红。")


if __name__ == "__main__":
    main()
