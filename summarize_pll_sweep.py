#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
summarize_pll_sweep.py — 性能扫描簿 → 带汇总页的 Excel

输入：一份仪器/脚本导出的性能扫描宽表（一行 = 一个测试项，列里有条件也有结果）。
输出：一份新的 .xlsx——
    第 1 页  原始数据（原封不动，就是输入那一页）
    汇总      compliance 版式的对比表：指标 × 各温度段的 Min/Max/Δ + 全温 + 判定
              Spec 两列留空给人填，填完 PASS/FAIL 由 Excel 公式自动出、超规自动标红
    温度明细  指标 × 温度矩阵（按段分列），既能翻数也是图表的数据源
    图表      每个指标一张 值-vs-温度 图（每段一条线）+ 相噪-vs-offset 图
    重锁对比  每次重锁的 前 / 锁定点 / 后 三态对照（重锁把指标拉回了多少）

为什么要按「段」而不是按温度汇总
    这类簿常见的跑法是：锁一次 → 不重锁跑完一整趟温度 → 到端点再重锁。
    于是同一个温度会在升温段、降温段里各出现一次甚至多次，
    直接按温度合并会把回滞（同温不同值）抹平，正好把要看的东西看没了。
    本脚本按「重锁事件」切段，段内再按温度排。

依赖：只用 openpyxl。   pip install openpyxl

用法：
    python summarize_pll_sweep.py <扫描簿.xlsx>                  # 出 <原名>_summary.xlsx
    python summarize_pll_sweep.py <扫描簿.xlsx> -o 汇总.xlsx
    python summarize_pll_sweep.py <扫描簿.xlsx> --dry-run        # 只打印识别结果，不写文件
    python summarize_pll_sweep.py <扫描簿.xlsx> --leg-col Mode --lock-pattern "_lock$"
    python summarize_pll_sweep.py <扫描簿.xlsx> --keep-test-item PLL_Test

先跑 --dry-run 看三件事对不对：列映射 / 分了几段 / 排除了哪些行。
排除的行不会被悄悄丢掉，汇总页底部会逐行列出原因。
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

# report-forge compliance 表的配色（黄表头 / 米色条件行 / 白结果行 / 红超规）
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
        khz = mhz * 1000.0
        return f"{fmt_num(khz)}kHz"
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

    def find(self, *patterns, exclude=()):
        """按正则找第一个匹配的表头，返回 (名字, 列下标)。"""
        for k in self.pos:
            kl = k.lower()
            if any(re.search(p, kl) for p in patterns) and \
               not any(re.search(p, kl) for p in exclude):
                return k, self.pos[k][0]
        return None, None


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


def build_items(cols, rows):
    """识别有数据的结果列。结果列整片是空的（占位列）自动丢掉，并记在 dropped 里。"""
    items, dropped = [], []

    def nonempty(ci):
        return sum(1 for r in rows if ci < len(r) and not is_blank(r[ci]))

    for name, cat, unit in SIMPLE_ITEMS:
        ci = cols.idx(name)
        if ci is None:
            continue
        if nonempty(ci) == 0:
            dropped.append((name, "整列没有数据"))
            continue
        items.append(Item(cat, name, unit, ci, name))

    for prefix, cat, unit, tpl in PAIRED_ITEMS:
        i = 1
        while True:
            fc, rc = cols.idx(f"{prefix}Freq{i}"), cols.idx(f"{prefix}Result{i}")
            if fc is None and rc is None:
                break
            i += 1
            if rc is None:
                continue
            n = nonempty(rc)
            if n == 0:
                dropped.append((f"{prefix}Result{i-1}", "整列没有数据"))
                continue
            freqs = {num(r[fc]) for r in rows
                     if fc is not None and fc < len(r) and num(r[fc]) is not None}
            f = fmt_hz(sorted(freqs)[0]) if len(freqs) == 1 else (
                "/".join(fmt_hz(x) for x in sorted(freqs)[:3]) or f"#{i-1}")
            items.append(Item(cat, tpl.format(f=f), unit, rc, f"{prefix}Result{i-1}"))
    return items, dropped


# ---------------------------------------------------------------- 行分段

class Row:
    __slots__ = ("xl", "temp", "leg", "kind", "vals", "raw")

    def __init__(self, xl, temp, kind, raw):
        self.xl, self.temp, self.kind, self.raw = xl, temp, kind, raw
        self.leg, self.vals = None, {}


class Leg:
    def __init__(self, n, lock_temp):
        self.n, self.lock_temp = n, lock_temp
        self.rows = []

    @property
    def temps(self):
        return [r.temp for r in self.rows if r.temp is not None]

    @property
    def direction(self):
        t = self.temps
        if len(t) < 2:
            return ""
        return "↑" if t[-1] > t[0] else ("↓" if t[-1] < t[0] else "")

    @property
    def title(self):
        lt = fmt_num(self.lock_temp)
        return f"段{self.n} 锁@{lt}℃" if lt is not None else f"段{self.n}"

    @property
    def stage(self):
        t = self.temps
        if not t:
            return ""
        return f"{fmt_num(t[0])}→{fmt_num(t[-1])}℃ {self.direction}".strip()


def segment(rows, leg_col_i, lock_re):
    """按重锁事件切段。重锁行本身也是测量点（它有结果），算作本段第一个点。"""
    legs, cur = [], None
    orphan = []
    for r in rows:
        mode = txt(r.raw[leg_col_i]) if leg_col_i is not None and leg_col_i < len(r.raw) else ""
        if lock_re.search(mode):
            cur = Leg(len(legs) + 1, r.temp)
            legs.append(cur)
            r.kind = "lock"
        if cur is None:
            orphan.append(r)          # 第一次重锁之前的行
            continue
        r.leg = cur
        cur.rows.append(r)
    return legs, orphan


# ---------------------------------------------------------------- 统计

def leg_series(leg, item):
    """本段的「温度 -> 值」去重视图：同段同温有多个点时取最后一个。

    什么时候会同温多点：重锁行和它后面那个测量点是同一个温度（在端点停下来
    重锁再测）。取后者＝重锁后稳定下来的那个点。

    ★ 汇总统计和温度明细页**共用这一份**。否则汇总按全部行取极值、明细一格
    只放得下一个值，就会出现「汇总报的极值在明细里查无此值」——报表一旦对不上账
    就没人敢信了。锁定瞬间的值不会丢，在「重锁对比」页里完整给出。
    """
    out = OrderedDict()
    for r in leg.rows:
        v = r.vals.get(item.col)
        if r.temp is not None and v is not None:
            out[r.temp] = v
    return out


def _extremes(pairs):
    if not pairs:
        return None
    lo = min(pairs, key=lambda x: x[0])
    hi = max(pairs, key=lambda x: x[0])
    return {"min": lo[0], "min_t": lo[1], "max": hi[0], "max_t": hi[1],
            "delta": hi[0] - lo[0], "n": len(pairs)}


def stats(leg, item):
    return _extremes([(v, t) for t, v in leg_series(leg, item).items()])


def stats_all(legs, item):
    return _extremes([(v, t) for lg in legs
                      for t, v in leg_series(lg, item).items()])


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


# ---------------------------------------------------------------- 汇总页

def write_summary(wb, legs, items, meta, st):
    from openpyxl.formatting.rule import FormulaRule
    from openpyxl.styles import PatternFill
    from openpyxl.utils import get_column_letter

    ws = wb.create_sheet("汇总")
    # 列规划：Category|Item|Unit|Spec Min|Spec Max| |段i Min/Max/Δ| … | |全温 5 列| |判定
    plan = [("cat", "Category", 16), ("item", "Item", 22), ("unit", "Unit", 9),
            ("spec_min", "Min", 10), ("spec_max", "Max", 10), ("sep", "", 2)]
    groups = []                       # (标题, stage, [列下标...])
    groups.append(("Spec（自己填）", "手工填写", [3, 4]))
    for lg in legs:
        base = len(plan)
        plan += [("min", "Min", 10), ("max", "Max", 10), ("delta", "Δ", 9), ("sep", "", 2)]
        groups.append((lg.title, lg.stage, [base, base + 1, base + 2]))
    base = len(plan)
    plan += [("min", "Min", 10), ("min_t", "@℃", 8), ("max", "Max", 10),
             ("max_t", "@℃", 8), ("delta", "Δ", 9), ("sep", "", 2)]
    groups.append(("全温合计", f"{len(legs)} 段并起来", [base, base + 1, base + 2, base + 3, base + 4]))
    col_judge = len(plan)
    plan += [("judge", "判定", 12)]

    for i, (_k, _l, w) in enumerate(plan):
        ws.column_dimensions[get_column_letter(i + 1)].width = w

    # ---- 3 行表头 ----
    for r in range(1, 4):
        for c in range(1, len(plan) + 1):
            put(ws, r, c, None, st, st["f_head"])
    for c0, label in ((0, "Category"), (1, "Item"), (2, "Unit")):
        ws.merge_cells(start_row=1, start_column=c0 + 1, end_row=3, end_column=c0 + 1)
        put(ws, 1, c0 + 1, label, st, st["f_head"], bold=True)
    ws.merge_cells(start_row=1, start_column=col_judge + 1, end_row=3, end_column=col_judge + 1)
    put(ws, 1, col_judge + 1, "判定", st, st["f_head"], bold=True)
    for title, stage, cc in groups:
        ws.merge_cells(start_row=1, start_column=cc[0] + 1, end_row=1, end_column=cc[-1] + 1)
        put(ws, 1, cc[0] + 1, title, st, st["f_head"], bold=True)
        ws.merge_cells(start_row=2, start_column=cc[0] + 1, end_row=2, end_column=cc[-1] + 1)
        put(ws, 2, cc[0] + 1, stage, st, st["f_head"], bold=True, size=9)
        for ci in cc:
            put(ws, 3, ci + 1, plan[ci][1], st, st["f_head"], bold=True, size=9)
    for i, (k, _l, _w) in enumerate(plan):
        if k == "sep":
            for r in range(1, 4):
                put(ws, r, i + 1, None, st, st["f_sep"])

    # ---- 数据行 ----
    r0 = 4
    spec_min_col, spec_max_col = 4, 5                    # 1 基
    all_cols = [i for i, (k, _l, _w) in enumerate(plan)]
    for n, it in enumerate(items):
        r = r0 + n
        for ci in all_cols:
            k = plan[ci][0]
            put(ws, r, ci + 1, None, st,
                st["f_sep"] if k == "sep" else st["f_res"])
        put(ws, r, 2, it.label, st, st["f_res"], align="left")
        put(ws, r, 3, it.unit, st, st["f_res"])
        gi = 1                                            # groups[0] 是 Spec
        for lg in legs:
            s = stats(lg, it)
            cc = groups[gi][2]
            gi += 1
            if s:
                put(ws, r, cc[0] + 1, fmt_num(s["min"]), st, st["f_res"])
                put(ws, r, cc[1] + 1, fmt_num(s["max"]), st, st["f_res"])
                put(ws, r, cc[2] + 1, fmt_num(s["delta"]), st, st["f_res"])
        s = stats_all(legs, it)
        cc = groups[-1][2]
        if s:
            put(ws, r, cc[0] + 1, fmt_num(s["min"]), st, st["f_res"])
            put(ws, r, cc[1] + 1, fmt_num(s["min_t"]), st, st["f_res"], size=9)
            put(ws, r, cc[2] + 1, fmt_num(s["max"]), st, st["f_res"])
            put(ws, r, cc[3] + 1, fmt_num(s["max_t"]), st, st["f_res"], size=9)
            put(ws, r, cc[4] + 1, fmt_num(s["delta"]), st, st["f_res"])
        # 判定：Spec 两格任意一格填了就开始判；只填一边就只判一边
        L = get_column_letter
        smin, smax = f"${L(spec_min_col)}{r}", f"${L(spec_max_col)}{r}"
        amin, amax = f"{L(cc[0] + 1)}{r}", f"{L(cc[2] + 1)}{r}"
        f = (f'=IF(AND({smin}="",{smax}=""),"",'
             f'IF(AND(OR({smin}="",{amin}>={smin}),OR({smax}="",{amax}<={smax})),'
             f'"PASS","FAIL"))')
        put(ws, r, col_judge + 1, f, st, st["f_res"], bold=True)

    # Category 列：同类连续行纵向合并
    i = 0
    while i < len(items):
        j = i
        while j + 1 < len(items) and items[j + 1].cat == items[i].cat:
            j += 1
        if j > i:
            ws.merge_cells(start_row=r0 + i, start_column=1, end_row=r0 + j, end_column=1)
        put(ws, r0 + i, 1, items[i].cat, st, st["f_res"], bold=True)
        i = j + 1

    # PASS/FAIL 上色
    rng = f"{get_column_letter(col_judge + 1)}{r0}:{get_column_letter(col_judge + 1)}{r0 + len(items) - 1}"
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f'{get_column_letter(col_judge + 1)}{r0}="FAIL"'],
        fill=PatternFill("solid", fgColor=FILL_FAIL, bgColor=FILL_FAIL),
        font=st["Font"](bold=True, color=COLOR_FLAG), stopIfTrue=False))
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f'{get_column_letter(col_judge + 1)}{r0}="PASS"'],
        fill=PatternFill("solid", fgColor=FILL_PASS, bgColor=FILL_PASS),
        font=st["Font"](bold=True, color=COLOR_PASS), stopIfTrue=False))
    ws.freeze_panes = "D4"

    # ---- 表下方：怎么用 + 哪些行没算进来 ----
    r = r0 + len(items) + 2
    put(ws, r, 1, "怎么用", st, st["f_group"], bold=True, align="left")
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
    for line in [
        "① Spec 的 Min / Max 两列自己填（只关心单边就只填一边，另一边留空）。",
        "② 填完「判定」列自动出 PASS / FAIL 并上色，不用重跑脚本。",
        "③ 判定用的是「全温合计」的 Min / Max，即各段所有温度点里的极值；"
        "这里的每个极值都能在「温度明细」页原样查到。",
        "④ 同一段同一温度若测了两次（重锁行 + 重锁后那点），取后者；"
        "锁定瞬间的值在「重锁对比」页完整给出。",
        f"⑤ 按段分开是有原因的：{meta['why_legs']}",
    ]:
        r += 1
        put(ws, r, 1, line, st, None, align="left")
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=min(12, len(plan)))

    r += 2
    put(ws, r, 1, "没算进汇总的行（逐行列出，不做静默丢弃）", st, st["f_group"], bold=True, align="left")
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    r += 1
    for h, c in (("原表行号", 1), ("原因", 2)):
        put(ws, r, c, h, st, st["f_head"], bold=True)
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
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=min(12, len(plan)))
    return ws


# ---------------------------------------------------------------- 温度明细

def write_detail(wb, legs, items, st):
    """每个指标一块：行=温度（所有段的温度并集，升序），列=段。图表就吃这个。"""
    from openpyxl.utils import get_column_letter
    ws = wb.create_sheet("温度明细")
    ws.column_dimensions["A"].width = 12
    for i in range(len(legs)):
        ws.column_dimensions[get_column_letter(2 + i)].width = 13

    temps = sorted({t for lg in legs for t in lg.temps})
    blocks = []                       # (item, 首行, 末行)
    r = 1
    for it in items:
        put(ws, r, 1, f"{it.label}  [{it.unit}]", st, st["f_group"], bold=True, align="left")
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=1 + len(legs))
        r += 1
        put(ws, r, 1, "温度℃", st, st["f_head"], bold=True)
        for i, lg in enumerate(legs):
            put(ws, r, 2 + i, f"{lg.title} {lg.direction}", st, st["f_head"], bold=True, size=9)
        head = r
        r += 1
        first = r
        series = [leg_series(lg, it) for lg in legs]
        for t in temps:
            put(ws, r, 1, fmt_num(t), st, st["f_res"])
            for i, _lg in enumerate(legs):
                put(ws, r, 2 + i, fmt_num(series[i].get(t)), st, st["f_res"])
            r += 1
        blocks.append((it, head, first, r - 1))
        r += 1
    ws.freeze_panes = "B1"
    return ws, blocks, temps


# ---------------------------------------------------------------- 图表

def write_charts(wb, ws_detail, blocks, legs, items, pn_items, st):
    from openpyxl.chart import Reference, ScatterChart, Series
    from openpyxl.chart.marker import Marker

    ws = wb.create_sheet("图表")
    put(ws, 1, 1, "每张图：横轴=温度，每段一条线。同温不同线分叉 = 回滞/温漂没回来。",
        st, st["f_group"], bold=True, align="left")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=10)

    row_anchor, col_anchor = 3, 1
    for n, (it, head, first, last) in enumerate(blocks):
        ch = ScatterChart()
        ch.title = f"{it.label} vs 温度"
        ch.style = 13
        ch.x_axis.title = "温度 (℃)"
        ch.y_axis.title = it.unit
        ch.x_axis.delete = False
        ch.y_axis.delete = False
        ch.height, ch.width = 7.5, 13
        ch.dispBlanksAs = "gap"
        xref = Reference(ws_detail, min_col=1, min_row=first, max_row=last)
        for i, lg in enumerate(legs):
            yref = Reference(ws_detail, min_col=2 + i, min_row=head, max_row=last)
            s = Series(yref, xref, title_from_data=True)
            s.marker = Marker(symbol="circle", size=5)
            ch.series.append(s)
        ws.add_chart(ch, f"{'A' if n % 2 == 0 else 'I'}{row_anchor + (n // 2) * 16}")
    # 相噪 vs offset：横轴对数，每条线一个温度
    if pn_items:
        _pn_offset_chart(wb, ws, legs, pn_items, st, row_anchor + ((len(blocks) + 1) // 2) * 16)
    return ws


def _pn_offset_chart(wb, ws_chart, legs, pn_items, st, anchor_row):
    """相噪-vs-offset：数据另开一小块（行=offset，列=选中的温度），再挂散点图。"""
    from openpyxl.chart import Reference, ScatterChart, Series
    from openpyxl.chart.marker import Marker

    lg = max(legs, key=lambda x: len(x.rows))            # 取点最多的那一段
    temps = sorted({r.temp for r in lg.rows if r.temp is not None})
    if not temps:
        return
    pick = sorted({temps[0], temps[len(temps) // 2], temps[-1]})

    ws = wb.create_sheet("PN曲线数据")
    put(ws, 1, 1, f"相噪 vs offset（取 {lg.title}，各温度一条线）",
        st, st["f_group"], bold=True, align="left")
    put(ws, 2, 1, "offset_MHz", st, st["f_head"], bold=True)
    for j, t in enumerate(pick):
        put(ws, 2, 2 + j, f"{fmt_num(t)}℃", st, st["f_head"], bold=True)
    for i, it in enumerate(pn_items):
        m = re.search(r"@([\d.]+)(k|M)Hz", it.label)
        off = float(m.group(1)) / (1000.0 if m and m.group(2) == "k" else 1.0) if m else None
        put(ws, 3 + i, 1, off, st, st["f_res"])
        for j, t in enumerate(pick):
            v = None
            for row in lg.rows:
                if row.temp == t and row.vals.get(it.col) is not None:
                    v = row.vals[it.col]
            put(ws, 3 + i, 2 + j, fmt_num(v), st, st["f_res"])

    ch = ScatterChart()
    ch.title = "相噪 vs offset"
    ch.style = 13
    ch.x_axis.title = "offset (MHz)"
    ch.y_axis.title = "dBc/Hz"
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ch.x_axis.scaling.logBase = 10
    ch.height, ch.width = 8, 14
    ch.dispBlanksAs = "gap"
    xref = Reference(ws, min_col=1, min_row=3, max_row=2 + len(pn_items))
    for j in range(len(pick)):
        yref = Reference(ws, min_col=2 + j, min_row=2, max_row=2 + len(pn_items))
        s = Series(yref, xref, title_from_data=True)
        s.marker = Marker(symbol="circle", size=5)
        ch.series.append(s)
    ws_chart.add_chart(ch, f"A{anchor_row}")


# ---------------------------------------------------------------- 重锁对比

def write_relock(wb, legs, items, st):
    """每次重锁：锁定前最后一点 / 锁定点 / 锁定后下一点，三态对照。

    重锁前后一般是同一个温度（在端点停下来重锁），所以差值基本就是
    「不重锁跑了一整趟温度之后，重锁能把指标拉回多少」。
    """
    ws = wb.create_sheet("重锁对比")
    events = []
    for k, lg in enumerate(legs):
        if not lg.rows:
            continue
        lock_row = lg.rows[0]
        before = legs[k - 1].rows[-1] if k > 0 and legs[k - 1].rows else None
        after = lg.rows[1] if len(lg.rows) > 1 else None
        if before is None:
            continue
        events.append((lg, before, lock_row, after))
    if not events:
        put(ws, 1, 1, "只有一段，没有重锁事件可比。", st, None, align="left")
        return ws

    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 9
    put(ws, 1, 1, "每次重锁的前 / 锁定点 / 后：差值 = 重锁把指标拉回了多少",
        st, st["f_group"], bold=True, align="left")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=3 + 4 * len(events))

    for j, (lg, b, l, a) in enumerate(events):
        c0 = 4 + j * 4
        ws.merge_cells(start_row=2, start_column=c0, end_row=2, end_column=c0 + 3)
        put(ws, 2, c0, f"{lg.title}（{fmt_num(l.temp)}℃）", st, st["f_head"], bold=True)
        for i, h in enumerate(["锁前", "锁定点", "锁后", "Δ(锁定点-锁前)"]):
            put(ws, 3, c0 + i, h, st, st["f_head"], bold=True, size=9)
            ws.column_dimensions[ws.cell(row=3, column=c0 + i).column_letter].width = 12
        put(ws, 4, c0, f"行{b.xl} @{fmt_num(b.temp)}℃", st, st["f_group"], size=8)
        put(ws, 4, c0 + 1, f"行{l.xl}", st, st["f_group"], size=8)
        put(ws, 4, c0 + 2, f"行{a.xl}" if a else "-", st, st["f_group"], size=8)
        put(ws, 4, c0 + 3, "", st, st["f_group"])
    for h, c in (("Category", 1), ("Item", 2), ("Unit", 3)):
        ws.merge_cells(start_row=2, start_column=c, end_row=4, end_column=c)
        put(ws, 2, c, h, st, st["f_head"], bold=True)

    for n, it in enumerate(items):
        r = 5 + n
        put(ws, r, 1, it.cat, st, st["f_res"], align="left")
        put(ws, r, 2, it.label, st, st["f_res"], align="left")
        put(ws, r, 3, it.unit, st, st["f_res"])
        for j, (lg, b, l, a) in enumerate(events):
            c0 = 4 + j * 4
            vb, vl = b.vals.get(it.col), l.vals.get(it.col)
            va = a.vals.get(it.col) if a else None
            put(ws, r, c0, fmt_num(vb), st, st["f_res"])
            put(ws, r, c0 + 1, fmt_num(vl), st, st["f_res"])
            put(ws, r, c0 + 2, fmt_num(va), st, st["f_res"])
            d = (vl - vb) if (vb is not None and vl is not None) else None
            put(ws, r, c0 + 3, fmt_num(d), st, st["f_res"], bold=True)
    ws.freeze_panes = "D5"
    return ws


# ---------------------------------------------------------------- main

def main():
    ap = argparse.ArgumentParser(description="性能扫描簿 → 带汇总页的 Excel")
    ap.add_argument("path", help="扫描簿 .xlsx")
    ap.add_argument("-o", "--out", default=None, help="输出路径（默认 <原名>_summary.xlsx）")
    ap.add_argument("--sheet", default=None, help="数据所在 sheet（默认第一个）")
    ap.add_argument("--header-row", type=int, default=1, help="表头行号（1 基，默认 1）")
    ap.add_argument("--leg-col", default="Mode", help="判断重锁用的列（默认 Mode）")
    ap.add_argument("--lock-pattern", default=r"_lock$",
                    help="该列匹配这个正则的行 = 一次重锁（默认 _lock$）")
    ap.add_argument("--temp-col", default=None, help="温度列（默认自动找含 Temperature 的列）")
    ap.add_argument("--keep-test-item", default=None,
                    help="只保留 Test Item 等于该值的行（默认取出现最多的那个值）")
    ap.add_argument("--keep-mode", default=None,
                    help="只保留该模式的行（默认取出现最多的那个值）；重锁行始终保留。"
                         "挡的是夹在扫描序列里的旁路/自检行——它们也带部分结果值，"
                         "不挡就会被算进相邻那一段，把段的走向和极值都带偏")
    ap.add_argument("--dry-run", action="store_true", help="只打印识别结果，不写文件")
    args = ap.parse_args()

    if not os.path.isfile(args.path):
        sys.exit(f"找不到文件: {args.path}")
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
            warnings.append(f"重复列名 {k}：出现在 {', '.join(gl(i+1) for i in v)}，"
                            f"按名字只取到第一个（{gl(v[0]+1)}）")

    # 温度列
    if args.temp_col:
        tname, tcol = args.temp_col, cols.idx(args.temp_col)
    else:
        tname, tcol = cols.find(r"temperature", r"^temp")
    if tcol is None:
        sys.exit("找不到温度列，用 --temp-col 指定")
    leg_i = cols.idx(args.leg_col)
    if leg_i is None:
        warnings.append(f"没有 {args.leg_col} 列，无法按重锁切段——全部行算作一段")
    ti_name, ti_col = (None, None)
    if cols.idx("Test Item") is not None:
        ti_name, ti_col = "Test Item", cols.idx("Test Item")

    # 行过滤：Test Item 少数派（收尾行/模板遗留行）踢掉，但逐行记原因
    excluded = []
    keep_ti = args.keep_test_item
    if ti_col is not None and keep_ti is None:
        cnt = {}
        for r in data:
            v = txt(r[ti_col]) if ti_col < len(r) else ""
            if v:
                cnt[v] = cnt.get(v, 0) + 1
        if cnt:
            keep_ti = max(cnt, key=lambda k: cnt[k])

    lock_re = re.compile(args.lock_pattern, re.I)

    # 主模式：默认取出现最多的那个值（重锁行不参与统计，它带 _lock 后缀）
    keep_mode = args.keep_mode
    if leg_i is not None and keep_mode is None:
        cnt = {}
        for r in data:
            v = txt(r[leg_i]) if leg_i < len(r) else ""
            if v and not lock_re.search(v):
                cnt[v] = cnt.get(v, 0) + 1
        if cnt:
            keep_mode = max(cnt, key=lambda k: cnt[k])

    rows = []
    for n, raw in enumerate(data):
        xl = args.header_row + 1 + n
        if all(is_blank(v) for v in raw):
            continue
        if ti_col is not None and keep_ti is not None:
            v = txt(raw[ti_col]) if ti_col < len(raw) else ""
            if v != keep_ti:
                excluded.append((xl, f"{ti_name} = {v!r}，不是主测试项 {keep_ti!r}"))
                continue
        if leg_i is not None and keep_mode is not None:
            v = txt(raw[leg_i]) if leg_i < len(raw) else ""
            if v != keep_mode and not lock_re.search(v):
                excluded.append((xl, f"{args.leg_col} = {v!r}，不是主模式 {keep_mode!r}"))
                continue
        rows.append(Row(xl, num(raw[tcol]) if tcol < len(raw) else None, "meas", raw))

    items, dropped = build_items(cols, [r.raw for r in rows])
    if not items:
        sys.exit("没识别出任何有数据的结果列")
    for r in rows:
        for it in items:
            r.vals[it.col] = num(r.raw[it.col]) if it.col < len(r.raw) else None

    # 没有任何结果值的行（纯配置行）也踢掉
    keep = []
    for r in rows:
        if all(v is None for v in r.vals.values()):
            excluded.append((r.xl, "所有结果列都是空的（配置/开关行，不是测量点）"))
        else:
            keep.append(r)
    rows = keep

    legs, orphan = segment(rows, leg_i, lock_re)
    for r in orphan:
        excluded.append((r.xl, f"排在第一次重锁之前（{args.leg_col}={txt(r.raw[leg_i]) if leg_i is not None else '?'}）"))
    if not legs:
        legs = [Leg(1, None)]
        legs[0].rows = rows
        warnings.append(f"没有匹配 {args.lock_pattern!r} 的行，整表按一段处理")

    excluded.sort(key=lambda x: x[0])
    meta = {
        "excluded": excluded,
        "warnings": warnings,
        "why_legs": ("锁一次跑一整趟温度、到端点才重锁，所以同一个温度会在不同段里"
                     "各出现一次；按温度合并会把回滞抹平。"),
    }

    # ---- 打印识别结果 ----
    print(f"源文件 : {os.path.basename(args.path)}   sheet={ws.title}  "
          f"{ws.max_row} 行 × {ws.max_column} 列")
    print(f"温度列 : {tname}    分段列: {args.leg_col}  规则: /{args.lock_pattern}/")
    if ti_col is not None:
        print(f"主测试项: {keep_ti!r}（其余行排除）")
    if keep_mode is not None:
        print(f"主模式  : {keep_mode!r}（+ 重锁行；其余行排除）")
    print(f"识别指标 {len(items)} 个:")
    for it in items:
        print(f"    [{it.cat:<12}] {it.label:<18} {it.unit:<7} <- 列 {it.src}")
    if dropped:
        print(f"  跳过的空列 {len(dropped)} 个: " + ", ".join(k for k, _ in dropped))
    print(f"分段 {len(legs)} 段:")
    for lg in legs:
        print(f"    {lg.title:<14} {lg.stage:<16} {len(lg.rows):>3} 点  "
              f"行 {lg.rows[0].xl}~{lg.rows[-1].xl}")
    if excluded:
        print(f"排除 {len(excluded)} 行:")
        for xl, why in excluded:
            print(f"    行{xl}: {why}")
    for w in warnings:
        print(f"  ⚠ {w}")
    if args.dry_run:
        print("\n--dry-run：没有写文件。")
        return

    # ---- 写出 ----
    st = _styles()
    write_summary(wb, legs, items, meta, st)
    ws_detail, blocks, _temps = write_detail(wb, legs, items, st)
    pn_items = [it for it in items if it.label.startswith("SpotPN@")]
    write_charts(wb, ws_detail, blocks, legs, items, pn_items, st)
    write_relock(wb, legs, items, st)

    out = args.out or os.path.splitext(args.path)[0] + "_summary.xlsx"
    wb.save(out)
    print(f"\n已写出: {os.path.abspath(out)}")
    print(f"  第 1 页「{ws.title}」= 原始数据原样保留；"
          f"新增 汇总 / 温度明细 / 图表 / PN曲线数据 / 重锁对比")
    print("  汇总页的 Spec Min/Max 两列留空，填进去 PASS/FAIL 自动出、超规自动标红。")


if __name__ == "__main__":
    main()
