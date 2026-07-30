#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
summarize_chips.py — 一个目录里的多颗芯片 → 一份给评审看的汇总 Excel

输入是这样一棵树（一层芯片目录，芯片编号＝目录名）：

    <根目录>/
        <芯片1>/  <模块>PLL_Temperature_Sweep_*.xlsx     PLL 温度扫描
                <模块>VCO_*.xlsx                       VCO 开环压控
                <模块>_Current_*.xlsx                  电流（本版只清点，不处理）
        <芯片2>/  ...

输出一份 .xlsx：

    PLL_Summary  上面一张 <模块A> PLL、下面一张 <模块B> PLL 的性能汇总表。
                 列 = 测试项 | Unit | Limit | Spec(Min/Typ/Max) | 仿真(Min/Typ/Max)
                      | 汇总(Min/Typ/Max) | 判定 | 每颗芯片(常温/最低/最高温 + 全温MAX
                      + MAX出现在哪) | 备注
                 Spec 与仿真列留空给人填，填完判定列（Excel 公式）自动出 PASS/FAIL。
    温巡          每颗芯片一个**竖条**，条内每个模块两张图（压控温巡 + 频率漂移），
                 一张图只画一颗芯片一个模块（不叠线）；图下面就是它们的数据源。
    _审计         每个数字出自哪一份文件、哪些文件被跳过。**默认隐藏**

出稿纪律（跟 summarize_pll_sweep / summarize_vco_sweep 一致）
    · 正表里不写使用说明、不写告警、不写排除记录——那些打在控制台。
    · 讲不清的参数不进表：压控电压/片上温度/电流都**不进**汇总表
      （压控看「温巡」页的图，电流另有专门的表格，格式定了再加页）。
    · 判定只看 Spec Min/Max 两头，Typ 与仿真列只作对照。

用法：
    python summarize_chips.py <根目录>
    python summarize_chips.py <根目录> --dry-run          # 只清点+核对识别结果
    python summarize_chips.py <根目录> -o 汇总.xlsx
    python summarize_chips.py <根目录> --chips <芯片1>,<芯片2> --modules <模块>
"""

import argparse
import os
import re
import sys

from summarize_vco_sweep import load_vco
from sweep_lib import (
    COLOR_FLAG, COLOR_MUTED, COLOR_PASS, FILL_FAIL, FILL_PASS,
    LEG_STYLE, apply_y, as_text, axis_bounds, blank_policy,
    fmt_num, leg_series, legend_bottom, load_sweep, median, nice_step, num, put,
    stats_all, styles, style_series, txt,
)
from xlsx_formula_cache import FormulaCache

try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

VCACHE = FormulaCache()

# ---------------------------------------------------------------- 发现层

KIND_PLL, KIND_VCO, KIND_CUR = "pll", "vco", "current"
KIND_LABEL = {KIND_PLL: "PLL 温扫", KIND_VCO: "VCO 开环", KIND_CUR: "电流"}

# 文件名里的时间戳 _2026-07-28-15-27-19
TS_RE = re.compile(r"_(\d{4})-(\d{2})-(\d{2})-(\d{2})-(\d{2})-(\d{2})")
# 自己的产物 / Excel 临时文件，别把它们当输入读回来
SKIP_RE = re.compile(r"(^~\$)|(_summary\.xlsx$)|(_chips_summary\.xlsx$)", re.I)


def natkey(s):
    """A015 < A087 < A105：数字段按数值比，别按字符串比。"""
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\d+)", str(s))]


def classify(basename):
    """文件名 -> (模块, 类型)。认不出来的返回 (None, None)。

    命名约定是 `<模块><类型>_...`：拿类型关键字（current / vco / pll）在名字里
    的位置一切两半，**前半就是模块名**——所以工具里不写死任何模块名，
    换个芯片、换个模块照样认（铁律：通用引擎零真实字面量）。
    `<模块>_VCO_...` 这种带下划线的也认，尾部的分隔符剥掉。

    ★ 芯片编号**只认目录名**，文件名里的编号不参与判断。原厂模板复制粘贴
      忘改名是这条线已经出过的事故（同一份模板里 REG ADDR7/8/9 重名同源），
      文件名里的编号跟目录名对不上时按目录名走、控制台报一句。
    """
    n = os.path.splitext(basename)[0].lower()
    for kw, kind in (("current", KIND_CUR), ("vco", KIND_VCO), ("pll", KIND_PLL)):
        i = n.find(kw)
        if i < 0:
            continue
        mod = basename[:i].strip(" _-.")
        return (mod or None), kind
    return None, None


class Book:
    __slots__ = ("chip", "module", "kind", "path", "name", "ts")

    def __init__(self, chip, module, kind, path):
        self.chip, self.module, self.kind, self.path = chip, module, kind, path
        self.name = os.path.basename(path)
        m = TS_RE.search(self.name)
        self.ts = m.group(0) if m else ""

    @property
    def sort_key(self):
        # 时间戳优先（字符串可比：年-月-日-时-分-秒 定宽），没有时间戳的看 mtime
        return (self.ts, os.path.getmtime(self.path))


def discover(root, only_chips=None, only_modules=None):
    """扫目录 -> (选中的, 同类被跳过的, 认不出来的, 根目录下散放没被扫的)。"""
    dirs = sorted((d for d in os.listdir(root)
                   if os.path.isdir(os.path.join(root, d))), key=natkey)
    loose = []
    if not dirs:
        # 根目录本身就装着一颗芯片的文件：目录名当芯片号
        dirs = [""]
    else:
        # ★ 有芯片目录时只扫子目录，根目录下的散放文件一律不读——但要报出来。
        #   重组目录结构之前的旧文件常常还躺在根目录里，静默跳过等于少算一颗芯片。
        loose = [f for f in sorted(os.listdir(root))
                 if os.path.isfile(os.path.join(root, f))
                 and f.lower().endswith((".xlsx", ".xlsm")) and not SKIP_RE.search(f)]
    picked, dropped, unknown = {}, [], []
    for d in dirs:
        chip = d or os.path.basename(os.path.abspath(root))
        if only_chips and chip not in only_chips:
            continue
        sub = os.path.join(root, d) if d else root
        for f in sorted(os.listdir(sub)):
            if not f.lower().endswith((".xlsx", ".xlsm")) or SKIP_RE.search(f):
                continue
            mod, kind = classify(f)
            if mod is None or kind is None:
                unknown.append((chip, f, "模块" if kind is not None else "类型"))
                continue
            if only_modules and mod not in only_modules:
                continue
            b = Book(chip, mod, kind, os.path.join(sub, f))
            key = (chip, mod, kind)
            old = picked.get(key)
            if old is None:
                picked[key] = b
            elif b.sort_key > old.sort_key:
                picked[key] = b
                dropped.append((old, b))          # 旧的那份让位
            else:
                dropped.append((b, old))
    return picked, dropped, unknown, loose


# ---------------------------------------------------------------- 指标挑选

# 只有这些进汇总表。★ 压控/片上温度/电流一律不进：
#   压控的结论是「距轨还剩多少」，一个 Min/Max 说不清，去「温巡」页看图；
#   电流另有专门的测试表格，格式定了再单独加页。
WANT_EXACT = ["Freq_MHz", "Power_dBm", "IPN_SSB", "IPN_Omit_SSB"]
WANT_PREFIX = ["SpotPN@", "Spur@"]
# 表里的分组带（顺序＝出现在页面上的顺序）
BANDS = [("Frequency / Output", ["Freq_MHz", "Power_dBm"]),
         ("Phase Noise", ["IPN_SSB", "IPN_Omit_SSB", "SpotPN@"]),
         ("Spur", ["Spur@"])]
# 单位 -> 小数位
ND = {"MHz": 3}


def _off_key(label):
    """SpotPN@100kHz / Spur@26MHz -> 按 offset 数值排序。"""
    m = re.search(r"@([\d.]+)(k|M)Hz", label)
    if not m:
        return 1e9
    return float(m.group(1)) * (0.001 if m.group(2) == "k" else 1.0)


def canon_items(sweeps):
    """把各颗芯片各自识别出的指标对齐成一份表的行序。

    ★ 跨簿子只能按**标签**对齐，不能按列位置：同一个模板不同批次导出，
      列位有可能挪，名字才是身份。某颗芯片缺某个指标就在那一列留空。
    """
    seen, rows = {}, []
    for sw in sweeps:
        for it in sw.items:
            lb = it.label
            if lb not in WANT_EXACT and not any(lb.startswith(p) for p in WANT_PREFIX):
                continue
            if lb not in seen:
                seen[lb] = (lb, it.unit)
    out = []
    for band, keys in BANDS:
        picked = []
        for k in keys:
            if k.endswith("@"):
                picked += sorted((v for lb, v in seen.items() if lb.startswith(k)),
                                 key=lambda v: _off_key(v[0]))
            elif k in seen:
                picked.append(seen[k])
        if picked:
            out.append((band, picked))
    rows = out
    return rows


def temp_view(sw, item):
    """{温度: [该温度经过的每一次的值]}。整趟温巡四段各走一遍，同温会有 2~4 个值。"""
    d = {}
    if item is None:
        return d
    for lg in sw.legs:
        for t, v in leg_series(lg, item).items():
            d.setdefault(t, []).append(v)
    return d


def val_at(view, t):
    """某个温度点的代表值＝该温度**全部经过点**的中位数。

    取中位数比随便挑一次稳，跟「常温 Typ 取中位」是同一条规矩。
    """
    vs = view.get(t)
    return median(vs) if vs else None


def extreme_label(view, nd, which="max"):
    """(全温极值, 它出现在哪) —— **并列要说出来**。

    ★ 频率这类量记录精度只有 1 kHz，极值经常在七八个温度上并列。
      只报其中一个温度，评审会读成"最坏就发生在 85℃"，那是个不存在的结论。
      并列按**显示精度**判（表里看到的是取整后的数，标签得跟它对得上）。
    """
    if not view:
        return None, None
    agg = max if which == "max" else min
    disp = fmt_num(agg(v for vs in view.values() for v in vs), nd)
    hit = sorted(t for t, vs in view.items() if fmt_num(agg(vs), nd) == disp)
    if not hit:
        return disp, None
    if len(hit) == 1:
        return disp, f"{fmt_num(hit[0])}℃"
    if len(hit) <= 3:
        return disp, "/".join(str(fmt_num(t)) for t in hit) + "℃"
    return disp, f"{fmt_num(hit[0])}℃ 等 {len(hit)} 处"


def chip_stat(sw, label):
    """一颗芯片一个指标的 (min, typ, max, min_t, max_t)；没这个指标返回 None。"""
    it = sw.item(label)
    if it is None:
        return None
    s = stats_all(sw.legs, it, sw.room_t)
    if not s:
        return None
    return s


# ---------------------------------------------------------------- 汇总页

C_ITEM, C_UNIT, C_LIMIT = 1, 2, 3
C_SPEC, C_GAP1 = 4, 7
C_SIM, C_GAP2 = 8, 11
C_SUM, C_JUDGE, C_GAP3 = 12, 15, 16
C_CHIP0 = 17           # 第一颗芯片的第一列
# ★ 每颗芯片 7 列：三个代表温度点 + 全温 MIN/MAX 各配一列"出现在哪个温度"。
#   为什么不是 Min/Typ/Max：那三个数看不出"在哪一头坏"，而三个代表温度点
#   （常温 / 最低 / 最高）本身就把趋势说清楚了，再给全温极值收口。
#   极值列各带 @℃ ——并列时那一列会写"等 N 处"，见 extreme_label()。
CHIP_AX = 7
CHIP_W = CHIP_AX + 1   # +1 = 组间间隔列
# 芯片组里哪几列是数字（要套数字格式）：三个温度 + MIN + MAX
CHIP_NUMCOL = (0, 1, 2, 3, 5)

AXES = ["Min", "Typ", "Max"]


def _cl(i):
    from openpyxl.utils import get_column_letter
    return get_column_letter(i)


def _align(h, wrap=True):
    # 不用 cell.alignment.copy()：openpyxl 3.1 起那个 copy 是 deprecated，
    # 每调一次在控制台喷一行 DeprecationWarning
    from openpyxl.styles import Alignment
    return Alignment(horizontal=h, vertical="center", wrap_text=wrap)


def chip_col(n):
    return C_CHIP0 + n * CHIP_W


def note_col(n_chips):
    return chip_col(n_chips)


def rail_cols(n_chips):
    """组与组之间的"竖栏"列号。填成中灰＝把每一片界定成一个视觉区域。

    ★ 40 列宽的表里最先丢的信息是"我现在看的是哪一颗芯片"。
      一条实心竖栏比给整块上底色更省视觉预算（Gestalt 的 common region），
      而且灰度打印下照样成立。
    """
    return [C_GAP1, C_GAP2, C_GAP3] + [chip_col(k) + CHIP_AX for k in range(n_chips)]


def _edges(ws, r0, r1, c_first, c_last):
    """给一个列组的左右两侧加中等粗细竖边框。

    ★ 汇总组是全表唯一要"跳出来"的（判定就是对着它做的）。只靠底色不够：
      灰度打印和色弱下几种浅色会撞，边框不会。
    """
    from openpyxl.styles import Border, Side
    th, md = Side(style="thin", color="FF000000"), Side(style="medium", color="FF000000")
    for r in range(r0, r1 + 1):
        for c, which in ((c_first, "l"), (c_last, "r")):
            b = ws.cell(r, c).border
            ws.cell(r, c).border = Border(left=md if which == "l" else (b.left or th),
                                          right=md if which == "r" else (b.right or th),
                                          top=b.top, bottom=b.bottom)


def _hguide(ws, r, c0, c1):
    """一条横向导引线（中等粗细的下边框）。

    ★ 人一眼能数清的上限是 4 个（subitizing）。SpotPN 有连续 8 行，
      在 40 列宽的表上横着扫很容易串行。每 4 行给一条横线，不上底色——
      竖向已经分区了，再加行斑马就成了网格噪声。
    """
    from openpyxl.styles import Border, Side
    md = Side(style="medium", color="FF000000")
    for c in range(c0, c1 + 1):
        b = ws.cell(r, c).border
        ws.cell(r, c).border = Border(left=b.left, right=b.right, top=b.top, bottom=md)


def _chip_axes(tlabels):
    """芯片组的 7 个轴名：三个代表温度点 + 全温 MIN/MAX 各配一个 @℃。"""
    return [(t or "—") for t in tlabels] + ["MIN\n(全温)", "MIN\n@℃",
                                            "MAX\n(全温)", "MAX\n@℃"]


def _header(ws, r0, chips, st, title, n_chips, tlabels):
    """三行表头：大标题 / 组名 / 轴名。"""
    last = note_col(n_chips)
    c = put(ws, r0, C_ITEM, title, st, st["f_sep"], bold=True, align="left", size=12)
    ws.merge_cells(start_row=r0, start_column=C_ITEM, end_row=r0, end_column=last)
    c.alignment = _align("left", wrap=False)

    hr, ar = r0 + 1, r0 + 2
    for col, name in ((C_ITEM, "测试项"), (C_UNIT, "Unit"),
                      (C_LIMIT, "Limit"), (C_JUDGE, "判定")):
        put(ws, hr, col, name, st, st["f_head"], bold=True)
        put(ws, ar, col, None, st, st["f_head"])
        ws.merge_cells(start_row=hr, start_column=col, end_row=ar, end_column=col)
    chip_ax = _chip_axes(tlabels)
    groups = [(C_SPEC, "Spec（留空，填完自动判定）", AXES),
              (C_SIM, "仿真（留空）", AXES),
              (C_SUM, f"汇总 · {n_chips} 片",
               ["Min\n(各片最小)", "Typ\n(各片中位)", "Max\n(各片最大)"])]
    for n, chip in enumerate(chips):
        groups.append((chip_col(n), chip, chip_ax))
    for col, name, axes in groups:
        put(ws, hr, col, name, st, st["f_head"], bold=True)
        for j in range(1, len(axes)):
            put(ws, hr, col + j, None, st, st["f_head"])
        ws.merge_cells(start_row=hr, start_column=col, end_row=hr,
                       end_column=col + len(axes) - 1)
        fill = st["f_in"] if col in (C_SPEC, C_SIM) else st["f_head"]
        for j, lb in enumerate(axes):
            put(ws, ar, col + j, lb, st, fill, bold=True, size=9)
    # 竖栏只画在两行表头上——r0 是横跨整表的大标题，已经 merge 过，
    # 往合并区里写会抛 MergedCell read-only
    for col in rail_cols(n_chips):
        for r in (hr, ar):
            put(ws, r, col, None, st, st["f_rail"])
    put(ws, hr, last, "备注", st, st["f_head"], bold=True)
    put(ws, ar, last, None, st, st["f_head"])
    ws.merge_cells(start_row=hr, start_column=last, end_row=ar, end_column=last)
    ws.row_dimensions[ar].height = 30
    return ar + 1


def _caption(ws, r, st, n_chips):
    """口径说明：这几个数是怎么取的。评审第一句必问，写在表头下面最省事。"""
    txt_ = ("每颗芯片：三个温度列＝该温度**全部经过点**的中位数（整趟温巡会多次经过同一温度）；"
            "MIN / MAX＝全温极值，各配一列给出它出现在哪个温度（多个温度并列时写「等 N 处」）。"
            "都不含重锁瞬间的读数。　"
            "汇总列：Min 取各片 MIN 的最小、Max 取各片 MAX 的最大、Typ 取各片常温值的中位数。　"
            "判定只看 Spec 的 Min / Max 两头；Typ 与仿真列只作对照。")
    c = put(ws, r, C_ITEM, txt_, st, st["f_group"], align="left", size=9)
    ws.merge_cells(start_row=r, start_column=C_ITEM, end_row=r,
                   end_column=note_col(n_chips))
    c.alignment = _align("left", wrap=True)
    ws.row_dimensions[r].height = 26
    return r + 1


def _band(ws, r, name, st, n_chips):
    put(ws, r, C_ITEM, name, st, st["f_sep"], bold=True, align="left")
    for c in range(C_ITEM + 1, note_col(n_chips) + 1):
        put(ws, r, c, None, st, st["f_sep"])
    return r + 1


def _cond_rows(ws, r, chips, data, st, n_chips):
    """条件行：这组数在什么条件下取的。一行一个条件，每颗芯片给自己的值。"""
    def per_chip(label, fn):
        nonlocal r
        put(ws, r, C_ITEM, label, st, st["f_group"], align="left")
        for c in (C_UNIT, C_LIMIT, C_JUDGE):
            put(ws, r, c, None, st, st["f_group"])
        for c in list(range(C_SPEC, C_SUM + 3)) + [C_JUDGE]:
            put(ws, r, c, None, st, st["f_group"])
        for c in rail_cols(n_chips):
            put(ws, r, c, None, st, st["f_rail"])
        for n, chip in enumerate(chips):
            sw = data.get(chip)
            v = fn(sw) if sw is not None else "未测"
            c0 = chip_col(n)
            put(ws, r, c0, as_text(v), st, st["f_group"], size=9)
            for j in range(1, CHIP_AX):
                put(ws, r, c0 + j, None, st, st["f_group"])
            ws.merge_cells(start_row=r, start_column=c0, end_row=r,
                           end_column=c0 + CHIP_AX - 1)
        put(ws, r, note_col(n_chips), None, st, st["f_group"])
        r += 1

    def temp_range(sw):
        ts = sw.temps
        return f"{fmt_num(ts[0])} ~ {fmt_num(ts[-1])}" if ts else ""

    def n_points(sw):
        n = sum(1 for lg in sw.legs for x in lg.rows if x.kind != "lock")
        return f"{len(sw.temps)} 档 / {n} 测点"

    def cond_val(sw, col, nd=3):
        ci = sw.cols.idx(col)
        if ci is None:
            return ""
        vals = {num(x.raw[ci]) for lg in sw.legs for x in lg.rows
                if ci < len(x.raw) and num(x.raw[ci]) is not None}
        vals.discard(None)
        return "/".join(str(fmt_num(v, nd)) for v in sorted(vals)) if vals else ""

    per_chip("温度范围 (℃)", temp_range)
    per_chip("温度点数", n_points)
    per_chip("测试频点 fLO (MHz)", lambda sw: cond_val(sw, "fLO_MHz"))
    per_chip("参考 fXO (MHz)", lambda sw: cond_val(sw, "fXO_MHz"))
    per_chip("锁定方式", lambda sw: f"{len(sw.legs)} 段，每段开头重锁一次")
    return r


def _judge_formula(r):
    """判定公式：Spec 一头没填就只判填了的那头，两头都没填就不判（留空）。

    只用 Spec 的 Min/Max 判，跟 Limit 列无关——Limit 只是给填 spec 的人
    提示方向（≤/≥/range）。判的对象是**汇总列**，不是逐片判。
    """
    dmin, dmax = f"${_cl(C_SPEC)}{r}", f"${_cl(C_SPEC + 2)}{r}"
    smin, smax = f"{_cl(C_SUM)}{r}", f"{_cl(C_SUM + 2)}{r}"
    over = f"AND(COUNT({dmax})>0,COUNT({smax})>0,{smax}>{dmax})"
    under = f"AND(COUNT({dmin})>0,COUNT({smin})>0,{smin}<{dmin})"
    return (f"=IF(COUNT({dmin},{dmax})=0,\"\","
            f"IF(OR({over},{under}),\"FAIL\",\"PASS\"))")


def _result_row(ws, r, label, unit, chips, data, st, n_chips, tpick):
    nd = ND.get(unit, 2)
    put(ws, r, C_ITEM, label, st, st["f_res"], align="left")
    put(ws, r, C_UNIT, unit, st, st["f_res"], size=9)
    # Limit 只是提示：相噪/杂散越小越好 -> 填上限；频率/功率两头都可能有要求
    lim = "≤" if any(label.startswith(p) for p in ("IPN", "SpotPN@", "Spur@")) else "range"
    put(ws, r, C_LIMIT, lim, st, st["f_in"], size=9)
    for j in range(3):
        put(ws, r, C_SPEC + j, None, st, st["f_in"])
        put(ws, r, C_SIM + j, None, st, st["f_in"])
    for c in rail_cols(len(chips)):
        put(ws, r, c, None, st, st["f_rail"])

    mins, typs, maxs, marks = [], [], [], []
    for n, chip in enumerate(chips):
        sw = data.get(chip)
        s = chip_stat(sw, label) if sw is not None else None
        c0 = chip_col(n)
        vals = [None] * CHIP_AX
        if s:
            # 三个代表温度点：该温度**全部经过点**的中位数。整趟温巡会多次经过
            # 同一个温度（四段各走一遍），取中位比随便挑一次稳。
            view = temp_view(sw, sw.item(label))
            vals = [fmt_num(val_at(view, t), nd) for t in tpick]
            vals += list(extreme_label(view, nd, "min"))
            vals += list(extreme_label(view, nd, "max"))
            # ★ 汇总列从**各片显示出来的值**再聚合（先按显示精度取整再聚合）。
            #   满精度聚合更"准"，但会出现「表里那两格取一下 ≠ 汇总那格」的
            #   末位差 0.01，评审一眼看见就得解释。报表宁可自洽。
            mins.append(vals[3])
            maxs.append(vals[5])
            if vals[0] is not None:
                typs.append(vals[0])          # tpick[0] 是常温
            marks.append((chip, s))
        for j, v in enumerate(vals):
            # 前 3 列（温度）白底、后 4 列（极值）浅灰底 —— 7 列切成 3+4，
            # 两边都在"一眼能数清"的 4 个以内（见 _hguide 的说明）。
            num_col = j in CHIP_NUMCOL
            cell = put(ws, r, c0 + j, v, st,
                       st["f_res"] if j < 3 else st["f_zone"],
                       align="right" if num_col else "center",
                       size=10 if num_col else 9,
                       color=None if num_col else COLOR_MUTED)
            if v is not None and num_col:
                cell.number_format = "0." + "0" * nd

    # Max 就是某一片 MAX 列里的那个值（原样搬过来，对得上账）；
    # Typ 是各片常温值的中位数，**偶数片时不再取整**——两片的中位落在半个显示位上
    # （-85.42 与 -85.43 的中位 = -85.425），再取整就变成"两格取一下对不上"。
    # 格子里放真值，显示交给数字格式。
    agg = [min(mins) if mins else None, median(typs) if typs else None,
           max(maxs) if maxs else None]
    for j, v in enumerate(agg):
        cell = put(ws, r, C_SUM + j, v, st, st["f_sum"], bold=True, align="right")
        if v is not None:
            cell.number_format = "0." + "0" * nd
    put(ws, r, C_JUDGE, _judge_formula(r), st, st["f_res"], bold=True)

    # 备注只写事实：包络的两头各出自哪一片（值本身各片列里就有，不重复写）。
    # ★ 跟 @℃ 同一个道理——并列要说出来。量化过的量（频率）常常各片 MIN 一样，
    #   写"Min 出自 某一片"会被读成"这片最差"，那是个不存在的结论。
    def _who(vals, agg):
        got = [(c, v) for (c, _s), v in zip(marks, vals) if v is not None]
        if not got:
            return None
        tgt = agg(v for _c, v in got)
        hit = [c for c, v in got if v == tgt]
        if len(hit) == len(got) and len(got) > 1:
            return "各片相同"
        return "/".join(hit) if len(hit) <= 2 else f"{len(hit)} 片并列"

    def _phrase(kind, w):
        return f"{kind} 各片相同" if w == "各片相同" else f"{kind} 出自 {w}"

    note = ""
    if len(marks) > 1:
        a, b = _who(mins, min), _who(maxs, max)
        note = ("Min/Max 各片相同" if a == b == "各片相同"
                else f"{_phrase('Min', a)} / {_phrase('Max', b)}")
    gone = [c for c in chips if c not in {x[0] for x in marks}]
    if gone:
        note += ("；" if note else "") + f"未测: {', '.join(gone)}"
    put(ws, r, note_col(n_chips), as_text(note), st, st["f_res"],
        align="left", size=9, color=COLOR_MUTED)
    return r + 1


def _over_spec_cf(ws, r0, r1, st):
    """超规的那一格自己标红（report-forge 的视觉语言：红粗体＝超规）。

    只有判定列变红的话，一行 20 多个格子里到底是哪头超了还得自己比，
    评审时那一秒的迟疑就会变成一个问题。
    """
    from openpyxl.formatting.rule import FormulaRule
    if r1 < r0:
        return
    red = st["Font"](bold=True, color=COLOR_FLAG)
    dmin, dmax = f"${_cl(C_SPEC)}{r0}", f"${_cl(C_SPEC + 2)}{r0}"
    for col, cond in ((C_SUM, f"AND(COUNT({dmin})>0,COUNT({_cl(C_SUM)}{r0})>0,"
                              f"{_cl(C_SUM)}{r0}<{dmin})"),
                      (C_SUM + 2, f"AND(COUNT({dmax})>0,COUNT({_cl(C_SUM+2)}{r0})>0,"
                                  f"{_cl(C_SUM+2)}{r0}>{dmax})")):
        ws.conditional_formatting.add(
            f"{_cl(col)}{r0}:{_cl(col)}{r1}",
            FormulaRule(formula=[cond], font=red, stopIfTrue=False))


def _pass_fail_cf(ws, col_letter, r0, r1, st):
    from openpyxl.formatting.rule import FormulaRule
    from openpyxl.styles import PatternFill
    if r1 < r0:
        return
    rng = f"{col_letter}{r0}:{col_letter}{r1}"
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f'{col_letter}{r0}="FAIL"'],
        fill=PatternFill("solid", fgColor=FILL_FAIL, bgColor=FILL_FAIL),
        font=st["Font"](bold=True, color=COLOR_FLAG), stopIfTrue=False))
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f'{col_letter}{r0}="PASS"'],
        fill=PatternFill("solid", fgColor=FILL_PASS, bgColor=FILL_PASS),
        font=st["Font"](bold=True, color=COLOR_PASS), stopIfTrue=False))


def pick_temps(sweeps, room_target=25.0):
    """挑三个代表温度：常温 / 最低 / 最高。从数据里取，不写死。

    取全部芯片温度点的**交集**——只有大家都测过的温度，同一列横着比才有意义。
    交集空了（各片温度栅格对不上）退回并集，缺的格子自然留空。
    """
    grids = [set(sw.temps) for sw in sweeps if sw.temps]
    if not grids:
        return [None] * 3, [None] * 3
    common = set.intersection(*grids) or set.union(*grids)
    ts = sorted(common)
    room = min(ts, key=lambda t: abs(t - room_target))
    pick = [room, ts[0], ts[-1]]
    # 去重但保住列数（只测过 1~2 档温度时后面几列空着）
    out, seen = [], set()
    for t in pick:
        out.append(None if t in seen else t)
        seen.add(t)
    return out, [f"{fmt_num(t)}℃" if t is not None else None for t in out]


def _limit_dropdown(ws, r0, r1):
    from openpyxl.worksheet.datavalidation import DataValidation
    dv = DataValidation(type="list", formula1='"≤,≥,range"', allow_blank=True)
    dv.error = "只能填 ≤ / ≥ / range"
    ws.add_data_validation(dv)
    dv.add(f"{_cl(C_LIMIT)}{r0}:{_cl(C_LIMIT)}{r1}")


def write_summary(wb, tables, chips, st):
    """tables = [(模块名, {芯片: Sweep}), ...]，一个模块一张表，从上到下排。"""
    ws = wb.create_sheet("PLL_Summary")
    n = len(chips)
    ws.column_dimensions[_cl(C_ITEM)].width = 24
    ws.column_dimensions[_cl(C_UNIT)].width = 8
    ws.column_dimensions[_cl(C_LIMIT)].width = 8
    for c in list(range(C_SPEC, C_SPEC + 3)) + list(range(C_SIM, C_SIM + 3)) + \
            list(range(C_SUM, C_SUM + 3)):
        ws.column_dimensions[_cl(c)].width = 10
    ws.column_dimensions[_cl(C_JUDGE)].width = 8
    for c in (C_GAP1, C_GAP2, C_GAP3):
        ws.column_dimensions[_cl(c)].width = 2
    for k in range(n):
        for j in range(CHIP_AX):
            ws.column_dimensions[_cl(chip_col(k) + j)].width = \
                10 if j in CHIP_NUMCOL else 11        # @℃ 列要放得下"等 N 处"
        ws.column_dimensions[_cl(chip_col(k) + CHIP_AX)].width = 2
    ws.column_dimensions[_cl(note_col(n))].width = 26

    r = 1
    judged = []
    for mod, data in tables:
        sweeps = [s for s in data.values() if s is not None]
        items = canon_items(sweeps)
        tpick, tlabels = pick_temps(sweeps)
        t0 = r                                   # 这张表的第一行（大标题）
        r = _header(ws, r, chips, st, f"{mod} PLL 性能汇总", n, tlabels)
        r = _caption(ws, r, st, n)
        r = _cond_rows(ws, r, chips, data, st, n)
        j0 = r
        for band, rows in items:
            r = _band(ws, r, band, st, n)
            b0 = r
            for label, unit in rows:
                r = _result_row(ws, r, label, unit, chips, data, st, n, tpick)
            # 组内超过 4 行就每 4 行给一条横向导引线（SpotPN 有 8 行）
            if r - b0 > 4:
                for rr in range(b0 + 3, r - 1, 4):
                    _hguide(ws, rr, C_ITEM, note_col(n))
        judged.append((j0, r - 1))
        _limit_dropdown(ws, j0, r - 1)
        # 汇总组＝判定的依据，给它左右两道中等粗细竖框（灰度打印下也分得开）
        _edges(ws, t0, r - 1, C_SUM, C_SUM + 2)
        r += 2
    for a, b in judged:
        _pass_fail_cf(ws, _cl(C_JUDGE), a, b, st)
        _over_spec_cf(ws, a, b, st)
    # ★ 只冻结列不冻结行：一页上下两张表，冻住行的话滚到下面那张时，
    #   顶上钉着的还是上面那张的标题（"<模块A> PLL 性能汇总"），指着下面那张
    #   模块的数写着上面那个模块的名字——比丢掉表头更容易看错。表本身只 20 来行，
    #   在一张表里滚动表头不会跑掉。
    ws.freeze_panes = f"{_cl(C_CHIP0)}1"
    return ws


# ---------------------------------------------------------------- 温巡页

# 一个芯片竖条的宽度：8 列数据 + 1 列间隔（≈ 一张图的宽度）
STRIP_W = 9
JCOLS = ["序", "温度℃", "事件", "Vtune_V", "重锁点", "Freq_MHz", "Δf (kHz)", "重锁点"]
CHART_H = 20            # 一张图占多少行


def _journey_rows(sw):
    """按实际测试顺序摊平：[(序, 温度, 事件, vtune, freq, Δf, 是否重锁), ...]

    Δf 相对**本颗芯片自己**的第一个测点——各片绝对频率不同，
    跨片比的是漂移量，不是绝对值。
    """
    vt = sw.item("Vtune_V") or next((i for i in sw.items
                                     if "vtune" in i.label.lower()), None)
    fq = sw.item("Freq_MHz") or next((i for i in sw.items
                                      if i.cat == "Frequency"), None)
    seq = [(lg, x) for lg in sw.legs for x in lg.rows]
    f0 = None
    if fq:
        for _lg, x in seq:
            if x.vals.get(fq.col) is not None:
                f0 = x.vals[fq.col]
                break
    out = []
    for i, (lg, x) in enumerate(seq):
        is_lock = x.kind == "lock"
        v = x.vals.get(vt.col) if vt else None
        f = x.vals.get(fq.col) if fq else None
        df = None if (f is None or f0 is None) else round((f - f0) * 1000.0, 3)
        out.append((i + 1, x.temp,
                    f"锁@{fmt_num(lg.lock_temp)}℃" if is_lock else None,
                    v, f, df, is_lock))
    return out, (vt.unit if vt else "V"), f0


def _jchart(ws, kind, chip, mod, col0, r_data, n_rows, bounds, st, title_extra=""):
    """一张温巡图：横轴=测试顺序（刻度标当时的温度），重锁点单独一条红三角。

    为什么不按"vs 温度"画：温巡是有先后的，同一个温度会经过好几次；
    按温度画就把先后这一维压没了，看不出"锁完一路漂到哪、下次重锁拉回多少"。
    ★ 同一模块所有芯片用**同一个纵轴范围**，否则并排的两张图没法比。
    """
    from openpyxl.chart import LineChart, Reference
    from openpyxl.chart.axis import ChartLines

    c_val = col0 + (3 if kind == "vt" else 6)      # 值列（数据块内第 4 / 第 7 列）
    ch = LineChart()
    ch.title = (f"{chip} · {mod} " +
                ("压控电压 温巡过程" if kind == "vt" else "输出频率漂移 温巡过程")
                + title_extra)
    ch.style = 13
    ch.height, ch.width = 8.2, 14.5
    blank_policy(ch)
    ch.y_axis.title = "Vtune (V)" if kind == "vt" else "Δf (kHz)"
    ch.x_axis.title = "温度 (℃)　—　从左到右＝实际测试先后"
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ch.y_axis.majorGridlines = ChartLines()
    for c in (c_val, c_val + 1):
        ch.add_data(Reference(ws, min_col=c, min_row=r_data - 1,
                              max_row=r_data + n_rows - 1), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=col0 + 1, min_row=r_data,
                                max_row=r_data + n_rows - 1))
    style_series(ch.series[0], LEG_STYLE[0][0], "circle", size=4)
    style_series(ch.series[1], COLOR_FLAG, "triangle", line=False, size=10)
    apply_y(ch, bounds)
    legend_bottom(ch)
    skip = max(1, n_rows // 14)
    ch.x_axis.tickLblSkip = skip
    ch.x_axis.tickMarkSkip = skip
    return ch


def write_journey(wb, tables, chips, st, no_charts=False):
    """一页里：每颗芯片一个竖条；条内每模块两张图 + 两块数据。

    横着一条 band = 同一个模块同一个指标的各芯片对照；竖着一条 = 同一颗芯片。
    """
    ws = wb.create_sheet("温巡")
    n = len(chips)
    for k in range(n):
        c0 = 1 + k * STRIP_W
        for j in range(len(JCOLS)):
            ws.column_dimensions[_cl(c0 + j)].width = 9 if j else 5
        ws.column_dimensions[_cl(c0 + STRIP_W - 1)].width = 2

    c = put(ws, 1, 1, "温度巡回过程 —— 一颗芯片一竖条；每张图只画一颗芯片的一个模块。"
                      "横轴按实际测试先后排，刻度标的是当时的温度；红三角 = 重锁点。"
                      "同一模块各片共用一个纵轴范围，可以直接横向比。",
            st, st["f_sep"], bold=True, align="left")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n * STRIP_W)
    c.alignment = _align("left", wrap=True)
    ws.row_dimensions[1].height = 30
    for k, chip in enumerate(chips):
        put(ws, 2, 1 + k * STRIP_W, chip, st, st["f_head"], bold=True, size=12)
        for j in range(1, STRIP_W):
            put(ws, 2, 1 + k * STRIP_W + j, None, st, st["f_head"])
        ws.merge_cells(start_row=2, start_column=1 + k * STRIP_W,
                       end_row=2, end_column=k * STRIP_W + STRIP_W)

    # 先把数据块写在下半页（图放上半页，图的引用照样指得到）
    prepared = {}
    r_data = 3 + (0 if no_charts else 2 * len(tables) * CHART_H)
    for mod, data in tables:
        head_row = r_data + 1
        maxn = 0
        for k, chip in enumerate(chips):
            sw = data.get(chip)
            c0 = 1 + k * STRIP_W
            put(ws, r_data, c0, f"{mod} 温巡数据（按测试顺序）" if sw is not None
                else f"{mod}：{chip} 未测", st, st["f_sep"], bold=True, align="left")
            for j in range(1, STRIP_W):
                put(ws, r_data, c0 + j, None, st, st["f_sep"])
            ws.merge_cells(start_row=r_data, start_column=c0,
                           end_row=r_data, end_column=c0 + STRIP_W - 1)
            for j, h in enumerate(JCOLS):
                put(ws, head_row, c0 + j, h, st, st["f_head"], bold=True, size=9)
            if sw is None:
                continue
            rows, vunit, f0 = _journey_rows(sw)
            maxn = max(maxn, len(rows))
            for i, (seq, temp, ev, v, f, df, is_lock) in enumerate(rows):
                r = head_row + 1 + i
                fill = st["f_group"] if is_lock else st["f_res"]
                put(ws, r, c0 + 0, seq, st, fill, size=9)
                put(ws, r, c0 + 1, fmt_num(temp), st, fill, size=9)
                put(ws, r, c0 + 2, ev, st, fill, size=9, bold=is_lock,
                    color=COLOR_FLAG if is_lock else None)
                for j, val in ((3, v), (4, v if is_lock else None),
                               (5, f), (6, df), (7, df if is_lock else None)):
                    cell = put(ws, r, c0 + j, fmt_num(val, 4 if j < 5 else 3), st, fill,
                               size=9)
                    if val is not None:
                        cell.number_format = "0.0000" if j < 5 else \
                            ("0.000" if j == 5 else "0.0")
            prepared.setdefault(mod, {})[chip] = (head_row + 1, len(rows), rows, f0)
        r_data = head_row + 1 + maxn + 1

    if no_charts:
        return ws

    # 纵轴范围：同一模块所有芯片一起算，横向才可比
    row = 3
    for m, (mod, data) in enumerate(tables):
        got = prepared.get(mod, {})
        vt_all = [x[3] for ch_ in got.values() for x in ch_[2]]
        df_all = [x[5] for ch_ in got.values() for x in ch_[2]]
        b_vt = axis_bounds(vt_all)
        dfs = [abs(x) for x in df_all if x is not None]
        # ★ 频率漂移图故意用宽窗：记录精度只有 1 kHz，贴着数据画会把量化噪声
        #   放大成满屏方波，评审第一句必然是"频率为什么在跳"。
        lim = max(3.0 * max(dfs), 5.0) if dfs else 5.0
        step = nice_step(2 * lim)
        lim = -(-lim // step) * step
        b_df = (-lim, lim, step)
        for kind, bounds, band in (("vt", b_vt, 0), ("df", b_df, 1)):
            for k, chip in enumerate(chips):
                if chip not in got:
                    continue
                first, cnt, _rows, f0 = got[chip]
                extra = (f"（相对首点 {fmt_num(f0, 6)} MHz）"
                         if kind == "df" and f0 is not None else "")
                ch = _jchart(ws, kind, chip, mod, 1 + k * STRIP_W, first, cnt,
                             bounds, st, extra)
                ws.add_chart(ch, f"{_cl(1 + k * STRIP_W)}{row + band * CHART_H}")
        row += 2 * CHART_H
    return ws


# ---------------------------------------------------------------- VCO 页

# 结论行的定义全部复用 summarize_vco_sweep.build_conclusion（Fmin/Fmax 怎么由两段
# 扫描拼出来、Margin 怎么算、Kvco 怎么逐区间取——都是逐轮改出来的，不重新发明）。
# 这里只做两件事：删掉不进评审表的行、把它按芯片横排。
VCO_DROP = {
    "CT Band Step (avg)",     # 平均步长下不了判断，spec 只约束最坏情况
    "Drift in CT Codes",      # 与 Freq Drift + CT Band Step 强冗余（≈ 前者÷后者）
    "Current_mA",             # 电流另有专门的测试表格
}
VCO_DROP_PREFIX = ("Kvco @",)  # 工作点那个值落在 Kvco min/max 之间，判定上冗余
# ★ 粗码温度下算不出真步长：高低温只测 5 个码（0/64/128/192/255），
#   "相邻两码的频率差"实际是跨 64 个码的平均，跟常温那个数不是一回事。
#   宁可留空——填一个差 50 倍的数比没有数糟得多。
VCO_COARSE_BLANK = {"CT Band Step (max)"}


def vco_rows(sw, ref_temp, op_vtune):
    """一颗芯片一个模块的结论行 -> [(cat, item, unit, dir, kind, {温度: 值})]"""
    from summarize_vco_sweep import build_conclusion
    rows, temps = build_conclusion(sw.by_kind, sw.items, sw.freq_item, ref_temp,
                                   sw.fvco, sw.fvco_ref, {}, op_vtune)
    coarse = coarse_temps(sw)
    out = []
    for d in rows:
        if d["item"] in VCO_DROP or d["item"].startswith(VCO_DROP_PREFIX):
            continue
        vals = dict(d["vals"])
        if d["item"] in VCO_COARSE_BLANK:
            for t in coarse:
                vals.pop(t, None)
        out.append((d["cat"], d["item"], d["unit"], d["dir"], d["kind"], vals))
    return out, temps


def coarse_temps(sw):
    """CT 扫的码不是逐码密扫的那些温度。相邻码步长 > 1 就算粗扫。"""
    from summarize_vco_sweep import group_series
    out = set()
    for kind, groups in sw.by_kind:
        if kind != "ct":
            continue
        for g in groups:
            xs = sorted(group_series(g, sw.freq_item))
            if len(xs) < 2:
                continue
            if min(b - a for a, b in zip(xs, xs[1:])) > 1:
                out.add(g.temp)
    return out


def op_vtune_of(sw):
    """CT 扫把 Vtune 钉在哪个值上——工作点就取它。"""
    for kind, groups in sw.by_kind:
        if kind == "ct":
            for g in groups:
                xs = [r.vt for r in g.rows if r.vt is not None]
                if xs:
                    return xs[0]
    return None


def write_vco_summary(wb, vtables, chips, st, vtemps):
    """VCO 汇总表：每颗芯片的组内轴＝三个温度（不是 Min/Typ/Max）。

    ★ 为什么这里用温度当轴、PLL 那页用极值当轴：VCO 这些量本来就是"一个温度
      算出一个数"，而温度只有 3 档。把 3 个数压成 Min/Typ/Max 还是占 3 列，
      却把"哪个温度"丢了——纯信息损失。PLL 那边是 49 个测点压成 3 个数，
      压缩有真收益。
    """
    ws = wb.create_sheet("VCO_Summary")
    n = len(chips)
    nax = len(vtemps)
    cw = nax + 1

    def vchip(k):
        return C_CHIP0 + k * cw

    def vnote():
        return vchip(n)

    def vrails():
        return [C_GAP1, C_GAP2, C_GAP3] + [vchip(k) + nax for k in range(n)]

    ws.column_dimensions[_cl(C_ITEM)].width = 26
    ws.column_dimensions[_cl(C_UNIT)].width = 9
    ws.column_dimensions[_cl(C_LIMIT)].width = 8
    for c in list(range(C_SPEC, C_SPEC + 3)) + list(range(C_SIM, C_SIM + 3)) + \
            list(range(C_SUM, C_SUM + 3)):
        ws.column_dimensions[_cl(c)].width = 10
    ws.column_dimensions[_cl(C_JUDGE)].width = 8
    for c in (C_GAP1, C_GAP2, C_GAP3):
        ws.column_dimensions[_cl(c)].width = 2
    for k in range(n):
        for j in range(nax):
            ws.column_dimensions[_cl(vchip(k) + j)].width = 11
        ws.column_dimensions[_cl(vchip(k) + nax)].width = 2
    ws.column_dimensions[_cl(vnote())].width = 34

    r = 1
    judged = []
    for mod, data in vtables:
        t0 = r
        # ---- 三行表头 ----
        c = put(ws, r, C_ITEM, f"{mod} VCO 开环特性汇总", st, st["f_sep"],
                bold=True, align="left", size=12)
        ws.merge_cells(start_row=r, start_column=C_ITEM, end_row=r, end_column=vnote())
        c.alignment = _align("left", wrap=False)
        hr, ar = r + 1, r + 2
        for col, name in ((C_ITEM, "测试项"), (C_UNIT, "Unit"),
                          (C_LIMIT, "Limit"), (C_JUDGE, "判定")):
            put(ws, hr, col, name, st, st["f_head"], bold=True)
            put(ws, ar, col, None, st, st["f_head"])
            ws.merge_cells(start_row=hr, start_column=col, end_row=ar, end_column=col)
        groups = [(C_SPEC, "Spec（留空，填完自动判定）", AXES, 3),
                  (C_SIM, "仿真（留空）", AXES, 3),
                  (C_SUM, f"汇总 · {n} 片",
                   ["Min\n(全部片全温)", "Typ\n(各片常温中位)", "Max\n(全部片全温)"], 3)]
        for k, chip in enumerate(chips):
            groups.append((vchip(k), chip,
                           [f"{fmt_num(t)}℃" for t in vtemps], nax))
        for col, name, axes, w in groups:
            put(ws, hr, col, name, st, st["f_head"], bold=True)
            for j in range(1, w):
                put(ws, hr, col + j, None, st, st["f_head"])
            ws.merge_cells(start_row=hr, start_column=col, end_row=hr,
                           end_column=col + w - 1)
            fill = st["f_in"] if col in (C_SPEC, C_SIM) else st["f_head"]
            for j, lb in enumerate(axes):
                put(ws, ar, col + j, lb, st, fill, bold=True, size=9)
        for col in vrails():
            for rr in (hr, ar):
                put(ws, rr, col, None, st, st["f_rail"])
        put(ws, hr, vnote(), "备注", st, st["f_head"], bold=True)
        put(ws, ar, vnote(), None, st, st["f_head"])
        ws.merge_cells(start_row=hr, start_column=vnote(), end_row=ar,
                       end_column=vnote())
        ws.row_dimensions[ar].height = 30
        r = ar + 1

        cap = ("每颗芯片三列＝三个温度各自算出来的值（不是极值，VCO 的量按温度报）。　"
               "汇总列：Min / Max 取全部芯片全部温度的包络，Typ 取各片常温值的中位数。　"
               "判定只看 Spec 的 Min / Max 两头。　"
               "CT 全码扫只在常温做（条件行「CT Code Range (points)」括号里就是测点数），"
               "所以按相邻码算的步长只有常温有。")
        c = put(ws, r, C_ITEM, cap, st, st["f_group"], align="left", size=9)
        ws.merge_cells(start_row=r, start_column=C_ITEM, end_row=r, end_column=vnote())
        c.alignment = _align("left", wrap=True)
        ws.row_dimensions[r].height = 30
        r += 1

        # ---- 行序：按 (cat, item) 对齐各片；cat 变了插一条分组带 ----
        order, seen = [], set()
        for chip in chips:
            for cat, item, unit, dr, kind, _v in (data.get(chip) or []):
                key = (cat, item)
                if key not in seen:
                    seen.add(key)
                    order.append((cat, item, unit, dr, kind))
        j0 = None
        cur_cat = None
        band0 = None
        for cat, item, unit, dr, kind in order:
            if cat != cur_cat:
                if band0 is not None and r - band0 > 4:
                    for rr in range(band0 + 3, r - 1, 4):
                        _hguide(ws, rr, C_ITEM, vnote())
                put(ws, r, C_ITEM, cat, st, st["f_sep"], bold=True, align="left")
                for cc in range(C_ITEM + 1, vnote() + 1):
                    put(ws, r, cc, None, st, st["f_sep"])
                r += 1
                cur_cat, band0 = cat, r
            is_res = kind == "result"
            if is_res and j0 is None:
                j0 = r
            nd = 3 if unit == "MHz" else 2
            body = st["f_res"] if is_res else st["f_group"]
            put(ws, r, C_ITEM, item, st, body, align="left")
            put(ws, r, C_UNIT, unit, st, body, size=9)
            put(ws, r, C_LIMIT, dr if is_res else None, st,
                st["f_in"] if is_res else body, size=9)
            for j in range(3):
                put(ws, r, C_SPEC + j, None, st, st["f_in"] if is_res else body)
                put(ws, r, C_SIM + j, None, st, st["f_in"] if is_res else body)
            for cc in vrails():
                put(ws, r, cc, None, st, st["f_rail"])

            allv, roomv = [], []
            for k, chip in enumerate(chips):
                got = {(c_, i_): v_ for c_, i_, _u, _d, _k, v_ in (data.get(chip) or [])}
                vals = got.get((cat, item), {})
                for j, t in enumerate(vtemps):
                    v = vals.get(t)
                    disp = fmt_num(v, nd) if isinstance(v, (int, float)) else v
                    cell = put(ws, r, vchip(k) + j, disp, st, body,
                               align="right" if isinstance(disp, (int, float))
                               else "center",
                               size=10 if isinstance(disp, (int, float)) else 9)
                    if isinstance(disp, (int, float)):
                        cell.number_format = "0." + "0" * nd
                        allv.append(disp)
                        if t == _room_of(vtemps):
                            roomv.append(disp)
                put(ws, r, vchip(k) + nax, None, st, st["f_rail"])

            agg = [min(allv) if allv else None,
                   median(roomv) if roomv else None,
                   max(allv) if allv else None] if is_res else [None] * 3
            for j, v in enumerate(agg):
                cell = put(ws, r, C_SUM + j, v, st,
                           st["f_sum"] if is_res else body, bold=True, align="right")
                if v is not None:
                    cell.number_format = "0." + "0" * nd
            put(ws, r, C_JUDGE, _judge_formula(r) if is_res else None, st,
                st["f_res"] if is_res else body, bold=True)
            # 备注只写"哪一片整行没数"。不要去说某一个温度格为什么空：
            # 有些空格是设计上的（温漂对参考温自己恒等于 0，那格不是测量结果；
            # 粗码温度算不出步长），写成"缺数据"是误报。
            note = ""
            if is_res:
                miss = []
                for c_ in chips:
                    got = {(a, b): v for a, b, _u, _d, _k, v in (data.get(c_) or [])}
                    vals_ = got.get((cat, item)) or {}
                    if not any(isinstance(v, (int, float)) for v in vals_.values()):
                        miss.append(c_)
                if miss:
                    note = f"没有这一项: {', '.join(miss)}"
            put(ws, r, vnote(), as_text(note), st, body, align="left", size=9,
                color=COLOR_MUTED)
            r += 1
        if band0 is not None and r - band0 > 4:
            for rr in range(band0 + 3, r - 1, 4):
                _hguide(ws, rr, C_ITEM, vnote())
        if j0 is not None:
            judged.append((j0, r - 1))
            _limit_dropdown(ws, j0, r - 1)
        _edges(ws, t0, r - 1, C_SUM, C_SUM + 2)
        r += 2
    for a, b in judged:
        _pass_fail_cf(ws, _cl(C_JUDGE), a, b, st)
        _over_spec_cf(ws, a, b, st)
    ws.freeze_panes = f"{_cl(C_CHIP0)}1"
    return ws


def _room_of(vtemps, target=25.0):
    return min(vtemps, key=lambda t: abs(t - target)) if vtemps else None


# ---- VCO 图 --------------------------------------------------------------

VSTRIP_W = 9            # 一颗芯片一竖条：8 列数据 + 1 列间隔
VCHART_H = 20


def _vco_series(sw, temps):
    """{温度: [(Vtune, F), ...]} 与 {温度: [(Vtune中点, Kvco), ...]} 与 {温度: [(码, F), ...]}"""
    from summarize_vco_sweep import group_series, slopes
    fv, kv, fc = {}, {}, {}
    for kind, groups in sw.by_kind:
        for g in groups:
            if g.temp is None:
                continue
            ser = group_series(g, sw.freq_item)
            pts = sorted(ser.items())
            if kind == "vtune":
                fv[g.temp] = pts
                kv[g.temp] = [(sl[0], abs(sl[1])) for sl in slopes(g, sw.freq_item)]
            elif kind == "ct":
                fc[g.temp] = pts
    return fv, kv, fc


def write_vco_charts(wb, vtables, chips, st, vtemps, no_charts=False):
    """一页里：每颗芯片一竖条，条内三张图（F-Vtune / Kvco-Vtune / F-CT码）+ 数据源。

    横着一条 band = 同一张图的各片对照；同一模块各片共用纵轴范围，才能直接比。
    """
    ws = wb.create_sheet("VCO压控")
    n = len(chips)
    for k in range(n):
        c0 = 1 + k * VSTRIP_W
        for j in range(VSTRIP_W - 1):
            ws.column_dimensions[_cl(c0 + j)].width = 10
        ws.column_dimensions[_cl(c0 + VSTRIP_W - 1)].width = 2

    c = put(ws, 1, 1, "VCO 开环压控 —— 一颗芯片一竖条，一张图只画一颗芯片的一个模块。"
                      "同一模块各片共用纵轴范围，可以直接横向比。"
                      "★CT 全码扫只在常温：高低温只测了几个粗码，那两条只打记号不连线"
                      "（5 个点跨 64 个码直连出来是一条不存在的曲线）。",
            st, st["f_sep"], bold=True, align="left")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n * VSTRIP_W)
    c.alignment = _align("left", wrap=True)
    ws.row_dimensions[1].height = 32
    for k, chip in enumerate(chips):
        put(ws, 2, 1 + k * VSTRIP_W, chip, st, st["f_head"], bold=True, size=12)
        for j in range(1, VSTRIP_W):
            put(ws, 2, 1 + k * VSTRIP_W + j, None, st, st["f_head"])
        ws.merge_cells(start_row=2, start_column=1 + k * VSTRIP_W,
                       end_row=2, end_column=(k + 1) * VSTRIP_W)

    nb = 3 * len(vtables)          # 每个模块 3 个 band
    r_data = 3 + (0 if no_charts else nb * VCHART_H)
    prepared = {}
    for mod, data in vtables:
        for tag, cols_of in (("v", lambda ts: ["Vtune (V)"] + [f"F@{fmt_num(t)}℃"
                                                               for t in ts]),
                             ("k", lambda ts: ["Vtune (V)"] + [f"Kvco@{fmt_num(t)}℃"
                                                               for t in ts]),
                             ("c", lambda ts: ["CT code"] + [f"F@{fmt_num(t)}℃"
                                                             for t in ts])):
            head = cols_of(vtemps)
            title = {"v": "频率 vs Vtune", "k": "Kvco vs Vtune",
                     "c": "频率 vs CT 码"}[tag]
            head_row = r_data + 1
            maxn = 0
            for k, chip in enumerate(chips):
                c0 = 1 + k * VSTRIP_W
                ser = (data.get(chip) or {}).get(tag) or {}
                put(ws, r_data, c0, f"{mod} · {title}" if ser
                    else f"{mod} · {title}：{chip} 未测", st, st["f_sep"],
                    bold=True, align="left")
                for j in range(1, VSTRIP_W):
                    put(ws, r_data, c0 + j, None, st, st["f_sep"])
                ws.merge_cells(start_row=r_data, start_column=c0,
                               end_row=r_data, end_column=c0 + VSTRIP_W - 1)
                for j, h in enumerate(head):
                    put(ws, head_row, c0 + j, h, st, st["f_head"], bold=True, size=9)
                if not ser:
                    continue
                xs = sorted({x for t in vtemps for x, _v in ser.get(t, [])})
                maxn = max(maxn, len(xs))
                idx = {x: i for i, x in enumerate(xs)}
                for i, x in enumerate(xs):
                    put(ws, head_row + 1 + i, c0, x, st, st["f_res"], size=9,
                        align="right")
                for j, t in enumerate(vtemps):
                    for x, v in ser.get(t, []):
                        cell = put(ws, head_row + 1 + idx[x], c0 + 1 + j,
                                   fmt_num(v, 3), st, st["f_res"], size=9,
                                   align="right")
                        cell.number_format = "0.000"
                prepared.setdefault((mod, tag), {})[chip] = (head_row + 1, len(xs), ser)
            r_data = head_row + 1 + maxn + 1

    if no_charts:
        return ws

    row = 3
    for mod, _data in vtables:
        for tag in ("v", "k", "c"):
            got = prepared.get((mod, tag), {})
            allv = [v for ch in got.values() for t in vtemps
                    for _x, v in ch[2].get(t, [])]
            bounds = axis_bounds(allv)
            for k, chip in enumerate(chips):
                if chip not in got:
                    continue
                first, cnt, ser = got[chip]
                ws.add_chart(_vco_chart(ws, tag, chip, mod, 1 + k * VSTRIP_W,
                                        first, cnt, vtemps, bounds, ser),
                             f"{_cl(1 + k * VSTRIP_W)}{row}")
            row += VCHART_H
    return ws


def _vco_chart(ws, tag, chip, mod, col0, r_data, n_rows, vtemps, bounds, ser):
    from openpyxl.chart import Reference, ScatterChart, Series
    ch = ScatterChart()
    ch.title = f"{chip} · {mod} " + {"v": "频率 vs Vtune",
                                     "k": "Kvco vs Vtune",
                                     "c": "频率 vs CT 码"}[tag]
    ch.style = 13
    ch.height, ch.width = 8.2, 14.5
    blank_policy(ch)
    ch.x_axis.title = "CT code" if tag == "c" else "Vtune (V)"
    ch.y_axis.title = {"v": "F (MHz)", "k": "Kvco (MHz/V)", "c": "F (MHz)"}[tag]
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    xref = Reference(ws, min_col=col0, min_row=r_data, max_row=r_data + n_rows - 1)
    for j, t in enumerate(vtemps):
        pts = ser.get(t, [])
        if not pts:
            continue
        yref = Reference(ws, min_col=col0 + 1 + j, min_row=r_data - 1,
                         max_row=r_data + n_rows - 1)
        s = Series(yref, xref, title_from_data=True)
        color, sym = LEG_STYLE[j % len(LEG_STYLE)]
        # ★ 点数少的那几条只打记号不连线：高低温 CT 只测 5 个粗码，
        #   跨 64 个码直连出来是一条不存在的曲线，看图的人会以为中间测过。
        sparse = tag == "c" and len(pts) < max(8, n_rows // 4)
        style_series(s, color, sym, line=not sparse,
                     size=8 if sparse else (3 if len(pts) > 60 else 5))
        ch.series.append(s)
    apply_y(ch, bounds)
    legend_bottom(ch)
    return ch


# ---------------------------------------------------------------- 审计页

def write_audit(wb, picked, dropped, unknown, failed, notes, st):
    """每个数出自哪份文件 + 哪些文件没用上。**隐藏页**——正表不写这些。"""
    ws = wb.create_sheet("_审计")
    ws.sheet_state = "hidden"
    for i, w in enumerate((44, 10, 12, 70, 14), 1):
        ws.column_dimensions[_cl(i)].width = w
    r = 1
    put(ws, r, 1, "数据来源（每颗芯片每个模块用的是哪一份文件）", st, st["f_sep"],
        bold=True, align="left")
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
    r += 1
    for i, h in enumerate(["芯片", "模块", "类型", "文件", "规模"], 1):
        put(ws, r, i, h, st, st["f_head"], bold=True)
    r += 1
    for (chip, mod, kind), b in sorted(picked.items(), key=lambda x: natkey(x[0])):
        put(ws, r, 1, chip, st, st["f_res"])
        put(ws, r, 2, mod, st, st["f_res"])
        put(ws, r, 3, KIND_LABEL[kind], st, st["f_res"])
        put(ws, r, 4, b.name, st, st["f_res"], align="left", size=9)
        put(ws, r, 5, notes.get(id(b), "未读（本版不处理这类）"), st, st["f_res"], size=9)
        r += 1
    r += 1
    for title, rowsrc in (
            ("同类多份，只用了时间戳最新的那份（下面是被跳过的）",
             [(b.chip, b.module, KIND_LABEL[b.kind], b.name, f"让位给 {w.ts or w.name}")
              for b, w in dropped]),
            ("文件名认不出模块/类型，没读", [(c, "", "", f, f"认不出{why}")
                                            for c, f, why in unknown]),
            ("读失败", [(b.chip, b.module, KIND_LABEL[b.kind], b.name, why)
                        for b, why in failed])):
        if not rowsrc:
            continue
        put(ws, r, 1, title, st, st["f_sep"], bold=True, align="left")
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
        r += 1
        for tup in rowsrc:
            for i, v in enumerate(tup, 1):
                put(ws, r, i, as_text(v), st, st["f_res"], align="left", size=9)
            r += 1
        r += 1
    return ws


# ---------------------------------------------------------------- main

def main():
    ap = argparse.ArgumentParser(
        description="多芯片测试目录 → 一份给评审看的汇总 Excel")
    ap.add_argument("root", help="根目录（下面一层是芯片目录，目录名＝芯片编号）")
    ap.add_argument("-o", "--out", default=None,
                    help="输出路径（默认 <根目录名>_chips_summary.xlsx，"
                         "放在根目录**旁边**，免得下次扫描把它当输入读回来）")
    ap.add_argument("--modules", default="",
                    help="只处理这几个模块，并按这个顺序上下排（逗号分隔）。"
                         "不给＝从文件名前缀自动认出全部模块，按名字排序")
    ap.add_argument("--chips", default="", help="只处理这几颗（逗号分隔）")
    ap.add_argument("--leg-col", default="Mode", help="判断重锁用的列（默认 Mode）")
    ap.add_argument("--lock-pattern", default=r"_lock$",
                    help="该列匹配这个正则的行 = 一次重锁（默认 _lock$）")
    ap.add_argument("--temp-col", default=None, help="温度列（默认自动找 Temperature）")
    ap.add_argument("--no-charts", action="store_true", help="图都不画，只出数据块")
    ap.add_argument("--no-vco", action="store_true",
                    help="不做 VCO 两页（只出 PLL 温扫那两页）")
    ap.add_argument("--ref-temp", type=float, default=25.0,
                    help="VCO 温漂的参考温度（默认 25）")
    ap.add_argument("--op-vtune", type=float, default=None,
                    help="VCO 工作点调谐电压 V（默认取 CT 扫钉住的那个值）")
    ap.add_argument("--no-audit", action="store_true", help="连隐藏的 _审计 页都不要")
    ap.add_argument("--dry-run", action="store_true", help="只清点和核对识别结果")
    args = ap.parse_args()

    root = os.path.abspath(args.root)
    if not os.path.isdir(root):
        sys.exit(f"不是目录: {root}")
    try:
        import openpyxl  # noqa: F401
    except ImportError:
        sys.exit("缺少 openpyxl，请先: pip install openpyxl")

    want_mod = [m.strip() for m in args.modules.split(",") if m.strip()]
    only = {c.strip() for c in args.chips.split(",") if c.strip()} or None
    picked, dropped, unknown, loose = discover(root, only, set(want_mod) or None)
    if not picked:
        sys.exit(f"{root} 下没找到能认出来的 .xlsx —— 文件名要长成 "
                 f"`<模块><类型>_...`（类型＝ PLL / VCO / Current）")

    chips = sorted({k[0] for k in picked}, key=natkey)
    # 模块名是从文件名前缀认出来的，不写死在代码里；--modules 可以钉死顺序
    modules = want_mod or sorted({k[1] for k in picked})
    print(f"根目录: {root}")
    print(f"芯片 {len(chips)} 颗: {', '.join(chips)}")
    print(f"模块 {len(modules)} 个: {', '.join(modules)}"
          + ("" if want_mod else "（从文件名认出来的，表的上下顺序按名字排；"
                                 "要换顺序用 --modules）"))
    grid = {}
    for (chip, mod, kind), b in picked.items():
        grid.setdefault(kind, {}).setdefault(mod, {})[chip] = b
    for kind in (KIND_PLL, KIND_VCO, KIND_CUR):
        for mod in modules:
            got = grid.get(kind, {}).get(mod, {})
            if got:
                miss = [c for c in chips if c not in got]
                print(f"  {mod:<6} {KIND_LABEL[kind]:<8} {len(got)} 份"
                      + (f"   缺: {', '.join(miss)}" if miss else ""))
    for b, w in dropped:
        print(f"  ↷ 跳过 {b.chip}/{b.name}（同类里有更新的 {w.ts or w.name}）")
    for chip, f, why in unknown:
        print(f"  ? 认不出{why}，没读: {chip}/{f}")
    if loose:
        print(f"  ⚠ 根目录下还散放着 {len(loose)} 个 .xlsx **没有被扫**"
              f"（有芯片目录时只扫子目录）——如果它们也是要算的芯片，"
              f"各自挪进对应的芯片目录：")
        for f in loose:
            print(f"      {f}")
    n_cur = sum(len(v) for v in grid.get(KIND_CUR, {}).values())
    if n_cur:
        print(f"  ⚠ 发现 {n_cur} 份电流文件——**本版不处理**（电流表格式未定，"
              f"定了再单独加页）")

    # ---- 读 PLL 温扫 ----
    tables, failed, notes, warn_seen, vcharts = [], [], {}, {}, []
    for mod in modules:
        books = grid.get(KIND_PLL, {}).get(mod, {})
        if not books:
            continue
        data = {}
        print(f"\n=== {mod} PLL 温扫 ===")
        for chip in chips:
            b = books.get(chip)
            if b is None:
                print(f"  {chip}: 没有这个模块的温扫文件")
                continue
            try:
                sw = load_sweep(b.path, leg_col=args.leg_col,
                                lock_pattern=args.lock_pattern,
                                temp_col=args.temp_col, keep_original=False)
            except Exception as e:                    # noqa: B902
                failed.append((b, f"{type(e).__name__}: {e}"))
                print(f"  {chip}: 读失败 —— {e}")
                continue
            data[chip] = sw
            n_meas = sum(1 for lg in sw.legs for x in lg.rows if x.kind != "lock")
            notes[id(b)] = f"{sw.n_rows}行×{sw.n_cols}列"
            print(f"  {chip}: {len(sw.legs)} 段 / {len(sw.temps)} 档温度 / "
                  f"{n_meas} 测点 / 指标 {len(sw.items)} 个 / 排除 {len(sw.excluded)} 行"
                  f"   [{b.name}]")
            flos = {txt(x.raw[sw.cols.idx('fLO_MHz')])
                    for lg in sw.legs for x in lg.rows
                    if sw.cols.idx('fLO_MHz') is not None} - {""}
            if len(flos) > 1:
                print(f"     ⚠ 这份簿里有多个 fLO：{sorted(flos)}"
                      f"——汇总把它们并成一行了，要分频点报的话得分块（本版未做）")
            for w in sw.warnings:
                warn_seen.setdefault(w, []).append(chip)
        if data:
            tables.append((mod, data))
        # 告警按条汇总打印：重名列这种模板老问题每份簿都会报，
        # 逐份逐条打 = 30 行同样的话，真正要看的东西反而被冲掉
        for w, who in warn_seen.items():
            print(f"  ⚠ {w}"
                  + (f"   （{len(who)}/{len(data)} 份都有）" if len(who) > 1
                     else f"   （{who[0]}）"))
        warn_seen.clear()

    # ---- 读 VCO 开环 ----
    vtables, vtemps = [], []
    for mod in modules:
        books = grid.get(KIND_VCO, {}).get(mod, {})
        if not books or args.no_vco:
            continue
        vdata, vrows = {}, {}
        print()
        print("=== %s VCO 开环 ===" % mod)
        op = None
        for chip in chips:
            b = books.get(chip)
            if b is None:
                print(f"  {chip}: 没有这个模块的开环文件")
                continue
            try:
                sw = load_vco(b.path, temp_col=args.temp_col, keep_original=False)
            except Exception as e:                    # noqa: B902
                failed.append((b, f"{type(e).__name__}: {e}"))
                print(f"  {chip}: 读失败 —— {e}")
                continue
            # 工作点用第一颗芯片的（CT 扫钉住的那个 Vtune）统一喂给全部芯片：
            # 否则各片的组名会变成「@ Vtune 0.4V」「@ Vtune 0.45V」，行对不齐
            if op is None:
                op = args.op_vtune if args.op_vtune else op_vtune_of(sw)
            rows, temps = vco_rows(sw, args.ref_temp, op)
            fv, kv, fc = _vco_series(sw, temps)
            vrows[chip] = rows
            vdata[chip] = {"v": fv, "k": kv, "c": fc}
            notes[id(b)] = f"{sw.ws_val.max_row}行×{sw.ws_val.max_column}列"
            coarse = coarse_temps(sw)
            print(f"  {chip}: 温度 {[fmt_num(t) for t in temps]} / "
                  f"结论行 {sum(1 for x in rows if x[4] == 'result')} 条 / "
                  f"指标 {len(sw.items)} 个   [{b.name}]")
            if coarse:
                print(f"     · CT 粗码温度 {[fmt_num(t) for t in sorted(coarse)]}"
                      f"（只测几个码，按相邻码算的步长那几格留空）")
            for w in sw.warnings:
                warn_seen.setdefault(w, []).append(chip)
            for t in temps:
                if t not in vtemps:
                    vtemps.append(t)
        for w, who in warn_seen.items():
            print(f"  ⚠ {w}   （{'/'.join(who)}）")
        warn_seen.clear()
        if vrows:
            vtables.append((mod, vrows))
            vcharts.append((mod, vdata))
    vtemps.sort()

    if not tables and not vtables:
        sys.exit("没有一份 PLL 温扫 / VCO 开环文件读成功")

    if args.dry_run:
        print("\n--dry-run：没有写文件。")
        return

    # ---- 写出 ----
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    st = styles()
    if tables:
        write_summary(wb, tables, chips, st)
        write_journey(wb, tables, chips, st, no_charts=args.no_charts)
    if vtables:
        write_vco_summary(wb, vtables, chips, st, vtemps)
        write_vco_charts(wb, vcharts, chips, st, vtemps, no_charts=args.no_charts)
    if not args.no_audit:
        write_audit(wb, picked, dropped, unknown, failed, notes, st)
    wb.calculation.fullCalcOnLoad = True

    out = args.out or os.path.join(os.path.dirname(root),
                                   os.path.basename(root) + "_chips_summary.xlsx")
    os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
    wb.save(out)
    n_fill, n_strip = VCACHE.inject(out)
    print(f"\n已写出: {os.path.abspath(out)}")
    print("  可见页: " + " / ".join(s.title for s in wb.worksheets
                                    if s.sheet_state == "visible"))
    print("  各汇总页的 Spec / 仿真 / Limit 列留空，填进 Spec Min/Max "
          "判定列自动出 PASS/FAIL 并上色。")
    if n_strip:
        print(f"  判定列 {n_strip} 格清掉了空缓存值（Excel 打开自己算）。")


if __name__ == "__main__":
    main()
