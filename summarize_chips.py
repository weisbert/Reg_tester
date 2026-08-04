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


def print_excluded(excluded, indent="     ", cap=14):
    """排除的行逐行列出原因，不静默丢弃。

    ★ 只打一个"排除 N 行"是不够的：N 对不上预期时没法判断到底丢了什么。
      这条纪律单簿脚本里一直有，跨芯片脚本一开始漏了——第一次跑真数据
      就撞上了（预期 3 行、实际 6 行，只能靠猜）。
    """
    for xl, why in excluded[:cap]:
        print(f"{indent}行{xl}: {why}")
    if len(excluded) > cap:
        print(f"{indent}…还有 {len(excluded) - cap} 行（全部内容在隐藏的 _审计 页）")


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
# ★★ 每颗芯片只有三列：常温 / 最低温 / 最高温。
#   曾经是 7 列（再带全温 MIN/MAX 各配一个 @℃），一轮轮加出来的，每列单看都有
#   理由，加在一起就成了"要读说明才看得懂的表"——而这份簿子的第一条要求就是
#   「不能出现给人 debug 用的信息」。被砍掉的是：@℃（极值在哪个温度，对下判断
#   没有决策价值，还会写出"等 7 处"这种噪声）、逐片 MIN/MAX（跟三个温度列是
#   两种统计量，摆在一行里读起来自相矛盾，得靠 caption 解释）、同温重复性。
#   现在一行里只有一种数：**该温度的实测值**。汇总列就是对这三列取，
#   谁都能用眼睛验证，不需要任何脚注。
CHIP_AX = 3
CHIP_W = CHIP_AX + 1   # +1 = 组间间隔列
CHIP_NUMCOL = (0, 1, 2)

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
    """芯片组的轴名就是三个温度，别的什么都不放。"""
    return [(t or "—") for t in tlabels]


def fold_temp_cols(ws, col_of, n_chips, nax, keep=0):
    """--slim：每片只留常温那一列，其余温度列收进 Excel 大纲（默认折起）。

    ★ 为什么是"折"不是"删"：汇总列的 Min/Max 就是这些格子里的值。真删掉，
      汇总那几个数在表上就没了出处（"导出来的数必须能用表上的格子推出来"），
      脚本自查第①条会当场 exit 1。折起来一个数都没少，点 ＋ 就能核。
      6 片时实测区 24 列 → 12 列，一屏看得完。
    ★ keep 不写死 0：PLL 页的第一列就是常温，VCO 页的三列是升序温度
      （−40 在最左），留错列的话折完剩下的是低温那列。
    """
    from openpyxl.worksheet.properties import Outline
    # summaryRight=False ＝ 汇总列在明细列**左边**，＋/− 按钮就画在留下的那列头上；
    # 默认的 True 会把按钮画到右边的灰竖栏上（点得到，但看着像点在缝里）
    ws.sheet_properties.outlinePr = Outline(summaryBelow=True, summaryRight=False)
    n = 0
    for k in range(n_chips):
        c0 = col_of(k)
        for j in range(nax):
            if j == keep:
                continue
            d = ws.column_dimensions[_cl(c0 + j)]
            d.outlineLevel = 1
            d.hidden = True
            n += 1
        ws.column_dimensions[_cl(c0 + keep)].collapsed = True
    return n


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
    # ★ 表头只写这一列**是什么**，不写"该怎么用"。"（留空，填完自动判定）"
    #   "(各片最小)" 这类是写给填表的人的，评审看见只会觉得奇怪——
    #   designer 本来就知道 Spec 列填什么，Min 是什么也不用括号解释。
    groups = [(C_SPEC, "Spec", AXES),
              (C_SIM, "仿真", AXES),
              (C_SUM, f"汇总 · {n_chips} 片", AXES)]
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


# ★ 这里原来有一行 `_caption()`：「每片三列＝该温度的实测值。汇总列 Min / Max 就是
#   这些格子里的最小 / 最大…填 Spec 的 Min / Max，判定列自动出 PASS / FAIL。」
#   2026-08-04 整行删掉——**从头到尾都是在教人怎么读表和怎么填表**，评审要的是数。
#   列名已经写着 25℃ / Min / Spec，再解释一遍只会招来"这行字是干嘛的"。
#   口径要留档就进隐藏的 _审计 页和控制台，正表里不写。


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
            # ★ 汇总列就是对**表上这几个格子**取最小/最大/常温。用满精度、或者
            #   改成对逐点取极值，都会出现"表里的格子跟汇总对不上"——那种表
            #   没人敢信，而且解释起来要写一段 caption。宁可自洽。
            got = [v for v in vals if v is not None]
            if got:
                mins.append(min(got))
                maxs.append(max(got))
            if vals[0] is not None:
                typs.append(vals[0])          # tpick[0] 是常温
            marks.append((chip, s))
        for j, v in enumerate(vals):
            cell = put(ws, r, c0 + j, v, st, st["f_res"])
            if v is not None:
                cell.number_format = "0." + "0" * nd

    # 三个数都是从上面那些格子里直接取的，看的人用眼睛就能验证
    agg = [min(mins) if mins else None, median(typs) if typs else None,
           max(maxs) if maxs else None]
    for j, v in enumerate(agg):
        cell = put(ws, r, C_SUM + j, v, st, st["f_sum"], bold=True)
        if v is not None:
            cell.number_format = "0." + "0" * nd
    put(ws, r, C_JUDGE, _judge_formula(r), st, st["f_res"], bold=True)

    # ★ 备注只在真的缺东西时才写。以前这里写「Min 出自 XX / 同温重复性 0.8@65℃」
    #   那类工程内部信息——那是 debug 用的，不该出现在给人 review 的表上。
    gone = [c for c in chips if c not in {x[0] for x in marks}]
    note = f"未测: {', '.join(gone)}" if gone else ""
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


def write_summary(wb, tables, chips, st, slim=False):
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
    if slim:
        # tpick = [常温, 最低温, 最高温]，留第 0 列
        fold_temp_cols(ws, chip_col, n, CHIP_AX, keep=0)
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


def _jchart(ws, kind, chip, mod, col0, r_data, n_rows, bounds, st, title_extra="",
            rows=None):
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
    # ★ 压控温巡图标出全程最高/最低两点的数值：这张图的全部意义就是"压控走到哪了、
    #   离轨还剩多少"，那两个数就是结论本身，不该让人在网格里用眼睛估。
    #   （频率漂移那张不标：它故意用了远宽于数据的窗口——记录精度只有 1 kHz，
    #   标出来是 ±1~2 kHz，等于给一条平线打标签。）
    if kind == "vt" and rows:
        from summarize_vco_sweep import _label_points
        pts = [(x[3], i) for i, x in enumerate(rows) if x[3] is not None]
        if pts:
            _label_points(ch.series[0], [min(pts)[1], max(pts)[1]], numfmt="0.0###")
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

    # ★ 只写这一页是什么，不写"该怎么看"。原来这里有一整句版式说明
    #   （一片一竖条／横轴按测试先后／共用纵轴范围）——芯片名就写在每条竖条上面，
    #   横轴刻度和图例自己会说话，那句话是写给我自己的。
    c = put(ws, 1, 1, "温度巡回过程", st, st["f_sep"], bold=True, align="left")
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
                first, cnt, _rows, f0 = got[chip]      # _rows 用来定位极值点
                extra = (f"（相对首点 {fmt_num(f0, 6)} MHz）"
                         if kind == "df" and f0 is not None else "")
                ch = _jchart(ws, kind, chip, mod, 1 + k * STRIP_W, first, cnt,
                             bounds, st, extra, rows=_rows)
                ws.add_chart(ch, f"{_cl(1 + k * STRIP_W)}{row + band * CHART_H}")
        row += 2 * CHART_H
    return ws


# ---------------------------------------------------------------- VCO 页

# 结论行的定义全部复用 summarize_vco_sweep.build_conclusion（Fmin/Fmax 怎么由两段
# 扫描拼出来、Margin 怎么算、Kvco 怎么逐区间取——都是逐轮改出来的，不重新发明）。
# 这里只做两件事：删掉不进评审表的行、把它按芯片横排。
VCO_DROP = {
    "CT Band Step (avg)",       # 平均步长下不了判断，spec 只约束最坏情况
    "Drift in CT Codes",        # 与 Freq Drift + CT Band Step 强冗余（≈ 前者÷后者）
    "Current_mA",               # 电流另有专门的测试表格
    "Margin to fVCO (low)",     # 用户 2026-07-30 删：Fmin/Fmax 直接对 Spec 判就够了
    "Margin to fVCO (high)",
    "Kvco max/min",             # 用户 2026-07-30 删
    # ★★ 下面两行一起删（用户两次说 Sub-band Tuning Range 看不懂）。
    #   它们唯一回答的问题是"有没有锁不上的盲区"（子带宽度必须大于相邻码步长），
    #   而**这份测试已经用实测回答了那个问题**：闭环锁定行在三个温度都锁上了、
    #   PLL 温扫在 16 个温度里一直保持锁定。真有盲区落在目标频率上，它就锁不上。
    #   所以这两行是设计余量的话题，不是这颗芯片的判定项。
    #   只删分子（子带宽度）会留下一个没有分子的分母，比两个都留更糟——
    #   所以 CT Band Step (max) 一起走。它还有个毛病：只有常温算得出来，
    #   一行三格里两格是空的，本身就长得像"缺数据"。
    #   要查随时有：单簿脚本 summarize_vco_sweep.py 的结论页两行都还在。
    "Sub-band Tuning Range",
    "CT Band Step (max)",
    "Target fVCO",              # 用户 2026-07-30 删（Margin 那两行已经没了，它成了孤立参照）
    # ★ Kvco min/max 删（用户 2026-07-30）：它们是"逐区间斜率的极值"，回答的是
    #   「带内增益平不平」——那是个**形状**问题，一条曲线一秒看明白、两个数字反而
    #   看不明白（用户看着它们问"这到底是啥"就是证据）。而且表里那两个数**就是**
    #   VCO压控 页 Kvco-vs-Vtune 那条曲线的最低点和最高点，同一件事说了两遍。
    #   形状交给图（图上给最高/最低点打了数值标注），表里换成两个能下判断的数：
    #   Kvco average（可核）+ Kvco @工作点（环路带宽真正用的那个）。
    "Kvco min",
    "Kvco max",
}
# 判定方向的补丁：build_conclusion 里 Kvco 那几行 dir 是空的（单簿页只作对照），
# 跨芯片表要判定就得写明方向。Kvco 的 spec 通常给一个范围：
# 太小则环路带宽不够/锁不住，太大则杂散与噪声恶化。
VCO_LIMIT_FIX = {"Kvco average": "range"}
# 备注补一句"这是什么"。★ 只补定义，不补解读——评审要的是"这个数怎么来的"。
VCO_NOTE_ADD = {}      # 解释性的尾巴全去掉了，备注只留算式
# ★ 备注整句改写：让算式里的每一项都**正好是表上某一行的名字**，
#   这样导出来的数用眼睛就能核（用户原话："缺少了计算他们的中间值，我有些没安全感"）。
# ★ 备注里**只留算式**。原来每条后面还跟着一句解释（"＝整个电容阵列能覆盖多宽"
#   "两项都在上面的实测端点行里，可以直接核"）——评审知道 Tuning Range 是什么，
#   也不用我告诉他去哪儿核；那些话跟表头里的"（各片最小）"是同一类。
VCO_NOTE_SET = {
    "Kvco average": "［F(Vtune=最高) − F(Vtune=最低)］÷ Vtune 扫的跨度",
    "Fmin": "CT 扫最低频 −［F(Vtune=0.4V) − F(Vtune=最低)］",
    "Fmax": "CT 扫最高频 +［F(Vtune=最高) − F(Vtune=0.4V)］",
    "Tuning Range": "Fmax − Fmin",
    "CT Band Coverage": "CT 扫最高频 − CT 扫最低频",
    # 纯出处（"原表 Temperature 列"）不进正表：那是给我自己对数用的
    "Temperature": "",
}
# ★ 空的：`Kvco @工作点` 请回来了。我当初提议删它的理由是"它落在 Kvco min/max
#   之间、判定上冗余"——那个理由依赖 min/max 存在。min/max 一走，它就是环路带宽
#   真正用的那个数，而且是全表唯一说得清"工作点增益"的行。
VCO_DROP_PREFIX = ()


def vco_rows(sw, ref_temp, op_vtune):
    """一颗芯片一个模块的结论行 -> [(cat, item, unit, dir, kind, {温度: 值})]"""
    from summarize_vco_sweep import build_conclusion
    rows, temps = build_conclusion(sw.by_kind, sw.items, sw.freq_item, ref_temp,
                                   sw.fvco, sw.fvco_ref, {}, op_vtune)
    out = []
    for d in rows:
        if d["item"] in VCO_DROP or d["item"].startswith(VCO_DROP_PREFIX):
            continue
        vals = dict(d["vals"])
        note = VCO_NOTE_SET.get(d["item"], d.get("note") or "")
        add = VCO_NOTE_ADD.get(d["item"]) if d["item"] not in VCO_NOTE_SET else None
        if add:
            note = f"{note}；{add}" if note else add
        out.append({"cat": d["cat"], "item": d["item"], "unit": d["unit"],
                    "dir": VCO_LIMIT_FIX.get(d["item"], d["dir"]),
                    "kind": d["kind"], "vals": vals, "note": note})
    # 原料行**各自插进自己那个组**的最前面（先摆实测端点、再摆拿它们算出来的数）。
    # 整块插在一处会让 Frequency Range / CT Band 两个组头各出现两次。
    eps = endpoint_rows(sw)
    for ep in reversed(eps):
        i = next((k for k, x in enumerate(out) if x["cat"] == ep["cat"]), len(out))
        out.insert(i, ep)
    out = _kvco_average(out, eps, sw)
    out = _drift_at_op(out, ref_temp)
    return out, temps


def _kvco_average(out, eps, sw):
    """给 Kvco 组补一行 `Kvco average`＝端点算的全程平均斜率。

    ★ 它是 Kvco 那一组里**唯一能用表上的格子核出来**的数：
      ［F(Vtune=最高) − F(Vtune=最低)］÷ Vtune 扫的跨度，两项都在实测端点行里。
      旁边的 `Kvco @工作点` 是工作点那一段的局部斜率——环路带宽用的是它，
      但它推不出来，靠 VCO压控 页那张 Kvco-vs-Vtune 图作证。
      两个都放：一个可核、一个有工程含义，缺哪个都不完整。
    ★ 曲线弯的时候这两个数会差很多（真数据里逐区间斜率有 7 倍散布），
      那不是矛盾——正是"一个数说不清 Kvco"的证据，图才是主证据。
    """
    lo = next((e for e in eps if e["item"] == "F(Vtune=最低)"), None)
    hi = next((e for e in eps if e["item"] == "F(Vtune=最高)"), None)
    if lo is None or hi is None:
        return out
    span = _vtune_span(sw)
    if not span:
        return out
    vals = {t: (hi["vals"][t] - lo["vals"][t]) / span
            for t in hi["vals"] if t in lo["vals"]}
    if not vals:
        return out
    row = {"cat": "Kvco", "item": "Kvco average", "unit": "MHz/V",
           "dir": VCO_LIMIT_FIX.get("Kvco average", ""), "kind": "result",
           "vals": {t: abs(v) for t, v in vals.items()},
           "note": VCO_NOTE_SET["Kvco average"]}
    i = next((k for k, x in enumerate(out) if x["cat"] == "Kvco"), len(out))
    out.insert(i, row)
    return out


def _vtune_span(sw):
    from summarize_vco_sweep import group_series
    xs = set()
    for kind, groups in sw.by_kind:
        if kind != "vtune":
            continue
        for g in groups:
            xs |= set(group_series(g, sw.freq_item))
    return (max(xs) - min(xs)) if len(xs) >= 2 else None


def _drift_at_op(out, ref_temp):
    """温漂改成**工作点**的漂移：Freq@T − Freq@常温，两项都是表上 `Freq` 那一行。

    ★ 原来是"14 个 Vtune 点各算一个差、取绝对值最大的那个"——14 个中间量全在表外，
      看的人只能选择信。改成工作点之后零新增行、完全可核，而且符合这条线早就定过的
      规矩：性能按工作点报，不跨扫描取最差。
      漂移量在不同 Vtune 上真差很多的话，那是另一个更该被单独看见的现象，
      不该被一个"取最坏"悄悄吸收掉。
    """
    fr = next((x for x in out if x["cat"].startswith("@ Vtune")
               and x["item"] == "Freq"), None)
    di = next((k for k, x in enumerate(out)
               if x["item"].startswith("Freq Drift")), None)
    if fr is None or di is None:
        return out
    ts = [t for t, v in fr["vals"].items() if isinstance(v, (int, float))]
    if not ts:
        return out
    ref = min(ts, key=lambda t: abs(t - ref_temp))
    base = fr["vals"][ref]
    out[di] = dict(out[di],
                   item=f"Freq Drift vs {fmt_num(ref)}℃",
                   vals={t: fr["vals"][t] - base for t in ts if t != ref},
                   note=f"@工作点 那一行的 Freq：F(T) − F({fmt_num(ref)}℃)")
    return out


def flat_vtune_temps(sw, ratio=0.1):
    """哪些温度的 Vtune 扫「频率几乎不动」。返回 [(温度, 该温度频率跨度, 各温中位)]。

    ★ 为什么要专门查这个：Kvco 是逐区间 ΔF/ΔVtune 算的。如果某个温度实际上没真正
      切到开环（环路还锁着 / DAC 没驱动到管子 / testmux 没切），那个温度的频率
      就不跟 Vtune 走，Kvco 会算出接近 0 的值——而 0.00 MHz/V 这种数看着像
      "低温增益低"，其实是"这个温度的扫描没生效"。工具不能安静地报它。
      判据用相对量（跟其他温度的中位数比），不写死绝对阈值：不同 VCO 差几十倍。
    """
    from summarize_vco_sweep import group_series
    spans = {}
    for kind, groups in sw.by_kind:
        if kind != "vtune":
            continue
        for g in groups:
            if g.temp is None:
                continue
            ser = group_series(g, sw.freq_item)
            if len(ser) >= 2:
                xs = sorted(ser)
                spans[g.temp] = abs(ser[xs[-1]] - ser[xs[0]])
    if len(spans) < 2:
        return []
    med = median(list(spans.values()))
    if not med:
        return []
    return [(t, sp, med) for t, sp in sorted(spans.items()) if sp < ratio * med]


def endpoint_rows(sw):
    """两条扫描的四个端点频率。

    ★ Fmin / Fmax / Tuning Range / CT Band Coverage 全是拿这四个数算出来的。
      只给算出来的结果、不给原料，看的人没法核对，只能选择信或者不信——
      用户原话「缺少了计算他们的中间值，我有些没安全感」。摆出来之后
      备注里的算式每一项都能在表上找到对应的行。
      这四行是**原料**不是被考核项，所以走条件行（米色），不带汇总和判定。
    """
    from summarize_vco_sweep import group_series
    out = []
    for kind, cat, lo_hi, note_fmt in (
            ("vtune", "Frequency Range",
             ("F(Vtune=最低)", "F(Vtune=最高)"),
             "Vtune 扫的{}设定点（%s V，CT 码固定不动）"),
            ("ct", "CT Band",
             ("CT 扫最低频", "CT 扫最高频"),
             "CT 扫{}的那一端（码 %s，Vtune 钉在 CT 扫用的那个值）")):
        per_t = {}
        for k, groups in sw.by_kind:
            if k != kind:
                continue
            for g in groups:
                if g.temp is None:
                    continue
                ser = group_series(g, sw.freq_item)
                if ser:
                    per_t[g.temp] = ser
        if not per_t:
            continue
        xs = sorted({x for ser in per_t.values() for x in ser})
        if len(xs) < 2:
            continue
        # CT 那条按"频率最低/最高"排（码号跟频率的方向不一定一致），Vtune 按设定值排
        if kind == "ct":
            f_of = {x: median([ser[x] for ser in per_t.values() if x in ser])
                    for x in (xs[0], xs[-1])}
            ends = sorted((xs[0], xs[-1]), key=lambda x: f_of[x])
        else:
            ends = [xs[0], xs[-1]]
        for j, x in enumerate(ends):
            out.append({
                "cat": cat, "item": lo_hi[j], "unit": "MHz", "dir": "",
                "kind": "cond",
                "vals": {t: ser[x] for t, ser in per_t.items() if x in ser},
                "note": note_fmt.format("最低" if j == 0 else "最高") % fmt_num(x, 4),
            })
    return out


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


def write_vco_summary(wb, vtables, chips, st, vtemps, slim=False):
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
    ws.column_dimensions[_cl(vnote())].width = 46

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
        groups = [(C_SPEC, "Spec", AXES, 3),
                  (C_SIM, "仿真", AXES, 3),
                  (C_SUM, f"汇总 · {n} 片", AXES, 3)]
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

        # ★ 这里原来有一行口径说明（每片三列是什么、汇总怎么取、判定看哪头、
        #   CT 全码扫只在常温）。2026-08-04 整行删掉：全是读表说明。
        #   "CT 全码扫只在常温" 这条**是测试条件、不是说明**，它已经在下面的
        #   条件行 `CT Code Range (points)` 里，不必在正表上再说一遍。

        # ---- 行序：按 (cat, item) 对齐各片；cat 变了插一条分组带 ----
        order, seen = [], set()
        notes_of = {}
        for chip in chips:
            for d in (data.get(chip) or []):
                key = (d["cat"], d["item"])
                notes_of.setdefault(key, d["note"])
                if key not in seen:
                    seen.add(key)
                    order.append((d["cat"], d["item"], d["unit"], d["dir"], d["kind"]))
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
                got = {(d["cat"], d["item"]): d["vals"] for d in (data.get(chip) or [])}
                vals = got.get((cat, item), {})
                for j, t in enumerate(vtemps):
                    v = vals.get(t)
                    disp = fmt_num(v, nd) if isinstance(v, (int, float)) else v
                    cell = put(ws, r, vchip(k) + j, disp, st, body,
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
                           st["f_sum"] if is_res else body, bold=True)
                if v is not None:
                    cell.number_format = "0." + "0" * nd
            put(ws, r, C_JUDGE, _judge_formula(r) if is_res else None, st,
                st["f_res"] if is_res else body, bold=True)
            # 备注只写"哪一片整行没数"。不要去说某一个温度格为什么空：
            # 有些空格是设计上的（温漂对参考温自己恒等于 0，那格不是测量结果；
            # 粗码温度算不出步长），写成"缺数据"是误报。
            # ★ 备注 = 这个数怎么算出来的（算式，来自单簿脚本的 note）。
            #   评审第一句就是"这怎么算出来的"；不写算式就得每次口头解释一遍。
            note = notes_of.get((cat, item), "")
            if is_res:
                miss = []
                for c_ in chips:
                    got = {(d["cat"], d["item"]): d["vals"]
                           for d in (data.get(c_) or [])}
                    vals_ = got.get((cat, item)) or {}
                    if not any(isinstance(v, (int, float)) for v in vals_.values()):
                        miss.append(c_)
                if miss:
                    note = (note + "；" if note else "") +                            f"没有这一项: {', '.join(miss)}"
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
    if slim and vtemps:
        # ★ 这里的列是**升序温度**，第 0 列是最低温。留常温＝留 _room_of 那一列，
        #   照搬 PLL 页的 keep=0 会折得只剩 −40℃。
        fold_temp_cols(ws, vchip, n, nax, keep=vtemps.index(_room_of(vtemps)))
    return ws


def _room_of(vtemps, target=25.0):
    return min(vtemps, key=lambda t: abs(t - target)) if vtemps else None


# ---- VCO 图 --------------------------------------------------------------

VSTRIP_W = 9            # 一颗芯片一竖条：8 列数据 + 1 列间隔
VCHART_H = 20
# 四张图：Vtune 轴给"值 + 斜率"，CT 轴也给"值 + 斜率"，对称
VCO_TAGS = ("v", "k", "c", "d")
VCO_TITLE = {"v": "频率 vs Vtune", "k": "Kvco vs Vtune", "c": "频率 vs CT 码",
             "d": "相邻码频率差 vs CT 码（仅常温）"}
VCO_XLABEL = {"v": "Vtune (V)", "k": "Vtune (V)", "c": "CT code", "d": "CT code"}
VCO_YHEAD = {"v": "F", "k": "Kvco", "c": "F", "d": "ΔF"}
VCO_YAXIS = {"v": "F (MHz)", "k": "Kvco (MHz/V)", "c": "F (MHz)",
             "d": "|ΔF| (MHz/code)"}
# 四张图都给**全图最低/最高两点**打数值标注。值图（F-vs-Vtune / F-vs-CT）的曲线
# 是单调的，全图最低/最高点正好落在两个端点上，要的就是它们；
# 斜率图上则是最平和最陡的那两个区间。
# 只标两点是这条线定过的规矩——每条线各标首末点会在同一个横轴端点上叠成一坨。
VCO_LABELED = ("v", "k", "c", "d")


def _vco_series(sw, temps):
    """四组曲线：F-vs-Vtune / Kvco-vs-Vtune / F-vs-CT码 / ΔF-vs-CT码。

    ★ ΔF-vs-CT码（相邻两个码的频率差）**只给密扫的那个温度**。高低温只测几个
      粗码，"相邻两点"跨了几十个码，算出来是那几十个码的平均，跟常温的逐码步长
      不是一个量——混在一张图上就是骗人。这也正是表里那一行被删掉的原因。
    """
    from summarize_vco_sweep import group_series, slopes
    fv, kv, fc, fd = {}, {}, {}, {}
    coarse = coarse_temps(sw)
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
                if g.temp not in coarse:
                    fd[g.temp] = [(sl[0], abs(sl[1]))
                                  for sl in slopes(g, sw.freq_item)]
    return fv, kv, fc, fd


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

    c = put(ws, 1, 1, "VCO 开环压控", st, st["f_sep"], bold=True, align="left")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n * VSTRIP_W)
    c.alignment = _align("left", wrap=True)
    ws.row_dimensions[1].height = 32
    for k, chip in enumerate(chips):
        put(ws, 2, 1 + k * VSTRIP_W, chip, st, st["f_head"], bold=True, size=12)
        for j in range(1, VSTRIP_W):
            put(ws, 2, 1 + k * VSTRIP_W + j, None, st, st["f_head"])
        ws.merge_cells(start_row=2, start_column=1 + k * VSTRIP_W,
                       end_row=2, end_column=(k + 1) * VSTRIP_W)

    nb = len(VCO_TAGS) * len(vtables)      # 每个模块几个 band
    r_data = 3 + (0 if no_charts else nb * VCHART_H)
    prepared = {}
    for mod, data in vtables:
        for tag in VCO_TAGS:
            head = [VCO_XLABEL[tag]] + [f"{VCO_YHEAD[tag]}@{fmt_num(t)}℃"
                                        for t in vtemps]
            title = VCO_TITLE[tag]
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
                    put(ws, head_row + 1 + i, c0, x, st, st["f_res"], size=9)
                for j, t in enumerate(vtemps):
                    for x, v in ser.get(t, []):
                        cell = put(ws, head_row + 1 + idx[x], c0 + 1 + j,
                                   fmt_num(v, 3), st, st["f_res"], size=9)
                        cell.number_format = "0.000"
                prepared.setdefault((mod, tag), {})[chip] = (head_row + 1, len(xs),
                                                            ser, xs)
            r_data = head_row + 1 + maxn + 1

    if no_charts:
        return ws

    row = 3
    for mod, _data in vtables:
        for tag in VCO_TAGS:
            got = prepared.get((mod, tag), {})
            allv = [v for ch in got.values() for t in vtemps
                    for _x, v in ch[2].get(t, [])]      # ch[2] = ser
            bounds = axis_bounds(allv)
            for k, chip in enumerate(chips):
                if chip not in got:
                    continue
                first, cnt, ser, xs = got[chip]
                ws.add_chart(_vco_chart(ws, tag, chip, mod, 1 + k * VSTRIP_W,
                                        first, cnt, vtemps, bounds, ser, xs),
                             f"{_cl(1 + k * VSTRIP_W)}{row}")
            row += VCHART_H
    return ws


def _vco_chart(ws, tag, chip, mod, col0, r_data, n_rows, vtemps, bounds, ser, xs):
    from openpyxl.chart import Reference, ScatterChart, Series
    # ★ 借单簿脚本那份数值标注（只标全图最低/最高两点——三条温度曲线在同一个
    #   横轴端点上值挨得很近，各标各的会叠成一坨）。放在那边是因为它先写出来的，
    #   这里不再抄一份，抄了必然漂移。
    from summarize_vco_sweep import _label_points

    ch = ScatterChart()
    ch.title = f"{chip} · {mod} " + VCO_TITLE[tag]
    ch.style = 13
    ch.height, ch.width = 8.2, 14.5
    blank_policy(ch)
    ch.x_axis.title = VCO_XLABEL[tag]
    ch.y_axis.title = VCO_YAXIS[tag]
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    xref = Reference(ws, min_col=col0, min_row=r_data, max_row=r_data + n_rows - 1)
    built = {}
    for j, t in enumerate(vtemps):
        pts = ser.get(t, [])
        if not pts:
            continue
        yref = Reference(ws, min_col=col0 + 1 + j, min_row=r_data - 1,
                         max_row=r_data + n_rows - 1)
        sr = Series(yref, xref, title_from_data=True)
        color, sym = LEG_STYLE[j % len(LEG_STYLE)]
        # ★ 点数少的那几条只打记号不连线：高低温 CT 只测几个粗码，
        #   跨几十个码直连出来是一条不存在的曲线，看图的人会以为中间测过。
        sparse = tag == "c" and len(pts) < max(8, n_rows // 4)
        style_series(sr, color, sym, line=not sparse,
                     size=8 if sparse else (3 if len(pts) > 60 else 5))
        ch.series.append(sr)
        built[t] = sr
    if tag in VCO_LABELED and built:
        pos = {x: i for i, x in enumerate(xs)}
        allp = [(y, t, x) for t in built for x, y in ser.get(t, [])]
        if allp:
            # ★ 同一个系列上的多个标注要**一次交给它**：_label_points 每次调用都
            #   重建 s.dLbls，分两次调等于第二次把第一次的覆盖掉（曲线是直线时
            #   最低点和最高点落在同一条系列上，就会只剩一个标注）。
            want = {}
            for y, t, x in {min(allp), max(allp)}:
                if x in pos:
                    want.setdefault(t, []).append(pos[x])
            for t, idxs in want.items():
                _label_points(built[t], idxs, numfmt="0.##")
    apply_y(ch, bounds)
    legend_bottom(ch)
    return ch


# ---------------------------------------------------------------- 审计页

def write_audit(wb, picked, dropped, unknown, failed, notes, st, excl_all=()):
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
                        for b, why in failed]),
            ("排除的行（逐行原因；这些行不进统计）",
             [(chip, mod, kind, f"行{xl}", why)
              for chip, mod, kind, ex in excl_all for xl, why in ex])):
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


# ---------------------------------------------------------------- 出稿自查

def selfcheck(path):
    """把刚写出来的簿子读回来自查两条。**跑完就查，不给人多敲一条命令的机会。**

    ① 汇总列的每个数字都必须能在同一行的格子里找到。
       （2026-07-30 翻过车：温度列放中位数、极值却对逐点取，出现「-40℃ 那格
       写 -54.19，MAX 写 -53.45 @-40℃」，用户当场判成算错了。）
    ② 判定列的公式格不能留下**空的缓存值**。openpyxl 有的版本会写
       `<f>公式</f><v></v>`＝在文件里写着"这公式的结果就是空"，自己机器上打开
       正常（触发了重算），发给别人就是一片空白。xlsx_formula_cache 负责清掉，
       这里负责确认真的清掉了——不同 openpyxl 版本行为不一样，不能只信一台机器。
    """
    import openpyxl
    wb = openpyxl.load_workbook(path)
    wv = openpyxl.load_workbook(path, data_only=True)
    fails, n_agg, n_f = [], 0, 0

    for name in ("PLL_Summary", "VCO_Summary"):
        if name not in wb.sheetnames:
            continue
        ws, wsv = wb[name], wv[name]
        ax = 0
        while ws.cell(3, C_CHIP0 + ax).value not in (None, ""):
            ax += 1
        w = ax + 1
        n = 0
        while ws.cell(2, C_CHIP0 + n * w).value not in (None, "", "备注"):
            n += 1
        # 常温在组里的第几列（轴头写的就是温度，挑最接近 25 的）
        ri, best = 0, None
        for j in range(ax):
            try:
                d = abs(float(str(ws.cell(3, C_CHIP0 + j).value).replace("℃", "")) - 25)
            except ValueError:
                continue
            if best is None or d < best:
                ri, best = j, d
        for r in range(1, ws.max_row + 1):
            item = ws.cell(r, 1).value
            got = [ws.cell(r, C_SUM + j).value for j in range(3)]
            if not isinstance(item, str) or not all(
                    v is None or isinstance(v, (int, float)) for v in got):
                continue
            jf = ws.cell(r, C_JUDGE).value
            if isinstance(jf, str) and jf.startswith("="):
                n_f += 1
                if wsv.cell(r, C_JUDGE).value is not None:
                    fails.append(f"{name} 第{r}行 判定列留了缓存值 "
                                 f"{wsv.cell(r, C_JUDGE).value!r}——别人打开会看到它")
            if all(v is None for v in got):
                continue
            cells = [v for v in (ws.cell(r, C_CHIP0 + k * w + j).value
                                 for k in range(n) for j in range(ax))
                     if isinstance(v, (int, float))]
            room = [v for v in (ws.cell(r, C_CHIP0 + k * w + ri).value
                                for k in range(n)) if isinstance(v, (int, float))]
            exp = [min(cells) if cells else None,
                   median(room) if room else None,
                   max(cells) if cells else None]
            for j, (x, y) in enumerate(zip(exp, got)):
                n_agg += 1
                if (x is None) != (y is None) or (
                        x is not None and abs(float(x) - float(y)) > 1e-9):
                    fails.append(f"{name}「{item}」第{j+1}个汇总格: "
                                 f"表上的格子给出 {x}，写的是 {y}")
    return fails, n_agg, n_f


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
    ap.add_argument("--slim", action="store_true",
                    help="两张汇总表每片只显示常温列，其余温度列折进 Excel 大纲"
                         "（点 ＋ 展开，数一个没少）；芯片多了用")
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
    excl_all = []
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
            print_excluded(sw.excluded)
            excl_all.append((chip, mod, KIND_LABEL[KIND_PLL], sw.excluded))
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
            fv, kv, fc, fd = _vco_series(sw, temps)
            vrows[chip] = rows
            vdata[chip] = {"v": fv, "k": kv, "c": fc, "d": fd}
            notes[id(b)] = f"{sw.ws_val.max_row}行×{sw.ws_val.max_column}列"
            coarse = coarse_temps(sw)
            print(f"  {chip}: 温度 {[fmt_num(t) for t in temps]} / "
                  f"结论行 {sum(1 for x in rows if x['kind'] == 'result')} 条 / "
                  f"指标 {len(sw.items)} 个   [{b.name}]")
            if coarse:
                print(f"     · CT 粗码温度 {[fmt_num(t) for t in sorted(coarse)]}"
                      f"（只测几个码）")
            for t, sp, med in flat_vtune_temps(sw):
                print(f"     ⚠⚠ {fmt_num(t)}℃ 的 Vtune 扫**频率几乎没动**："
                      f"全程只变了 {fmt_num(sp, 4)} MHz，其他温度是 {fmt_num(med, 1)} MHz。"
                      f"这个温度多半没真正切到开环（环路还锁着／DAC 没驱动／testmux "
                      f"没切），它的 Kvco 会算出接近 0 的值——那不是低温增益低。"
                      f"对照表上 F(Vtune=最低) 与 F(Vtune=最高) 这两行就能确认")
            print(f"     排除 {len(sw.excluded)} 行:")
            print_excluded(sw.excluded)
            excl_all.append((chip, mod, KIND_LABEL[KIND_VCO], sw.excluded))
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

    # ★★ 每一页只列**这一页真有数据**的芯片，不拿全局芯片表去铺列。
    #   性能和电流常常不是同一个人测的、测的也不是同一颗 die（一个目录只有温扫、
    #   另一个只有电流）。用全局表铺列的话，只测了电流的那颗会在 PLL / VCO 四张页上
    #   各占一整列"未测"——一列全是"未测"不是信息，是噪声，还会让人以为那颗片子
    #   测坏了。反过来电流页也一样。
    #   ★ 这不是把差异藏起来：哪颗片子有哪类数据，隐藏的 _审计 页逐份列着，
    #     控制台每次也打。正表上只摆有数的列。
    def _with_data(tbls):
        return [c for c in chips
                if any(d.get(c) is not None for _m, d in tbls)]

    chips_pll = _with_data(tables)
    chips_vco = _with_data(vtables)

    # ---- 写出 ----
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    st = styles()
    if tables:
        write_summary(wb, tables, chips_pll, st, slim=args.slim)
        write_journey(wb, tables, chips_pll, st, no_charts=args.no_charts)
    if vtables:
        write_vco_summary(wb, vtables, chips_vco, st, vtemps, slim=args.slim)
        write_vco_charts(wb, vcharts, chips_vco, st, vtemps,
                         no_charts=args.no_charts)
    if not args.no_audit:
        write_audit(wb, picked, dropped, unknown, failed, notes, st, excl_all)
    wb.calculation.fullCalcOnLoad = True

    out = args.out or os.path.join(os.path.dirname(root),
                                   os.path.basename(root) + "_chips_summary.xlsx")
    os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
    wb.save(out)
    n_fill, n_strip = VCACHE.inject(out)
    print(f"\n已写出: {os.path.abspath(out)}")
    print("  可见页: " + " / ".join(s.title for s in wb.worksheets
                                    if s.sheet_state == "visible"))
    for nm, cs in (("PLL_Summary / 温巡", chips_pll), ("VCO 两页", chips_vco)):
        if cs and set(cs) != set(chips):
            gone = [c for c in chips if c not in cs]
            print(f"  {nm} 只列了 {', '.join(cs)}；{', '.join(gone)} 没有这类数据，"
                  f"不给它留空列（谁有哪类数据见隐藏的 _审计 页）。")
    print("  各汇总页的 Spec / 仿真 / Limit 列留空，填进 Spec Min/Max "
          "判定列自动出 PASS/FAIL 并上色。")
    if args.slim:
        print("  --slim: 两张汇总表每片只显示常温列，其余温度列已折起——"
              "点表头上方的 ＋（或左上角的「2」）展开，数一个都没少。"
              "温巡页 / VCO压控页的竖条不受影响。")
    elif len(chips) >= 5:
        print(f"  提示: {len(chips)} 颗芯片＝实测区 {len(chips) * CHIP_W} 列，"
              f"横着容易数不清第几片。加 --slim 每片只显示常温列"
              f"（其余折起来，点 ＋ 就能展开核对）。")
    fails, n_agg, n_f = selfcheck(out)
    print(f"  自查: {n_agg} 个汇总格子都能在同一行的格子里找到；"
          f"{n_f} 个判定公式没有留下缓存值"
          + (f"（清掉了 {n_strip} 个空缓存）" if n_strip else "") + "。")
    if fails:
        print()
        print("  ✗ 自查没过，这份簿子先别发出去：")
        for f in fails[:20]:
            print(f"      {f}")
        sys.exit(1)


if __name__ == "__main__":
    main()
