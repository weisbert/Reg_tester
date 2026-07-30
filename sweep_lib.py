#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
sweep_lib.py — 扫描簿报表的公共层（读数 / 分段 / 统计 / 画格子 / 图表原语）

从 summarize_pll_sweep.py 抽出来的。抽的理由：同一份原厂宽表现在有两个消费者
（单簿脚本 summarize_pll_sweep.py，跨芯片汇总 summarize_chips.py），
读取和统计的口径必须只有一份——两份必然漂移，而漂移出来的是"两份报表同一个
指标报不同的数"，那种账最难对。

这里只放**通用引擎**：没有任何真实模块名/寄存器名/频点常量。
真实的东西全靠表头名和命令行参数进来。

分四段：
    ① 取值      is_blank / num / txt / fmt_num / fmt_hz
    ② 结构      Columns（按表头定位、重名列报出来）/ Item / build_items
    ③ 行与段    Row / Leg / segment / room_temp / leg_series / stats
    ④ 装载      load_sweep()  ——  一行把一份簿子读成 (rows, legs, items)
    ⑤ 出稿      styles / put / 图表原语（配色、纵轴范围、图例）

依赖：只用 openpyxl（且只在真的要读写 xlsx 时才 import）。
"""

import re
from collections import OrderedDict

# report-forge compliance 表的配色（黄表头 / 米色条件行 / 白结果行 / 红超规）
FILL_HEADER = "FFFF00"
FILL_GROUP = "EEECE1"
FILL_RESULT = "FFFFFF"
FILL_SEP = "B8CCE4"
COLOR_FLAG = "FF0000"
COLOR_PASS = "006100"
FILL_FAIL = "FFC7CE"
FILL_PASS = "C6EFCE"
FILL_INPUT = "FFF2CC"       # 要人填的格子（Spec / 仿真值 / 判据），浅黄一眼看得出
# 宽表分区用的灰阶与低饱和色。★色相总数压在 3 个以内（黄=表头 / 浅黄=要人填 /
# 浅蓝=判定依据），其余分区一律用灰阶深浅 + 分隔，否则 40 列的表变成花布，
# 而且灰度打印和色弱下几种色相会互相撞。
FILL_ZONE = "F2F2F2"        # 组内的第二个功能区（极值列）——极浅灰
FILL_RAIL = "D9D9D9"        # 组与组之间的竖栏（中灰），界定"这一块是一片"
FILL_SUM = "DDEBF7"         # 汇总组：全表唯一需要跳出来的东西
COLOR_MUTED = "595959"      # 注释性文字（@℃ 这类），要退到背景里去

# 表里表示"没测/不适用"的占位符，一律当空值
BLANK_TOKENS = {"", "-", "--", "—", "n/a", "na", "null", "none", "#n/a"}


class SweepError(Exception):
    """簿子读不下去（缺列/没数据行）。调用方决定是退出还是跳过这一份。"""


# ---------------------------------------------------------------- ① 取值

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


def median(vals):
    v = sorted(x for x in vals if x is not None)
    if not v:
        return None
    n = len(v)
    return v[n // 2] if n % 2 else (v[n // 2 - 1] + v[n // 2]) / 2.0


# ---------------------------------------------------------------- ② 结构

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

    def dup_report(self):
        from openpyxl.utils import get_column_letter as gl
        return [f"重复列名 {k}：出现在 {', '.join(gl(i + 1) for i in v)}，"
                f"按名字只取到第一个（{gl(v[0] + 1)}）"
                for k, v in self.duplicates.items()]


class Item:
    __slots__ = ("cat", "label", "unit", "col", "src", "text_src")

    def __init__(self, cat, label, unit, col, src):
        self.cat, self.label, self.unit, self.col, self.src = cat, label, unit, col, src
        self.text_src = False       # 原始列是文本型数字？是的话引用要加双负号

    @property
    def key(self):
        """跨簿子对齐用的键。列位置在不同芯片的簿子里可能不同，名字才是身份。"""
        return self.label


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

    # 标出「文本型数字」列：Python 解得动、Excel 不认。不标出来的话，
    # 引用过去 COUNT/MIN 全落空，极值温度那格直接 #N/A。
    for it in items:
        n_txt = sum(1 for r in rows
                    if it.col < len(r) and isinstance(r[it.col], str)
                    and num(r[it.col]) is not None)
        it.text_src = n_txt > 0
    return items, dropped


# ---------------------------------------------------------------- ③ 行与段

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


def room_temp(legs, target=25.0):
    """"常温"取实际测过的、最接近 target 的那个温度。

    别用"排序后的中位数"——那是统计意义的中间，不是工程上的常温（曾挑到 35℃）。
    """
    ts = sorted({r.temp for lg in legs for r in lg.rows if r.temp is not None})
    return min(ts, key=lambda t: abs(t - target)) if ts else None


# ---------------------------------------------------------------- 统计

def leg_series(leg, item, with_lock=False):
    """本段的「温度 -> 值」去重视图：同段同温有多个点时取最后一个。

    ★ 默认剔掉重锁行。重锁是这组数据的**取得条件**，不是被考核的性能：
    把锁定瞬间的读数混进"全温最坏值"里，等于拿锁的过程去判性能的规格。
    （实际上每个重锁点后面都紧跟一个同温的稳定测量点，所以剔掉不丢温度。）

    ★ 汇总统计和温度明细页**共用这一份**。否则汇总按全部行取极值、明细一格
    只放得下一个值，就会出现「汇总报的极值在明细里查无此值」——报表一旦对不上账
    就没人敢信了。
    """
    out = OrderedDict()
    for r in leg.rows:
        if r.kind == "lock" and not with_lock:
            continue
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


def stats_all(legs, item, room_t=None):
    """全温统计。TYP 取常温点的中位数——常温在温巡里被经过好几次，
    取中位数比随便挑一次稳。"""
    s = _extremes([(v, t) for lg in legs
                   for t, v in leg_series(lg, item).items()])
    if s and room_t is not None:
        rv = [v for lg in legs for t, v in leg_series(lg, item).items() if t == room_t]
        if rv:
            s["typ"] = median(rv)
            s["typ_n"] = len(rv)
    return s


# ---------------------------------------------------------------- ④ 装载

class Sweep:
    """一份读完的扫描簿。字段名跟原来 main() 里的局部变量一一对应。"""

    def __init__(self, **kw):
        self.__dict__.update(kw)

    def item(self, *labels):
        """按标签取指标（给多个名字＝依次尝试）。"""
        for n in labels:
            for it in self.items:
                if it.label == n:
                    return it
        return None

    @property
    def temps(self):
        return sorted({r.temp for lg in self.legs for r in lg.rows
                       if r.temp is not None})


def load_sweep(path, sheet=None, header_row=1, leg_col="Mode",
               lock_pattern=r"_lock$", temp_col=None,
               keep_test_item=None, keep_mode=None, keep_original=True):
    """把一份扫描簿读成 Sweep（行已过滤、段已切好、指标已识别）。

    keep_original=True 时额外用 data_only=False 再读一遍工作簿并挂在 .wb 上，
    给"输出簿第一页保留原表"用。跨芯片汇总不需要那份，给 False 省一半内存。
    """
    import openpyxl

    # 读两份：一份取缓存值用来算，一份原封不动用来存。
    # 只用 data_only=True 那份去存的话，原表里若有公式会被替换成计算结果——
    # 「第 1 页保留原始 excel」就不成立了。
    wb_val = openpyxl.load_workbook(path, data_only=True, read_only=False)
    ws_val = wb_val[sheet] if sheet else wb_val[wb_val.sheetnames[0]]
    wb, ws = None, None
    if keep_original:
        wb = openpyxl.load_workbook(path, data_only=False)
        ws = wb[ws_val.title]

    all_rows = [list(r) for r in ws_val.iter_rows(values_only=True)]
    if len(all_rows) < header_row + 1:
        raise SweepError("表里没有数据行")
    header = all_rows[header_row - 1]
    data = all_rows[header_row:]

    cols = Columns(header)
    warnings = list(cols.dup_report())

    if temp_col:
        tname, tcol = temp_col, cols.idx(temp_col)
    else:
        tname, tcol = cols.find(r"temperature", r"^temp")
    if tcol is None:
        raise SweepError("找不到温度列，用 --temp-col 指定")

    leg_i = cols.idx(leg_col)
    if leg_i is None:
        warnings.append(f"没有 {leg_col} 列，无法按重锁切段——全部行算作一段")
    ti_name, ti_col = (None, None)
    if cols.idx("Test Item") is not None:
        ti_name, ti_col = "Test Item", cols.idx("Test Item")

    # 行过滤：Test Item 少数派（收尾行/模板遗留行）踢掉，但逐行记原因
    excluded = []
    keep_ti = keep_test_item
    if ti_col is not None and keep_ti is None:
        keep_ti = _majority(data, ti_col)

    lock_re = re.compile(lock_pattern, re.I)

    # 主模式：默认取出现最多的那个值（重锁行不参与统计，它带 _lock 后缀）
    if leg_i is not None and keep_mode is None:
        keep_mode = _majority(data, leg_i, skip_re=lock_re)

    rows, cut = [], []          # cut: (行号, 原因, 原始行) —— 稍后补"带几个结果值"
    for n, raw in enumerate(data):
        xl = header_row + 1 + n
        if all(is_blank(v) for v in raw):
            continue
        if ti_col is not None and keep_ti is not None:
            v = txt(raw[ti_col]) if ti_col < len(raw) else ""
            if v != keep_ti:
                cut.append((xl, f"{ti_name} = {v!r}，不是主测试项 {keep_ti!r}", raw))
                continue
        if leg_i is not None and keep_mode is not None:
            v = txt(raw[leg_i]) if leg_i < len(raw) else ""
            if v != keep_mode and not lock_re.search(v):
                cut.append((xl, f"{leg_col} = {v!r}，不是主模式 {keep_mode!r}", raw))
                continue
        rows.append(Row(xl, num(raw[tcol]) if tcol < len(raw) else None, "meas", raw))

    items, dropped = build_items(cols, [r.raw for r in rows])
    if not items:
        raise SweepError("没识别出任何有数据的结果列")
    for r in rows:
        for it in items:
            r.vals[it.col] = num(r.raw[it.col]) if it.col < len(r.raw) else None

    # ★ 被过滤掉的行到底有没有带测量结果，必须说出来。
    #   "排除了 6 行"这句话本身不足以判断有没有丢数据——原厂模板常在数据后面
    #   追加页脚/图例行（条件列全空），那种排掉是对的；但旁路自检行是**带部分
    #   结果值**的，把它算进相邻那一段会把段的走向和极值带偏。两者只看行号
    #   分不出来，得看它带不带结果。
    for xl, why, raw in cut:
        k = sum(1 for it in items
                if it.col < len(raw) and num(raw[it.col]) is not None)
        # 带 4/16 跟带 16/16 是两件事：前者是配置行顺手回读了几个量，
        # 后者是**另一个配置下的一整套测量**（比如关掉 test mux 再测一遍）。
        # 只打绝对个数分不出来，分母必须给。
        excluded.append((xl, why + (f"（这行带 {k}/{len(items)} 个结果值）" if k else
                                    "（这行没有任何结果值）")))

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
        excluded.append((r.xl, f"排在第一次重锁之前（"
                               f"{leg_col}={txt(r.raw[leg_i]) if leg_i is not None else '?'}）"))
    if not legs:
        legs = [Leg(1, None)]
        legs[0].rows = rows
        warnings.append(f"没有匹配 {lock_pattern!r} 的行，整表按一段处理")
    excluded.sort(key=lambda x: x[0])

    return Sweep(path=path, wb=wb, ws=ws, wb_val=wb_val, ws_val=ws_val,
                 src_title=ws_val.title, header=header, data=data, cols=cols,
                 temp_name=tname, temp_col=tcol, leg_col=leg_col, leg_i=leg_i,
                 lock_pattern=lock_pattern, lock_re=lock_re,
                 ti_name=ti_name, ti_col=ti_col, keep_ti=keep_ti, keep_mode=keep_mode,
                 rows=rows, items=items, dropped=dropped, legs=legs, orphan=orphan,
                 excluded=excluded, warnings=warnings, room_t=room_temp(legs),
                 n_rows=ws_val.max_row, n_cols=ws_val.max_column)


def _majority(data, ci, skip_re=None):
    cnt = {}
    for r in data:
        v = txt(r[ci]) if ci < len(r) else ""
        if v and not (skip_re and skip_re.search(v)):
            cnt[v] = cnt.get(v, 0) + 1
    return max(cnt, key=lambda k: cnt[k]) if cnt else None


# ---------------------------------------------------------------- ⑤ 出稿

def styles():
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    thin = Side(style="thin", color="FF000000")
    return {
        "border": Border(left=thin, right=thin, top=thin, bottom=thin),
        "center": Alignment(horizontal="center", vertical="center", wrap_text=True),
        "left": Alignment(horizontal="left", vertical="center", wrap_text=True),
        # ★ 数字列右对齐：居中的数字列小数点不对齐，比大小得逐个读；右对齐后
        #   位数差直接变成视觉长度差，扫一眼就知道谁大。这条比任何配色都管用。
        "right": Alignment(horizontal="right", vertical="center", wrap_text=False),
        "f_head": PatternFill("solid", fgColor=FILL_HEADER, bgColor=FILL_HEADER),
        "f_group": PatternFill("solid", fgColor=FILL_GROUP, bgColor=FILL_GROUP),
        "f_res": PatternFill("solid", fgColor=FILL_RESULT, bgColor=FILL_RESULT),
        "f_sep": PatternFill("solid", fgColor=FILL_SEP, bgColor=FILL_SEP),
        "f_in": PatternFill("solid", fgColor=FILL_INPUT, bgColor=FILL_INPUT),
        "f_zone": PatternFill("solid", fgColor=FILL_ZONE, bgColor=FILL_ZONE),
        "f_rail": PatternFill("solid", fgColor=FILL_RAIL, bgColor=FILL_RAIL),
        "f_sum": PatternFill("solid", fgColor=FILL_SUM, bgColor=FILL_SUM),
        "Font": Font,
    }


def put(ws, r, c, v, st, fill=None, bold=False, color=None, align="center", size=10):
    cell = ws.cell(row=r, column=c)
    cell.value = v
    cell.border = st["border"]
    cell.alignment = st.get(align) or st["center"]
    cell.font = st["Font"](bold=bold, color=color, size=size)
    if fill is not None:
        cell.fill = fill
    return cell


def as_text(v):
    """要当纯文本写进去的串。

    ★ 以 = 开头的字符串 openpyxl 会当公式写进 <f>：备注里写
    "= Fmax − Fmin" 这种算式样子，Excel 就去解析它，整列变成 #REF!。
    垫一个空格最省事，看不出来也不会被解析。
    只垫 `=`：实测 openpyxl 只按 startswith("=") 判公式，`+ - @` 照常存成字符串，
    多垫的话 "-40 ~ 105" 这种条件值会平白多一个前导空格。
    """
    if isinstance(v, str) and v[:1] == "=":
        return " " + v
    return v


# ---- 图表原语 ------------------------------------------------------------

# 每段/每条线一个颜色 + 一个记号形状：几条线叠在一张图上，光靠颜色分不开
# （打印和色弱都糊），形状也得不一样。
LEG_STYLE = [("1F77B4", "circle"), ("D62728", "square"),
             ("2CA02C", "triangle"), ("FF7F0E", "diamond"),
             ("9467BD", "x"), ("8C564B", "plus")]


def nice_step(span):
    """给定跨度挑一个好看的刻度步长（1/2/2.5/5 × 10^n）。"""
    import math
    if span <= 0:
        return 1.0
    raw = span / 5.0
    e = 10.0 ** math.floor(math.log10(raw))
    for f in (1, 2, 2.5, 5, 10):
        if raw <= f * e:
            return f * e
    return 10 * e


def axis_bounds(vals, pad=0.10):
    """按数据自己算 (min, max, 步长)，两头留一点余量再对齐到整刻度。

    ★ 不要指望 Excel 自动缩放。它常把值轴从 0 起画，量本身只在 0.69~0.73
    晃的话就被压成一条平线，什么都看不出来。范围必须由数据算出来钉死。
    """
    import math
    vals = [v for v in vals if v is not None]
    if not vals:
        return None
    lo, hi = min(vals), max(vals)
    if hi == lo:
        d = abs(hi) * 0.05 or 1.0
        lo, hi = lo - d, hi + d
    m = (hi - lo) * pad
    lo, hi = lo - m, hi + m
    step = nice_step(hi - lo)
    return math.floor(lo / step) * step, math.ceil(hi / step) * step, step


def apply_y(chart, bounds):
    if not bounds:
        return
    chart.y_axis.scaling.min, chart.y_axis.scaling.max = bounds[0], bounds[1]
    chart.y_axis.majorUnit = bounds[2]


def legend_bottom(chart):
    if chart.legend is not None:
        chart.legend.position = "b"
        chart.legend.overlay = False


def blank_policy(chart, data_sheet_hidden=False):
    """断点留空 + 数据页隐藏了也照画。

    ★ 属性名必须是 openpyxl 的 display_blanks / visible_cells_only。
      写成 OOXML 里的 dispBlanksAs / plotVisOnly **不报错**，只是给对象挂了个
      没人读的属性：dispBlanksAs 缺省是 "zero"，于是没测的格子被当 0 画进图，
      曲线在测点之间扎到零成一串尖刺。
    """
    chart.display_blanks = "gap"
    if data_sheet_hidden:
        chart.visible_cells_only = False


def style_series(series, color, symbol, line=True, dash=None, size=6):
    from openpyxl.chart.marker import Marker
    from openpyxl.chart.shapes import GraphicalProperties
    from openpyxl.drawing.line import LineProperties
    gp = GraphicalProperties()
    gp.line = LineProperties(noFill=True) if not line else \
        LineProperties(solidFill=color, w=20000, prstDash=dash)
    series.graphicalProperties = gp
    if symbol:
        m = Marker(symbol=symbol, size=size)
        m.graphicalProperties = GraphicalProperties(solidFill=color)
        series.marker = m
    else:
        series.marker = Marker(symbol="none")
    series.smooth = False
