#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
xlsx_formula_cache.py — 给 openpyxl 写出的公式格补「缓存值」

★ 这个模块存在的唯一理由（踩过的坑）
    openpyxl 写公式格时，写出来的是：

        <c r="L15" s="15"><f>IF(COUNT(...)=0,"",MEDIAN(...))</f><v></v></c>

    注意那个 **空的** `<v></v>`——它是「缓存值」字段，等于在文件里明确写着
    "这个公式算出来的结果就是空"。后果：

      · 自己机器上打开看着正常——Excel 因为 fullCalcOnLoad 做了一次全量重算；
        （fullCalcOnLoad 本来就是 openpyxl 的默认值，所以它**不是**解药）
      · 把文件发给别人，格子一片空白，虽然公式确实在里面。受保护视图
        （邮件/IM 收到的文件默认就是）、手动计算模式、网页版/WPS 之类的
        阅读器，都会直接信那个空缓存值；
      · 在格子里敲一次回车强制该格重算，值才出现。

    修法两条，本模块都做：
      ① 知道结果的 → 把真值写进 `<v>`。这样**任何环境**打开都直接看到数，
         压根不依赖重算。公式还在，原始数据改了照样能重算。
      ② 不知道结果的（例如判定列要等人填 Spec）→ 把空 `<v></v>` **删掉**。
         删掉反而对：Excel 见一个公式格没有缓存值，就必须自己算。

用法
    from xlsx_formula_cache import FormulaCache
    vc = FormulaCache()
    ...
    ws.cell(r, c).value = "=MIN(A1:A9)"
    vc.remember(ws, r, c, 3.14)          # 边写公式边把 Python 算好的结果记下
    ...
    wb.save(path)
    n_fill, n_strip = vc.inject(path)    # 必须在 save 之后

只用标准库（zipfile / re / xml.etree），不依赖 openpyxl 的内部结构。
"""

import re
import shutil
import zipfile
from xml.etree import ElementTree as ET

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
RNS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
RELNS = "{http://schemas.openxmlformats.org/package/2006/relationships}"

_CELL_RE = re.compile(rb"<c\b[^>]*?r=\"([A-Z]+[0-9]+)\"[^>]*>(.*?)</c>", re.S)


def _col_letter(idx):
    s = ""
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        s = chr(65 + rem) + s
    return s


def _fmt(v):
    """写进 <v> 的数字文本。round 到 12 位小数，收掉浮点减法的噪声尾巴
    （max-lock 之类会算出 -0.00046999999999997）；只动第 12 位以后，
    肉眼和 Excel 重算的结果都看不出差别。"""
    if isinstance(v, float):
        v = round(v, 12)
        if v == int(v) and abs(v) < 1e15:
            v = int(v)
    return str(v) if isinstance(v, int) else repr(v)


class Formula(object):
    """(公式, Python 算好的值) 一对。

    让写格子的地方只管把它交给 put()，由 put() 一处负责
    「公式写进格子、值记进缓存」，不必每个调用点都记一次。
    """

    __slots__ = ("formula", "value")

    def __init__(self, formula, value):
        self.formula, self.value = formula, value

    def __str__(self):
        return str(self.formula)


class FormulaCache(object):
    def __init__(self):
        self.d = {}

    def remember(self, ws, row, col, value):
        """记下某个公式格的计算结果。非数值（None / 字符串 / 布尔）直接忽略——
        字符串结果没法安全地当数字缓存，交给 Excel 自己算更稳。"""
        if isinstance(value, bool) or not isinstance(value, (int, float)):
            return
        self.d.setdefault(ws.title, {})[_col_letter(col) + str(row)] = value

    # ------------------------------------------------------------------
    def inject(self, path):
        """存盘后把缓存值补进 xlsx。返回 (补了几个, 清掉几个空缓存)。"""
        with zipfile.ZipFile(path) as z:
            parts = {n: z.read(n) for n in z.namelist()}

        sheet_part = self._sheet_parts(parts)
        n_fill = n_strip = 0

        for title, part in sheet_part.items():
            if not part:
                continue
            vals = self.d.get(title, {})

            def repl(m, vals=vals):
                nonlocal n_fill, n_strip
                inner = m.group(2)
                if b"<f" not in inner:          # 不是公式格，别碰
                    return m.group(0)
                v = vals.get(m.group(1).decode())
                if v is None:
                    if b"<v></v>" in inner:     # 不知道结果 -> 清掉空缓存
                        n_strip += 1
                        return m.group(0).replace(b"<v></v>", b"")
                    return m.group(0)
                new = ("<v>%s</v>" % _fmt(v)).encode()
                n_fill += 1
                if b"<v></v>" in inner:
                    return m.group(0).replace(b"<v></v>", new)
                if b"<v>" in inner:             # 已经有缓存值了，不动
                    return m.group(0)
                return m.group(0)[: -len(b"</c>")] + new + b"</c>"

            parts[part] = _CELL_RE.sub(repl, parts[part])

        tmp = path + ".tmp"
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
            for n, d in parts.items():
                z.writestr(n, d)
        shutil.move(tmp, path)
        return n_fill, n_strip

    @staticmethod
    def _sheet_parts(parts):
        """工作表名 -> zip 里对应的 sheetN.xml 路径。

        不能假设 sheet 名的顺序就是 sheetN.xml 的编号顺序，得老实走
        workbook.xml 的 r:id -> workbook.xml.rels 的 Target。
        """
        rels = {}
        rx = parts.get("xl/_rels/workbook.xml.rels")
        if rx:
            for r in ET.fromstring(rx):
                if r.tag == RELNS + "Relationship":
                    rels[r.get("Id")] = r.get("Target") or ""
        out = {}
        wbx = parts.get("xl/workbook.xml")
        if not wbx:
            return out
        for sh in ET.fromstring(wbx).iter(NS + "sheet"):
            tgt = rels.get(sh.get(RNS + "id"), "")
            if not tgt:
                continue
            # Target 可能是包内绝对路径 /xl/worksheets/sheetN.xml，
            # 也可能是相对 xl/ 的 worksheets/sheetN.xml——两种都得认
            p = tgt.lstrip("/") if tgt.startswith("/") else "xl/" + tgt
            out[sh.get("name")] = p if p in parts else None
        return out
