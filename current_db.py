#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
current_db.py — 电流数据库（v4：全模式单文件多温度 + 仿真 tier/stage 过滤 + 可读汇总簿）

把两类数据统一进一个 SQLite 库，并导出对比/汇总 Excel：
  1) 仿真长表：某工作簿的 Current_data 页（ID/Module/Trim/Mode/simulation/Tier/Current/Unit；
     simulation=pre/post，Tier 为电流档位——config.sim_tier 选定与实测同档的数据参与对比）
  2) 实测结果：Result*.xlsx。两种形态自动识别：
     a. 全模式单文件（可含多温度段）：按行分段——`*_sigN` 签名行 / Init 行 NO.=模式名 开段，
        SET_TEMP·Chamber 行闭段；同段温度取自 Temperature 列 → 每 (模式,温度) 一个 run
     b. 旧单模式文件：整表一段（Init 行 NO. 给模式名，文件夹名兜底）
     模式名与仿真表 Mode 自动对齐：大小写/下划线无关、UNSYNC≡NOSYNC、尾部裸 SYNC 可省
     （BT2GRX_unSync≡BT_2G_RX_noSYNC；BT2GRX_sync≡BT_2G_RX）；config.mode_map 可强制指定。

实测解析规则（在每个模式段内独立执行）：
  - 一个序列从 Init 行开始；基线 = 第一个 OFF 行之前最后一行的电流（通常是最后一个 Lock_step）
  - 模块电流 = 上一行电流 - 本行电流（逐级关断做差），统一换算成 uA
  - 第二个及以后的 Init 段 = 锁定复验，忽略（原始行仍入库审计；全模式文件按段分割后天然不会触发）
  - SET_TEMP / chamber 行只控温箱，不参与做差
  - NO. 列多个编号（如 "45,46"）= 该步同时关断的一组模块，按组对比（仿真侧求和）
  - LDO 归并（config.ldo_reparent，如 28->26）：子模块不在被测 LDO 下，
    其实测 delta 并入父模块组；对比时仿真侧同样求和（meas(26)+meas(28) vs sim(26)+sim(28)）
  - NO. 列非数字标签（如 "DCO5G"）：config.label_groups 映射到仿真模块 ID；
    值可以是 ID 列表（同模式），也可以是 {"mode": "CK_ADPLL_DCO2G", "ids": "*"}
    （跨模式取该仿真 Mode 的全部/指定模块合计）

用法（在数据所在机器上）：
  python current_db.py build   --root D:\\Excel --chip C1
    首次运行会在 root 下生成 current_config.json（sim_tier/模式映射/LDO 归并等都在里面改），
    并输出 current.db + Current_compare_pivot.xlsx
  python current_db.py summary --db current.db --out 各模式功耗表.xlsx
    人直接读的汇总簿：总览矩阵(模块×模式×温度+仿真对比) / 温度趋势图 / 对比明细

  也可分步：
  python current_db.py ingest-sim --db current.db --xlsx Current_all_mode_v2.xlsx
  python current_db.py ingest-run --db current.db --xlsx <Result文件> --mode BT_5G_TX --chip C1
  python current_db.py export     --db current.db --out out.xlsx [--all-runs]

依赖：openpyxl（与本仓库其余工具一致，无其他第三方依赖）
"""
import argparse
import datetime
import fnmatch
import json
import os
import re
import sqlite3
import sys

import openpyxl
from openpyxl.chart import LineChart, Reference
from openpyxl.chart.text import RichText, Text
from openpyxl.chart.title import Title
from openpyxl.drawing.text import (CharacterProperties, Font as DrawFont, Paragraph,
                                   ParagraphProperties, RegularTextRun)
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------- 常量/工具

UNIT_TO_UA = {"ua": 1.0, "µa": 1.0, "μa": 1.0, "ma": 1000.0, "na": 0.001, "a": 1e6, "": 1.0}

DEFAULT_CONFIG = {
    "_说明": {
        "sim_workbook": "仿真长表所在工作簿（相对 root 或绝对路径）；null=自动找 Current_all_mode*.xlsx",
        "sim_sheet": "仿真长表的 tab 名",
        "result_glob": "每个模式文件夹里匹配实测文件的通配符",
        "result_sheet": "实测数据所在 tab；null=自动扫描含 NO./Current 表头的第一个 tab",
        "skip_dirs": "扫描 root 子目录时跳过的文件夹",
        "mode_map": "文件夹名 -> 仿真表 Mode 名 的映射（同名可省略）",
        "ldo_reparent": "子模块ID -> 父模块ID：子模块不在被测 LDO 下，其实测 delta 并入父模块组",
        "ldo_reparent_sim_add_child": "归并时仿真侧是否把子模块电流也加进对比和（子模块电流不在被测轨上时应为 false）",
        "label_groups": "NO. 列非数字标签 -> 仿真模块ID列表如 {\"DCO5G\": [21]}，"
                        "或跨模式 {\"DCO2G\": {\"mode\": \"CK_ADPLL_DCO2G\", \"ids\": \"*\"}}",
        "exclude_globs": "扫描时按文件名跳过的通配符（本工具自己的输出必须在内，防自吞）",
        "sim_tier": "参与对比的仿真电流档位（如 Tier2，与实测一致）；空=不过滤（多档共存会重复求和！）",
        "sim_temp_note": "仿真数据的温度/corner 标注，只用于表头展示（如 55C/TT/0.9V）",
        "delta_flag_pct": "汇总簿里 |偏差%| 超过该值标红（默认 20）",
        "delta_ref_temp": "汇总簿偏差列用哪个实测温度对仿真（null=自动取离 55℃ 最近的温度点）",
        "mode_freq": "汇总簿测试频率条件行：模式名 -> 显示文本；不在表里的按名字推断"
                     "（含 2G -> 2.5GHz，含 5G -> 5.8GHz）",
    },
    "sim_workbook": None,
    "sim_sheet": "Current_data",
    "result_glob": "Result*.xlsx",
    "result_sheet": None,
    "skip_dirs": ["Simulation", "自动化"],
    "mode_map": {},
    "ldo_reparent": {"8": "6", "28": "26"},
    "ldo_reparent_sim_add_child": False,
    "label_groups": {},
    "exclude_globs": ["Current_compare_pivot*.xlsx", "probe_dump*", "*功耗表*.xlsx"],
    "sim_tier": "Tier2",
    "sim_temp_note": "55C",
    "delta_flag_pct": 20,
    "delta_ref_temp": None,
    "mode_freq": {},
}


def norm(v):
    return str(v).strip().lower() if v is not None else ""


def cell(row, idx):
    """read_only 模式下行元组可能比表头短，安全取值。"""
    if idx is None or idx >= len(row):
        return None
    return row[idx]


def as_float(v):
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(",", "")
    try:
        return float(s)
    except ValueError:
        return None


def parse_ids(no_raw):
    """NO. 单元格 -> 模块ID列表；解析失败返回 None（说明是标签）。"""
    if no_raw is None:
        return None
    if isinstance(no_raw, (int, float)) and not isinstance(no_raw, bool):
        return [int(no_raw)] if float(no_raw).is_integer() else None
    parts = re.split(r"[,，、;；\s]+", str(no_raw).strip())
    ids = []
    for p in parts:
        if not p:
            continue
        if not re.fullmatch(r"\d+", p):
            return None
        ids.append(int(p))
    return ids or None


def canon_mode(name):
    """模式名规范化键：大小写/下划线等分隔无关；UNSYNC≡NOSYNC；尾部裸 SYNC（默认态）可省。
    例：BT2GRX_unSync 与 BT_2G_RX_noSYNC 同键；BT2GRX_sync 与 BT_2G_RX 同键。"""
    s = re.sub(r"[^0-9A-Za-z]", "", str(name or "")).upper()
    s = s.replace("UNSYNC", "NOSYNC")
    return re.sub(r"(?<!NO)SYNC$", "", s)


def resolve_mode(label, sim_modes, mode_map, folder=None):
    """实测段标签 -> 仿真表 Mode 名。优先 config.mode_map（键可以是段标签或文件夹名），
    其次 canon 规范化唯一匹配。返回 (resolved, how)，how ∈ config/auto/ambig/none。"""
    mode_map = mode_map or {}
    for key in (label, folder):
        if key and key in mode_map:
            return mode_map[key], "config"
    c = canon_mode(label)
    hits = sorted({m for m in sim_modes or () if canon_mode(m) == c})
    if len(hits) == 1:
        return hits[0], "auto"
    return label, ("ambig" if hits else "none")


def now_iso():
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


# ---------------------------------------------------------------- 配置

def load_config(root, path=None):
    cfg_path = path or os.path.join(root, "current_config.json")
    if os.path.exists(cfg_path):
        with open(cfg_path, "r", encoding="utf-8") as f:
            cfg = json.load(f)
        merged = dict(DEFAULT_CONFIG)
        merged.update(cfg)
        return merged, cfg_path, False
    with open(cfg_path, "w", encoding="utf-8") as f:
        json.dump(DEFAULT_CONFIG, f, ensure_ascii=False, indent=2)
    return dict(DEFAULT_CONFIG), cfg_path, True


# ---------------------------------------------------------------- 数据库

SCHEMA = """
CREATE TABLE IF NOT EXISTS runs(
    run_id INTEGER PRIMARY KEY,
    mode TEXT, chip TEXT, temp_c REAL,
    src_file TEXT, run_ts TEXT, ingested_ts TEXT, mode_raw TEXT);
CREATE TABLE IF NOT EXISTS meas_raw(
    id INTEGER PRIMARY KEY,
    run_id INTEGER, row_idx INTEGER, seq_idx INTEGER, kind TEXT,
    no_raw TEXT, mode_label TEXT, current_ma REAL, delta_ma REAL,
    temp_c REAL, note TEXT);
CREATE TABLE IF NOT EXISTS meas_module(
    id INTEGER PRIMARY KEY,
    run_id INTEGER, step_order INTEGER,
    group_disp TEXT, step_name TEXT,
    module_ids TEXT, sim_ids TEXT,
    current_ua REAL, note TEXT, sim_mode TEXT);
CREATE TABLE IF NOT EXISTS sim_current(
    id INTEGER PRIMARY KEY,
    module_id INTEGER, module_name TEXT, trim TEXT, mode TEXT,
    stage TEXT, tier TEXT, current_ua REAL, unit_raw TEXT, src_file TEXT);
"""


def open_db(path):
    conn = sqlite3.connect(path)
    conn.executescript(SCHEMA)
    for alter in ("ALTER TABLE runs ADD COLUMN mode_raw TEXT",
                  "ALTER TABLE meas_module ADD COLUMN sim_mode TEXT"):
        try:
            conn.execute(alter)  # v4 前建的库补列
        except sqlite3.OperationalError:
            pass
    return conn


# ---------------------------------------------------------------- 仿真表导入

def match_sim_header(row):
    """仿真长表表头：ID + Mode + Current*，且有 simulation/Unit/Tier 佐证。"""
    names = [norm(c) for c in row]
    return ("id" in names and "mode" in names
            and any(n.startswith("current") for n in names)
            and (any(n.startswith("simulation") for n in names)
                 or "unit" in names))  # 注意: 不认 tier——本工具导出的 Sim_long 页有 tier 列


def find_sim_header(ws):
    for i, row in enumerate(ws.iter_rows(min_row=1, max_row=30, values_only=True), 1):
        names = [norm(c) for c in row]
        if "id" in names and "mode" in names and any(n.startswith("current") for n in names):
            cols = {}
            for j, n in enumerate(names):
                if n == "id":
                    cols["id"] = j
                elif n == "module":
                    cols["module"] = j
                elif n == "trim":
                    cols["trim"] = j
                elif n == "mode":
                    cols["mode"] = j
                elif n.startswith("simulation") or n == "sim":
                    cols["sim"] = j
                elif n == "tier":
                    cols["tier"] = j
                elif n.startswith("current"):
                    cols["current"] = j
                elif n == "unit":
                    cols["unit"] = j
            return i, cols
    return None, None


def ingest_sim(conn, xlsx, sheet_name):
    wb = openpyxl.load_workbook(xlsx, read_only=True, data_only=True)
    try:
        ws = None
        if sheet_name and sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
        else:
            for sn in wb.sheetnames:
                if norm(sn) == norm(sheet_name or "current_data"):
                    ws = wb[sn]
                    break
        if ws is None:  # 按名字没找到 -> 按表头内容扫
            for sn in wb.sheetnames:
                cand = wb[sn]
                if any(match_sim_header(r) for r in
                       cand.iter_rows(min_row=1, max_row=30, values_only=True)):
                    ws = cand
                    break
        if ws is None:
            raise SystemExit(f"[错误] {os.path.basename(xlsx)} 里找不到仿真 tab（按名 {sheet_name!r} "
                             f"或按表头 ID/Mode/Current 都没命中），现有 tab: {wb.sheetnames}")
        hdr, cols = find_sim_header(ws)
        if hdr is None:
            raise SystemExit(f"[错误] 仿真 tab {ws.title!r} 找不到表头行（需含 ID/Mode/Current 列）")
        conn.execute("DELETE FROM sim_current WHERE src_file=?", (os.path.abspath(xlsx),))
        n = 0
        for row in ws.iter_rows(min_row=hdr + 1, values_only=True):
            cur = as_float(cell(row, cols.get("current")))
            mid = cell(row, cols.get("id"))
            mode = cell(row, cols.get("mode"))
            if cur is None or mode is None:
                continue
            unit = norm(cell(row, cols.get("unit")))
            factor = UNIT_TO_UA.get(unit)
            if factor is None:
                factor = 1.0
            stage_raw = norm(cell(row, cols.get("sim")))
            stage = "pre" if stage_raw.startswith("pre") else ("post" if stage_raw else "")
            try:
                mid_i = int(mid)
            except (TypeError, ValueError):
                mid_i = None
            conn.execute(
                "INSERT INTO sim_current(module_id,module_name,trim,mode,stage,tier,current_ua,unit_raw,src_file)"
                " VALUES(?,?,?,?,?,?,?,?,?)",
                (mid_i, str(cell(row, cols.get("module")) or ""),
                 str(cell(row, cols.get("trim")) or ""), str(mode).strip(),
                 stage, str(cell(row, cols.get("tier")) or ""),
                 cur * factor, unit, os.path.abspath(xlsx)))
            n += 1
        conn.commit()
        return n
    finally:
        wb.close()


# ---------------------------------------------------------------- 实测表解析

def match_result_header(row):
    """实测表表头。带优先级：Current_mA(带单位) > Current(裸开关列)；Temperature* > temp
    （防撞 Vtemp）。返回列映射 dict（含 unit）或 None。"""
    no_col = mode_col = None
    cur_exact = cur_bare = temp_exact = temp_bare = None
    for j, c in enumerate(row):
        n = norm(c)
        if n in ("no.", "no", "no．") and no_col is None:
            no_col = j
        elif n == "mode" and mode_col is None:
            mode_col = j
        elif re.fullmatch(r"current[_\s]*[munµμ]?a", n) and cur_exact is None:
            cur_exact = j
        elif n == "current" and cur_bare is None:
            cur_bare = j
        elif n.startswith("temperature") and temp_exact is None:
            temp_exact = j
        elif n == "temp" and temp_bare is None:
            temp_bare = j
    cur_col = cur_exact if cur_exact is not None else cur_bare
    temp_col = temp_exact if temp_exact is not None else temp_bare
    if no_col is None or cur_col is None:
        return None
    if mode_col is None:
        mode_col = no_col + 1
    unit = "ma"
    m = re.fullmatch(r"current[_\s]*([munµμ]?a)", norm(row[cur_col]))
    if m and m.group(1):
        unit = m.group(1)
    return {"no": no_col, "mode": mode_col, "cur": cur_col, "temp": temp_col, "unit": unit}


def find_result_sheet(wb, sheet_name):
    """返回 (worksheet, 表头行号, 列映射)。按表头名定位，不按列字母。"""
    names = [sheet_name] if sheet_name else wb.sheetnames
    for sn in names:
        if sn not in wb.sheetnames:
            continue
        ws = wb[sn]
        for i, row in enumerate(ws.iter_rows(min_row=1, max_row=30, values_only=True), 1):
            cols = match_result_header(row)
            if cols is not None:
                return ws, i, cols
    return None, None, None


def read_raw_rows(ws, hdr, cols):
    """表体 -> [(行号, no_raw, label, 电流float, 温度float)]，跳过全空行。"""
    raw = []
    for i, row in enumerate(ws.iter_rows(min_row=hdr + 1, values_only=True), hdr + 1):
        no_raw = cell(row, cols["no"])
        label = cell(row, cols["mode"])
        cur = as_float(cell(row, cols["cur"]))
        temp = as_float(cell(row, cols["temp"])) if cols["temp"] is not None else None
        if no_raw is None and label is None and cur is None:
            continue
        raw.append((i, no_raw, label, cur, temp))
    return raw


SIG_RE = re.compile(r"(.+)_sig\d+$")


def split_allmode(raw):
    """全模式单文件按行分段 -> [{mode, temp, raw:[...]}]；整表无边界时返回 []。
    开段：`*_sigN` 签名行（最可靠）/ Init 行 NO.=模式名（兜底，兼容旧单模式文件）。
    闭段：SET_TEMP 设温行、chamber 行（只控温箱，不入任何段）。
    相邻且 (canon模式,温度) 相同的段合并——签名行用 unSync、用户原行用 noSYNC 之类的
    命名分裂在这里收口，段名取带 Init 原行那段的写法（与仿真表命名一致的那个）。"""
    segs, cur_seg = [], None
    for rec in raw:
        _i, no_raw, label, _cur, _temp = rec
        ls = str(label).strip() if label is not None else ""
        ln, nn = norm(label), norm(no_raw)
        if ln.startswith("set_temp") or nn.startswith("set_temp") \
                or "chamber" in ln or "chamber" in nn:
            cur_seg = None
            continue
        m = SIG_RE.fullmatch(ls)
        if m:
            if cur_seg is None or canon_mode(cur_seg["mode"]) != canon_mode(m.group(1)):
                cur_seg = {"mode": m.group(1), "raw": []}
                segs.append(cur_seg)
        elif ln.startswith("init") and isinstance(no_raw, str) and no_raw.strip() \
                and not re.fullmatch(r"[\d ,，、;；]+", no_raw.strip()):
            name = no_raw.strip()
            if cur_seg is None or canon_mode(cur_seg["mode"]) != canon_mode(name):
                cur_seg = {"mode": name, "raw": []}
                segs.append(cur_seg)
            else:
                cur_seg["mode"] = name  # 同段：签名行段名让位给 Init 原行段名
        if cur_seg is not None:
            cur_seg["raw"].append(rec)
    out = []
    for s in segs:
        temps = [t for (_i, _n, _l, _c, t) in s["raw"] if t is not None]
        s["temp"] = temps[0] if temps else None
        if out and canon_mode(out[-1]["mode"]) == canon_mode(s["mode"]) \
                and out[-1]["temp"] == s["temp"]:
            out[-1]["raw"].extend(s["raw"])
            out[-1]["mode"] = s["mode"]
        else:
            out.append(s)
    return out


def classify_rows(ws, hdr, cols):
    """兼容入口：整表当一个序列。返回 (rows, temp)。"""
    factor_to_ma = UNIT_TO_UA.get(cols["unit"], 1000.0) / 1000.0  # 原始单位 -> mA
    return classify_raw(read_raw_rows(ws, hdr, cols), factor_to_ma)


def classify_raw(raw, factor_to_ma):
    """逐行分类并做差。返回 (rows, temp)。
    rows: dict(row_idx, no_raw, label, cur_ma, delta_ma, temp, seq, kind)
    两遍扫描：先看有没有显式 Init 行——有的话序列只从 Init 行开始，
    Init 之前的带电流行（其他测试项/签名行）留在 seq=0 不参与做差；
    全表都没有 Init 行时，才把第一个带电流行当作序列起点。"""
    has_init = any(norm(label).startswith("init") for _i, _n, label, _c, _t in raw)

    out = []
    seq = 0
    prev_cur = None
    temp_first = None
    for i, no_raw, label, cur, temp in raw:
        ln, nn = norm(label), norm(no_raw)
        if "chamber" in ln or "chamber" in nn:
            kind = "chamber"
        elif ln.startswith("init"):
            kind = "init"
            seq += 1
        elif "lock" in ln:
            kind = "lock"
        elif ln.startswith("off"):
            kind = "off"
        else:
            kind = "other"
        if seq == 0 and not has_init and cur is not None:
            seq = 1  # 整表无显式 Init 行时，第一段视为正式测量
            if kind == "other":
                kind = "init"
        delta = None
        if seq == 1 and kind in ("lock", "off", "other") and cur is not None and prev_cur is not None:
            delta = prev_cur * factor_to_ma - cur * factor_to_ma
        if seq == 1 and cur is not None and kind != "chamber":
            prev_cur = cur
        if temp is not None and temp_first is None and seq >= 1:
            temp_first = temp
        if kind == "other" and cur is None and seq == 0:
            continue  # Init 之前的测试计划行，不入库
        out.append(dict(row_idx=i, no_raw=no_raw, label=label,
                        cur_ma=(cur * factor_to_ma) if cur is not None else None,
                        delta_ma=delta, temp=temp, seq=seq, kind=kind))
    return out, temp_first


def build_groups(rows, config):
    """从 seq==1 的 OFF 行生成模块组，套用 LDO 归并。返回 (groups, absorbed_notes)。"""
    reparent = {}
    for c, p in (config.get("ldo_reparent") or {}).items():
        try:
            reparent[int(c)] = int(p)
        except (TypeError, ValueError):
            pass
    label_groups = {str(k): list(v) for k, v in (config.get("label_groups") or {}).items()}

    steps = []
    for r in rows:
        if r["seq"] != 1 or r["kind"] != "off" or r["delta_ma"] is None:
            continue
        ids = parse_ids(r["no_raw"])
        disp = ",".join(str(i) for i in ids) if ids else str(r["no_raw"]).strip()
        step_name = re.sub(r"(?i)^off\s*", "", str(r["label"] or "")).strip()
        note = ""
        sim_ids = list(ids) if ids else None
        sim_mode = None  # 该步仿真值来自其他仿真 Mode 时（如 DCO 标签对 CK_ADPLL_*）
        if ids is None:
            mapped = label_groups.get(disp)
            if isinstance(mapped, dict):
                sim_mode = str(mapped.get("mode") or "").strip() or None
                v = mapped.get("ids")
                sim_ids = ["*"] if v in (None, "*", ["*"]) else [int(x) for x in v]
                note = (f"标签 {disp} 按 config.label_groups 映射到仿真"
                        f"{' Mode ' + sim_mode if sim_mode else ''} ID {sim_ids}")
            elif mapped:
                sim_ids = [int(x) for x in mapped]
                note = f"标签 {disp} 按 config.label_groups 映射到仿真 ID {sim_ids}"
            else:
                note = "标签未映射仿真模块（可在 current_config.json 的 label_groups 补充）"
        steps.append(dict(row_idx=r["row_idx"], ids=ids, sim_ids=sim_ids, sim_mode=sim_mode,
                          disp=disp, step_name=step_name, delta_ua=r["delta_ma"] * 1000.0,
                          note=note))

    # LDO 归并：单独成步的子模块，实测 delta 与仿真 ID 都并入父模块所在组
    absorbed = {}  # 子步 row_idx -> 父组 disp
    by_single_id = {s["ids"][0]: s for s in steps if s["ids"] and len(s["ids"]) == 1}
    for child, parent in reparent.items():
        child_step = by_single_id.get(child)
        parent_step = next((s for s in steps if s["ids"] and parent in s["ids"]), None)
        if child_step is None or parent_step is None or child_step is parent_step:
            if child_step is not None and parent_step is None:
                child_step["note"] = (child_step["note"] + "；" if child_step["note"] else "") + \
                    f"模块{child}不在被测LDO下（父模块{parent}本次未测，未归并）"
            continue
        parent_step["delta_ua"] += child_step["delta_ua"]
        add_child_sim = bool(config.get("ldo_reparent_sim_add_child", False))
        if add_child_sim:
            parent_step["sim_ids"] = (parent_step["sim_ids"] or []) + [child]
        # 编号列保持父模块本来的编号（用户定：6 就是 6），归并信息只进备注列
        parent_step["note"] = (parent_step["note"] + "；" if parent_step["note"] else "") + \
            f"含模块{child}的实测delta（{child}不在被测LDO下，仿真侧{'已并入' if add_child_sim else '不计'}{child}）"
        absorbed[child_step["row_idx"]] = parent_step["disp"]
    steps = [s for s in steps if s["row_idx"] not in absorbed]
    for order, s in enumerate(steps, 1):
        s["order"] = order
    return steps, absorbed


def _run_ts_of(xlsx):
    m = re.search(r"(\d{4}-\d{2}-\d{2})-(\d{2})-(\d{2})-(\d{2})", os.path.basename(xlsx))
    if m:
        return f"{m.group(1)} {m.group(2)}:{m.group(3)}:{m.group(4)}"
    return datetime.datetime.fromtimestamp(os.path.getmtime(xlsx)).strftime("%Y-%m-%d %H:%M:%S")


def _delete_runs_of(conn, src, chip):
    for (rid,) in conn.execute("SELECT run_id FROM runs WHERE src_file=? AND chip=?",
                               (src, chip)).fetchall():
        conn.execute("DELETE FROM meas_raw WHERE run_id=?", (rid,))
        conn.execute("DELETE FROM meas_module WHERE run_id=?", (rid,))
        conn.execute("DELETE FROM runs WHERE run_id=?", (rid,))


def _insert_run(conn, src, mode, mode_raw, chip, temp, rows, steps, absorbed, run_ts):
    cur = conn.execute(
        "INSERT INTO runs(mode,chip,temp_c,src_file,run_ts,ingested_ts,mode_raw)"
        " VALUES(?,?,?,?,?,?,?)",
        (mode, chip, temp, src, run_ts, now_iso(), mode_raw))
    run_id = cur.lastrowid
    for r in rows:
        note = ""
        if r["seq"] >= 2:
            note = "锁定复验段，忽略"
        elif r["row_idx"] in absorbed:
            note = f"并入组 {absorbed[r['row_idx']]}"
        conn.execute(
            "INSERT INTO meas_raw(run_id,row_idx,seq_idx,kind,no_raw,mode_label,current_ma,delta_ma,temp_c,note)"
            " VALUES(?,?,?,?,?,?,?,?,?,?)",
            (run_id, r["row_idx"], r["seq"], r["kind"],
             str(r["no_raw"]) if r["no_raw"] is not None else None,
             str(r["label"]) if r["label"] is not None else None,
             r["cur_ma"], r["delta_ma"], r["temp"], note))
    for s in steps:
        conn.execute(
            "INSERT INTO meas_module(run_id,step_order,group_disp,step_name,module_ids,sim_ids,"
            "current_ua,note,sim_mode) VALUES(?,?,?,?,?,?,?,?,?)",
            (run_id, s["order"], s["disp"], s["step_name"],
             json.dumps(s["ids"]) if s["ids"] else None,
             json.dumps(s["sim_ids"]) if s["sim_ids"] else None,
             s["delta_ua"], s["note"], s.get("sim_mode")))
    return run_id


def ingest_run(conn, xlsx, mode, chip, config, sheet_name=None):
    """单模式显式入库（ingest-run 子命令）：整表当一个序列，模式名由调用者给。"""
    wb = openpyxl.load_workbook(xlsx, read_only=True, data_only=True)
    try:
        ws, hdr, cols = find_result_sheet(wb, sheet_name or config.get("result_sheet"))
        if ws is None:
            raise SystemExit(f"[错误] {os.path.basename(xlsx)} 里找不到含 NO./Current 表头的 tab")
        rows, temp = classify_rows(ws, hdr, cols)
        steps, absorbed = build_groups(rows, config)
        src = os.path.abspath(xlsx)
        run_ts = _run_ts_of(xlsx)
        _delete_runs_of(conn, src, chip)
        run_id = _insert_run(conn, src, mode, mode, chip, temp, rows, steps, absorbed, run_ts)
        conn.commit()
        return run_id, len(steps), temp, run_ts
    finally:
        wb.close()


def ingest_result_file(conn, xlsx, chip, config, sim_modes, folder_mode=None, sheet_name=None):
    """一个 Result 文件 -> 若干 run（全模式单文件按 (模式,温度) 分段；旧单模式文件=1 段，
    模式名取 Init 行 NO.，退无可退才用文件夹名）。
    返回 [(run_id, mode, mode_raw, how, temp, n_steps)]。"""
    wb = openpyxl.load_workbook(xlsx, read_only=True, data_only=True)
    try:
        ws, hdr, cols = find_result_sheet(wb, sheet_name or config.get("result_sheet"))
        if ws is None:
            raise SystemExit(f"[错误] {os.path.basename(xlsx)} 里找不到含 NO./Current 表头的 tab")
        raw = read_raw_rows(ws, hdr, cols)
    finally:
        wb.close()
    factor_to_ma = UNIT_TO_UA.get(cols["unit"], 1000.0) / 1000.0
    return _ingest_raw(conn, raw, factor_to_ma, os.path.abspath(xlsx), _run_ts_of(xlsx),
                       chip, config, sim_modes, folder_mode)


def _ingest_raw(conn, raw, factor_to_ma, src, run_ts, chip, config, sim_modes, folder_mode=None):
    segs = split_allmode(raw)
    if not segs:  # 无任何段边界：整表一个序列，文件夹名当模式
        segs = [{"mode": folder_mode or "?", "raw": raw, "temp": None}]
    # 旧单模式目录（整文件一段）：文件夹名是操作者意图；段内 Init 标签见过模板复制错名
    # （DCO2G 文件夹里写着 DCO5G），此时文件夹名优先，映射表里标出来。全模式多段文件只信段标签。
    single_legacy = (len(segs) == 1 and folder_mode
                     and canon_mode(folder_mode) != canon_mode(segs[0]["mode"]))
    _delete_runs_of(conn, src, chip)
    out = []
    mode_map = config.get("mode_map") or {}
    for s in segs:
        rows, temp0 = classify_raw(s["raw"], factor_to_ma)
        temp = s["temp"] if s.get("temp") is not None else temp0
        steps, absorbed = build_groups(rows, config)
        if single_legacy:
            mode, _ = resolve_mode(folder_mode, sim_modes, mode_map)
            how = "folder"
        else:
            mode, how = resolve_mode(s["mode"], sim_modes, mode_map, folder=folder_mode)
        run_id = _insert_run(conn, src, mode, s["mode"], chip, temp, rows, steps, absorbed, run_ts)
        out.append((run_id, mode, s["mode"], how, temp, len(steps)))
    conn.commit()
    return out


def ingest_probe_json(conn, json_path, chip, config):
    """probe_allmode_result.py --json 的产物入库（黄区只带 JSON 回来时的开发机路径）。
    行流按原行号重排后走与 xlsx 完全相同的分段/分类管线。"""
    with open(json_path, "r", encoding="utf-8") as f:
        d = json.load(f)
    recs = []
    for seg in (d.get("segments") or []):
        for r in seg.get("rows") or []:
            recs.append((r.get("row"), r.get("no"), r.get("label"),
                         as_float(r.get("current")), as_float(r.get("temp"))))
    for r in (d.get("orphans") or []):
        recs.append((r.get("row"), r.get("no"), r.get("label"),
                     as_float(r.get("current")), as_float(r.get("temp"))))
    recs.sort(key=lambda x: (x[0] if x[0] is not None else 0))
    unit = "ma"
    kc = d.get("key_cols_1based") or {}
    if isinstance(kc.get("unit"), str):
        unit = kc["unit"]
    factor_to_ma = UNIT_TO_UA.get(unit, 1000.0) / 1000.0
    src = os.path.abspath(json_path)
    m = re.search(r"(\d{4}-\d{2}-\d{2})-(\d{2})-(\d{2})-(\d{2})", str(d.get("file") or ""))
    run_ts = (f"{m.group(1)} {m.group(2)}:{m.group(3)}:{m.group(4)}" if m
              else _run_ts_of(json_path))
    sim_modes = {r[0] for r in conn.execute("SELECT DISTINCT mode FROM sim_current")}
    return _ingest_raw(conn, recs, factor_to_ma, src, run_ts, chip, config, sim_modes)


# ---------------------------------------------------------------- 导出

HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(bold=True, color="FFFFFF")


def style_sheet(ws, widths=None):
    for c in ws[1]:
        c.fill = HEADER_FILL
        c.font = HEADER_FONT
        c.alignment = Alignment(vertical="center")
    ws.freeze_panes = "A2"
    if ws.max_row > 1:
        ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"
    if widths:
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w


def rnd(v, n=3):
    return round(v, n) if isinstance(v, float) else v


def sim_lookup(conn, mode, ids, stage, tier=""):
    """返回 (合计uA, 缺失ID列表, trim集合, tier集合)。
    tier 非空时只取该档位；该档位没有但有无档位('')的旧数据时退回旧数据（兼容旧仿真表）。
    ids 含 "*" = 该 Mode 全部模块合计（跨模式 label_groups 用）。"""
    def rows_of(where, params):
        if tier:
            rows = conn.execute(f"SELECT current_ua,trim,tier FROM sim_current WHERE {where}"
                                " AND tier=?", params + (tier,)).fetchall()
            if rows:
                return rows
            return conn.execute(f"SELECT current_ua,trim,tier FROM sim_current WHERE {where}"
                                " AND tier=''", params).fetchall()
        return conn.execute(f"SELECT current_ua,trim,tier FROM sim_current WHERE {where}",
                            params).fetchall()

    total, missing, trims, tiers = 0.0, [], set(), set()
    found_any = False
    if ids and any(str(x) == "*" for x in ids):
        buckets = [rows_of("mode=? AND stage=?", (mode, stage))]
        if not buckets[0]:
            return None, ["*"], set(), set()
    else:
        buckets = []
        for mid in ids:
            rows = rows_of("mode=? AND module_id=? AND stage=?", (mode, mid, stage))
            if not rows:
                missing.append(mid)
            else:
                buckets.append(rows)
    for rows in buckets:
        found_any = True
        for cur, trim, t in rows:
            total += cur
            if trim:
                trims.add(trim)
            if t:
                tiers.add(t)
    return (total if found_any else None), missing, trims, tiers


def module_names(conn, ids, skip_missing=False):
    """ids -> 仿真模块名串。skip_missing=True 时略过仿真表没有的 ID（组值整组记在首
    编号上时，其余成员没有独立行，人读表里不该出现 ID15 这类噪音）；全缺才回退 IDn。"""
    names, fallback = [], []
    for mid in ids or []:
        rows = conn.execute(
            "SELECT DISTINCT module_name FROM sim_current WHERE module_id=? AND module_name!=''"
            " ORDER BY module_name", (mid,)).fetchall()
        if rows:
            n = "/".join(r[0] for r in rows)
            if n not in names:
                names.append(n)
        else:
            fallback.append(f"ID{mid}")
    if not skip_missing:
        names += fallback
    return " + ".join(names or fallback)


def latest_runs(conn, all_runs=False):
    """runs 行（含 mode_raw）；默认同 (mode,chip,温度) 取 run_ts 最新的一次。"""
    runs = conn.execute(
        "SELECT run_id,mode,chip,temp_c,src_file,run_ts,mode_raw FROM runs"
        " ORDER BY mode,chip,run_ts,run_id").fetchall()
    if not all_runs:
        latest = {}
        for r in runs:
            latest[(r[1], r[2], r[3])] = r  # 同 mode+chip+temp 取最新（已按 run_ts 升序）
        runs = sorted(latest.values(), key=lambda r: (r[1], r[2], r[3] if r[3] is not None else 0))
    return runs


def export_xlsx(conn, out_path, all_runs=False, config=None):
    tier = (config or {}).get("sim_tier") or ""
    runs = latest_runs(conn, all_runs)

    wb = openpyxl.Workbook()

    # ---- ReadMe
    ws = wb.active
    ws.title = "ReadMe"
    ws.append(["电流对比数据库 · 导出说明"])
    ws["A1"].font = Font(bold=True, size=14)
    for line in [
        "",
        f"导出时间：{now_iso()}    数据来源：current.db（current_db.py 生成）",
        "",
        "【Sheet 说明】",
        "  Compare    —— 每模式每模块组一行：仿真 pre/post、实测、偏差%（可直接用于汇报）",
        "  Long       —— 透视长表：Source 列区分 sim_pre / sim_post / meas，做透视图用",
        "  Sim_long   —— 仿真长表原样（单位已统一为 uA）",
        "  Meas_steps —— 实测逐行审计：原始电流、做差、行分类（复验段/归并行也在）",
        "  Runs       —— 本次导出包含的测试 run 列表",
        "",
        "【计算规则】",
        "  1. 基线 = 第一个 OFF 行之前最后一行的电流（通常是最后一个 Lock_step）",
        "  2. 模块电流 = 上一行电流 - 本行电流（逐级关断做差），统一为 uA",
        "  3. NO. 列多个编号（如 45,46）= 一组模块同时关断，仿真侧按组求和对比",
        "  4. LDO 归并（current_config.json 的 ldo_reparent）：子模块不在被测 LDO 下，",
        "     其实测 delta 并入父模块组（编号仍显示父模块号，归并详情见 Note 列）",
        "  5. 第二个及以后的 Init 段 = 锁定复验，忽略（Meas_steps 里有原始行）",
        "  6. 多 run 时默认每个 模式×芯片×温度 取最新一次；--all-runs 可导出全部",
        f"  7. 仿真对比只取档位 sim_tier={tier or '(未过滤)'}（current_config.json 里改）",
        "",
        "【透视建议（Long 页）】",
        "  行=Group/Modules，列=Source（或 Chip/Temp_C），值=Current_uA（用平均值，防多 run 重复计数）",
    ]:
        ws.append([line])
    ws.column_dimensions["A"].width = 100

    # ---- Compare / Long
    cmp_ws = wb.create_sheet("Compare")
    cmp_ws.append(["Mode", "Chip", "Temp_C", "Run_TS", "Step", "Group", "Step_Name", "Modules",
                   "Sim_pre_uA", "Sim_post_uA", "Meas_uA", "Meas-Post_uA", "Meas/Post", "Dev_%", "Note"])
    long_ws = wb.create_sheet("Long")
    long_ws.append(["Mode", "Chip", "Temp_C", "Run_TS", "Step", "Group", "Step_Name", "Modules",
                    "Source", "Trim", "Tier", "Current_uA", "Note"])

    sim_modes = {r[0] for r in conn.execute("SELECT DISTINCT mode FROM sim_current")}
    for run_id, mode, chip, temp, _src, run_ts, _mode_raw in runs:
        groups = conn.execute(
            "SELECT step_order,group_disp,step_name,module_ids,sim_ids,current_ua,note,sim_mode"
            " FROM meas_module WHERE run_id=? ORDER BY step_order", (run_id,)).fetchall()
        for order, disp, step_name, _mids, sim_ids_j, meas_ua, note, sim_mode in groups:
            sim_ids = json.loads(sim_ids_j) if sim_ids_j else None
            names = module_names(conn, sim_ids) if sim_ids else ""
            notes = [note] if note else []
            sim_pre = sim_post = None
            trims, tiers = set(), set()
            lk_mode = sim_mode or mode  # 跨模式 label_groups 的步查它指定的仿真 Mode
            if sim_ids and lk_mode in sim_modes:
                sim_pre, miss_pre, t1, r1 = sim_lookup(conn, lk_mode, sim_ids, "pre", tier)
                sim_post, miss_post, t2, r2 = sim_lookup(conn, lk_mode, sim_ids, "post", tier)
                trims, tiers = t1 | t2, r1 | r2
                miss = sorted(set(miss_pre) & set(miss_post))
                if miss:
                    notes.append(f"仿真表缺ID: {miss}")
            elif sim_ids and lk_mode not in sim_modes:
                notes.append("仿真表未导入" if not sim_modes else f"仿真表无模式 {lk_mode}")
            note_s = "；".join(n for n in notes if n)
            diff = (meas_ua - sim_post) if (sim_post is not None) else None
            ratio = (meas_ua / sim_post) if sim_post else None
            dev = (diff / sim_post * 100.0) if sim_post else None
            cmp_ws.append([mode, chip, temp, run_ts, order, disp, step_name, names,
                           rnd(sim_pre), rnd(sim_post), rnd(meas_ua), rnd(diff),
                           rnd(ratio, 3), rnd(dev, 1), note_s])
            trim_s = ",".join(sorted(trims))
            tier_s = ",".join(sorted(tiers))
            long_ws.append([mode, chip, temp, run_ts, order, disp, step_name, names,
                            "meas", "", "", rnd(meas_ua), note_s])
            if sim_pre is not None:
                long_ws.append([mode, chip, temp, run_ts, order, disp, step_name, names,
                                "sim_pre", trim_s, tier_s, rnd(sim_pre), ""])
            if sim_post is not None:
                long_ws.append([mode, chip, temp, run_ts, order, disp, step_name, names,
                                "sim_post", trim_s, tier_s, rnd(sim_post), ""])

    style_sheet(cmp_ws, [16, 8, 8, 17, 6, 12, 22, 34, 12, 12, 12, 13, 10, 8, 40])
    style_sheet(long_ws, [16, 8, 8, 17, 6, 12, 22, 34, 10, 8, 8, 12, 40])
    if cmp_ws.max_row > 1:
        cmp_ws.conditional_formatting.add(
            f"N2:N{cmp_ws.max_row}",
            ColorScaleRule(start_type="num", start_value=-50, start_color="63BE7B",
                           mid_type="num", mid_value=0, mid_color="FFFFFF",
                           end_type="num", end_value=50, end_color="F8696B"))

    # ---- Sim_long
    ws = wb.create_sheet("Sim_long")
    ws.append(["ID", "Module", "Trim", "Mode", "Stage", "Tier", "Current_uA", "Unit_raw"])
    for r in conn.execute(
            "SELECT module_id,module_name,trim,mode,stage,tier,current_ua,unit_raw FROM sim_current"
            " ORDER BY mode,module_id,stage"):
        ws.append([r[0], r[1], r[2], r[3], r[4], r[5], rnd(r[6]), r[7]])
    style_sheet(ws, [6, 28, 8, 18, 8, 8, 12, 9])

    # ---- Meas_steps
    ws = wb.create_sheet("Meas_steps")
    ws.append(["Mode", "Chip", "Run_TS", "Row", "Seq", "Kind", "NO_raw", "Mode_label",
               "Current_mA", "Delta_mA", "Temp_C", "Note"])
    for run_id, mode, chip, _temp, _src, run_ts, _mraw in runs:
        for r in conn.execute(
                "SELECT row_idx,seq_idx,kind,no_raw,mode_label,current_ma,delta_ma,temp_c,note"
                " FROM meas_raw WHERE run_id=? ORDER BY row_idx", (run_id,)):
            ws.append([mode, chip, run_ts, r[0], r[1], r[2], r[3], r[4],
                       rnd(r[5], 4), rnd(r[6], 4), r[7], r[8]])
    style_sheet(ws, [16, 8, 17, 6, 5, 8, 12, 24, 11, 10, 8, 30])

    # ---- Runs
    ws = wb.create_sheet("Runs")
    ws.append(["Run_ID", "Mode", "Mode_raw", "Chip", "Temp_C", "Run_TS", "Src_file"])
    for r in runs:
        ws.append([r[0], r[1], r[6], r[2], r[3], r[5], r[4]])
    style_sheet(ws, [7, 16, 16, 8, 8, 17, 70])

    wb.save(out_path)


# ---- 汇总簿视觉语言（参考评审报告表：黄表头带/条件行米色/结果白/合计蓝/超差红粗/细边框） ----
C_HEADER, C_SETTING, C_RESULT, C_SEP, C_FLAG = "FFFF00", "EEECE1", "FFFFFF", "B8CCE4", "FF0000"
FONT_NAME = "微软雅黑"
_THIN = Side(style="thin", color="FF000000")
BORDER_ALL = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
FMT_UA, FMT_MA, FMT_PCT = "#,##0.0", "0.000", "0.0%"


def _cell(ws, r, c, val=None, bold=False, fill=None, fmt=None, align="center", size=10):
    cc = ws.cell(row=r, column=c)
    if val is not None:
        cc.value = val
    cc.font = Font(name=FONT_NAME, size=size, bold=bold)
    cc.border = BORDER_ALL
    cc.alignment = Alignment(horizontal=align, vertical="center")
    if fill:
        cc.fill = PatternFill("solid", fgColor=fill)
    if fmt:
        cc.number_format = fmt
    return cc


def _t(v):
    """温度显示：25.0 -> '25℃'。"""
    return ("%g" % v) + "℃"


def _mode_freq(mode, config):
    """测试频率条件行文本。config.mode_freq 优先；否则按模式名首个频段 token 推断
    （2G -> 2.5GHz，5G -> 5.8GHz；BT_2G_TX_DCO5G 这类先出现 2G 算 2G 模式）。"""
    fmap = config.get("mode_freq") or {}
    if mode in fmap:
        return str(fmap[mode])
    m = re.search(r"([25])\s*G", str(mode).upper())
    if m:
        return "2.5GHz" if m.group(1) == "2" else "5.8GHz"
    return ""


def _chart_title(text, sz=1100, bold=True):
    """图表标题（显式字体/字号的富文本）。openpyxl 默认标题不带字体属性，
    Excel 对中文按默认字体度量排版会把文本框排溢出、裁掉开头几个字。"""
    rpr = CharacterProperties(sz=sz, b=bold,
                              latin=DrawFont(typeface=FONT_NAME),
                              ea=DrawFont(typeface=FONT_NAME))
    p = Paragraph(pPr=ParagraphProperties(defRPr=rpr),
                  r=[RegularTextRun(rPr=rpr, t=text)])
    return Title(tx=Text(rich=RichText(p=[p])))


def cmd_summary_export(conn, out_path, config):
    """人直接读的汇总簿：说明 / 总览(模块×模式×温度 + 仿真对比) / 温度趋势(图) / 对比明细。"""
    tier = config.get("sim_tier") or ""
    sim_note = config.get("sim_temp_note") or ""
    thr = float(config.get("delta_flag_pct") or 20) / 100.0

    runs = latest_runs(conn)
    if not runs:
        raise SystemExit("[错误] 库里没有任何实测 run")
    sim_modes = {r[0] for r in conn.execute("SELECT DISTINCT mode FROM sim_current")}
    multi_chip = len({r[2] for r in runs}) > 1
    temps = sorted({r[3] for r in runs if r[3] is not None})
    if not temps:
        temps = [None]
    ref_temp = config.get("delta_ref_temp")
    if ref_temp is None:
        ref_temp = min(temps, key=lambda t: abs((t if t is not None else 25) - 55))
    n_t = len(temps)

    # 列组 = (mode, chip)，按首个 run_id 排（=入库顺序=测试顺序）
    first_id = {}
    for r in runs:
        k = (r[1], r[2])
        if k not in first_id or r[0] < first_id[k]:
            first_id[k] = r[0]
    col_keys = sorted(first_id, key=first_id.get)

    def col_title(mode, chip):
        return f"{mode}({chip})" if multi_chip else mode

    # 逐 run 取基线/末态/模块组
    base_ma, init_ma, end_ma, per_groups = {}, {}, {}, {}   # 键 (mode,chip,temp)
    src_files = set()
    for run_id, mode, chip, temp, src, _ts, _mraw in runs:
        k = (mode, chip, temp)
        src_files.add(os.path.basename(src))
        row = conn.execute(
            "SELECT current_ma FROM meas_raw WHERE run_id=? AND seq_idx=1 AND kind='init'"
            " AND current_ma IS NOT NULL ORDER BY row_idx LIMIT 1", (run_id,)).fetchone()
        init_ma[k] = row[0] if row else None
        row = conn.execute(
            "SELECT current_ma FROM meas_raw WHERE run_id=? AND seq_idx=1 AND kind='lock'"
            " AND current_ma IS NOT NULL ORDER BY row_idx DESC LIMIT 1", (run_id,)).fetchone()
        base_ma[k] = row[0] if row else init_ma[k]
        row = conn.execute(
            "SELECT current_ma FROM meas_raw WHERE run_id=? AND seq_idx=1 AND kind='off'"
            " AND current_ma IS NOT NULL ORDER BY row_idx DESC LIMIT 1", (run_id,)).fetchone()
        end_ma[k] = row[0] if row else None      # 全部关断后的末态电流（末个 OFF 步实测）
        per_groups[k] = conn.execute(
            "SELECT step_order,group_disp,step_name,sim_ids,sim_mode,current_ua,note"
            " FROM meas_module WHERE run_id=? ORDER BY step_order", (run_id,)).fetchall()

    # 行 universe：(disp, step_name)，按平均步序排
    matrix, order_sum, order_cnt, notes, siminfo = {}, {}, {}, {}, {}
    for (mode, chip, temp), groups in per_groups.items():
        for step_order, disp, step_name, sim_ids_j, sim_mode, ua, note in groups:
            key = (disp, step_name)
            matrix.setdefault(key, {})[(mode, chip, temp)] = ua
            order_sum[key] = order_sum.get(key, 0) + step_order
            order_cnt[key] = order_cnt.get(key, 0) + 1
            if sim_ids_j and key not in siminfo:
                siminfo[key] = (json.loads(sim_ids_j), sim_mode)
            if note:
                keep = "；".join(x for x in note.split("；")
                                 if "不在被测LDO" in x or ("映射" in x and "未映射" not in x))
                if keep:
                    notes.setdefault(key, set()).add(keep)
    def _row_sort_key(k):
        """按 buffer 编号从低到高；组合/归并组（"45,46"、"26+28"）取组内最小号；
        非数字标签（DCO2G 等）排最后，按平均步序。"""
        ids = parse_ids(str(k[0]).replace("+", ","))
        if ids:
            return (0, min(ids), order_sum[k] / order_cnt[k])
        return (1, 0, order_sum[k] / order_cnt[k])
    row_keys = sorted(matrix, key=_row_sort_key)

    # 仿真值缓存：行×模式 -> pre/post 合计 µA（tier 过滤）；行名
    sim_val, sim_pre_val, sim_names = {}, {}, {}
    for key in row_keys:
        ids, override = siminfo.get(key, (None, None))
        sim_names[key] = module_names(conn, [i for i in (ids or []) if str(i) != "*"],
                                      skip_missing=True) if ids else ""
        for mode, chip in col_keys:
            lk_mode = override or mode
            if ids and lk_mode in sim_modes:
                sim_val[(key, mode)], _, _, _ = sim_lookup(conn, lk_mode, ids, "post", tier)
                sim_pre_val[(key, mode)], _, _, _ = sim_lookup(conn, lk_mode, ids, "pre", tier)
    # 标签行（DCO 等）排在末尾——Σ 分两段：LO 模块（可与仿真对）/ 总合计（含标签行，只有实测）
    n_label = sum(1 for k in row_keys if parse_ids(str(k[0]).replace("+", ",")) is None)

    wb = openpyxl.Workbook()

    # ================= 说明 =================
    ws = wb.active
    ws.title = "说明"
    ws["A1"] = "各模式功耗汇总簿 · 读法"
    ws["A1"].font = Font(name=FONT_NAME, bold=True, size=14)
    lines = [
        "",
        f"导出时间 {now_iso()}；数据源 current.db（current_db.py summary 生成）",
        f"实测：{len(runs)} 个 run（模式×温度），温度点 {', '.join(_t(t) for t in temps if t is not None) or '未知'}；"
        f"源文件 {', '.join(sorted(src_files))}",
        f"仿真：档位 {tier or '未过滤'} / post-sim（pre-sim 在「对比明细」页）/ 温度 {sim_note or '未标注'}",
        "",
        "【总览页】行=模块（按 buffer 编号从低到高，组合组取组内最小号，DCO 标签行在最后），",
        "  列=模式×温度的实测电流 + 仿真参考 + 偏差%。顶部条件行：测试频率（2G=2.5GHz/5G=5.8GHz）、",
        "  锁定后总电流（做差基线）、全关残留电流（末个 OFF 步实测=全部关断后的末态）。",
        f"  偏差% = (实测@{_t(ref_temp) if ref_temp is not None else '?'} − 仿真post) / 仿真post，"
        f"|偏差%| > {thr * 100:.0f}% 标红（阈值在 current_config.json 的 delta_flag_pct 改）。",
        f"  注意：仿真是 {sim_note or '单温度'} 单点，与各实测温度直接对比含系统性温差。",
        "",
        "【颜色】黄=表头；米色=条件/汇总行（mA）；白=模块行（µA）；蓝=Σ合计；红粗=超阈值偏差。",
        "",
        "【计算规则】",
        "  基线 = 每模式段第一个 OFF 前最后一行（末个 Lock_step）；模块电流 = 上一行 − 本行。",
        "  底部 Σ 分两行：Σ LO 模块合计（不含 DCO 等标签行，口径与仿真一致、带偏差%）；",
        "  Σ 总合计（含标签行，只有实测）。锁定后总电流是整机电流，不与仿真直接对比。",
        f"  LDO 归并 {config.get('ldo_reparent')}：子模块实测并入父组，仿真侧"
        f"{'也求和' if config.get('ldo_reparent_sim_add_child') else '不计子模块'}。",
        "  多 run 时每个 模式×芯片×温度 取时间最新一次。",
        "",
        "偏差% 与 Σ 合计是公式，改动数值后 Excel 会自动重算。",
    ]
    for i, line in enumerate(lines, 2):
        ws.cell(row=i, column=1, value=line).font = Font(name=FONT_NAME, size=10)
    ws.column_dimensions["A"].width = 110

    # ================= 总览 =================
    ws = wb.create_sheet("总览")
    FIX = 4                                  # 编号/模块/仿真模块名/单位
    grp_w = n_t + 2                          # 每模式：温度列 + 仿真 + 偏差%
    note_col = FIX + len(col_keys) * grp_w + 1

    def ref(r, c):
        return f"{get_column_letter(c)}{r}"

    # -- 表头带（3 行黄）：先整片打底，再填字做合并
    for r in range(1, 4):
        for c in range(1, note_col + 1):
            _cell(ws, r, c, fill=C_HEADER)
    for c, (title, w) in enumerate(zip(["编号", "模块 (OFF 步)", "仿真模块名", "单位"],
                                       [9, 24, 32, 6]), 1):
        ws.merge_cells(start_row=1, start_column=c, end_row=3, end_column=c)
        _cell(ws, 1, c, title, bold=True, fill=C_HEADER)
        ws.column_dimensions[get_column_letter(c)].width = w
    for gi, (mode, chip) in enumerate(col_keys):
        c0 = FIX + 1 + gi * grp_w
        ws.merge_cells(start_row=1, start_column=c0, end_row=1, end_column=c0 + grp_w - 1)
        _cell(ws, 1, c0, col_title(mode, chip), bold=True, fill=C_HEADER)
        if n_t > 1:
            ws.merge_cells(start_row=2, start_column=c0, end_row=2, end_column=c0 + n_t - 1)
        _cell(ws, 2, c0, "实测", bold=True, fill=C_HEADER)
        _cell(ws, 2, c0 + n_t, "仿真", bold=True, fill=C_HEADER)
        _cell(ws, 2, c0 + n_t + 1, "偏差", bold=True, fill=C_HEADER)
        for ti, t in enumerate(temps):
            _cell(ws, 3, c0 + ti, _t(t) if t is not None else "?", bold=True, fill=C_HEADER)
        _cell(ws, 3, c0 + n_t, (f"{sim_note} {tier}".strip() or "post"), bold=True, fill=C_HEADER)
        _cell(ws, 3, c0 + n_t + 1,
              f"vs{_t(ref_temp) if ref_temp is not None else ''}", bold=True, fill=C_HEADER)
        for ti in range(n_t):
            ws.column_dimensions[get_column_letter(c0 + ti)].width = 9.5
        ws.column_dimensions[get_column_letter(c0 + n_t)].width = 10
        ws.column_dimensions[get_column_letter(c0 + n_t + 1)].width = 8.5
    ws.merge_cells(start_row=1, start_column=note_col, end_row=3, end_column=note_col)
    _cell(ws, 1, note_col, "备注", bold=True, fill=C_HEADER)
    ws.column_dimensions[get_column_letter(note_col)].width = 42

    r_freq, r_base, r_end = 4, 5, 6
    r_mod0 = 7
    r_sum = r_mod0 + len(row_keys)          # Σ LO 模块行（不含末尾标签行）
    r_all = r_sum + 1 if n_label else None  # Σ 总合计行（含 DCO 等标签行，只有实测）
    r_lo_last = r_sum - 1 - n_label         # 最后一个 LO 模块行

    # -- 条件/汇总行（米色）：测试频率 / 锁定后总电流 / 全关残留
    for rr, name, unit in ((r_freq, "测试频率", ""),
                           (r_base, "锁定后总电流", "mA"),
                           (r_end, "全关残留电流", "mA")):
        _cell(ws, rr, 1, fill=C_SETTING)
        _cell(ws, rr, 2, name, bold=True, fill=C_SETTING, align="left")
        _cell(ws, rr, 3, "", fill=C_SETTING)
        _cell(ws, rr, 4, unit, fill=C_SETTING)
    for gi, (mode, chip) in enumerate(col_keys):
        c0 = FIX + 1 + gi * grp_w
        for cc in range(c0, c0 + grp_w):     # 频率行整组打底再合并
            _cell(ws, r_freq, cc, fill=C_SETTING)
        ws.merge_cells(start_row=r_freq, start_column=c0,
                       end_row=r_freq, end_column=c0 + grp_w - 1)
        _cell(ws, r_freq, c0, _mode_freq(mode, config), bold=True, fill=C_SETTING)
        for ti, t in enumerate(temps):
            v = base_ma.get((mode, chip, t))
            _cell(ws, r_base, c0 + ti, rnd(v, 3) if v is not None else "",
                  fill=C_SETTING, fmt=FMT_MA)
            v = end_ma.get((mode, chip, t))
            _cell(ws, r_end, c0 + ti, rnd(v, 3) if v is not None else "",
                  fill=C_SETTING, fmt=FMT_MA)
        # 基线/末态是含 DCO 与杂项的整机电流，仿真只覆盖 LO 模块——不做行内对比，
        # 仿真对比见底部「Σ LO 模块合计」行
        _cell(ws, r_base, c0 + n_t, "", fill=C_SETTING)
        _cell(ws, r_end, c0 + n_t, "", fill=C_SETTING)
        _cell(ws, r_base, c0 + n_t + 1, "", fill=C_SETTING)
        _cell(ws, r_end, c0 + n_t + 1, "", fill=C_SETTING)
    _cell(ws, r_freq, note_col, "2G 模式=2.5GHz，5G 模式=5.8GHz（config.mode_freq 可改）",
          fill=C_SETTING, align="left")
    _cell(ws, r_base, note_col, "基线=末个 Lock_step（含 DCO 及未列模块，仿真对比看 Σ LO 行）",
          fill=C_SETTING, align="left")
    _cell(ws, r_end, note_col, "末个 OFF 步实测=全部关断后仍在流的电流（≈基线−Σ模块）",
          fill=C_SETTING, align="left")

    # -- 模块行（白，µA）
    for i, key in enumerate(row_keys):
        rr = r_mod0 + i
        disp, step_name = key
        _cell(ws, rr, 1, disp)
        _cell(ws, rr, 2, step_name, align="left")
        _cell(ws, rr, 3, sim_names.get(key, ""), align="left")
        _cell(ws, rr, 4, "µA")
        for gi, (mode, chip) in enumerate(col_keys):
            c0 = FIX + 1 + gi * grp_w
            for ti, t in enumerate(temps):
                v = matrix[key].get((mode, chip, t))
                _cell(ws, rr, c0 + ti, rnd(v, 1) if v is not None else "", fmt=FMT_UA)
            sv = sim_val.get((key, mode))
            _cell(ws, rr, c0 + n_t, rnd(sv, 1) if sv is not None else "", fmt=FMT_UA)
            ci_ref = c0 + temps.index(ref_temp) if ref_temp in temps else c0
            m_ref, s_ref = ref(rr, ci_ref), ref(rr, c0 + n_t)
            _cell(ws, rr, c0 + n_t + 1,
                  f'=IF(OR({m_ref}="",{s_ref}=""),"",({m_ref}-{s_ref})/{s_ref})', fmt=FMT_PCT)
        note_parts = sorted(notes.get(key, ()))
        if key not in siminfo and parse_ids(str(disp).replace("+", ",")) is None:
            note_parts.append("仿真无对应项（标签未映射，config.label_groups 可配）")
        _cell(ws, rr, note_col, "；".join(note_parts), align="left")

    # -- Σ合计（蓝）：LO 模块行（可与仿真对）+ 总合计行（含 DCO 等标签行，只有实测）
    _cell(ws, r_sum, 1, fill=C_SEP)
    _cell(ws, r_sum, 2, "Σ LO 模块合计" if n_label else "Σ 模块合计",
          bold=True, fill=C_SEP, align="left")
    _cell(ws, r_sum, 3, fill=C_SEP)
    _cell(ws, r_sum, 4, "µA", bold=True, fill=C_SEP)
    for gi, (mode, chip) in enumerate(col_keys):
        c0 = FIX + 1 + gi * grp_w
        for ti in range(n_t + 1):        # 温度列 + 仿真列都求和（标签行仿真为空不影响）
            col = c0 + ti
            _cell(ws, r_sum, col,
                  f"=SUM({ref(r_mod0, col)}:{ref(r_lo_last, col)})",
                  bold=True, fill=C_SEP, fmt=FMT_UA)
        ci_ref = c0 + temps.index(ref_temp) if ref_temp in temps else c0
        m_ref, s_ref = ref(r_sum, ci_ref), ref(r_sum, c0 + n_t)
        _cell(ws, r_sum, c0 + n_t + 1,
              f'=IF(OR({m_ref}=0,{s_ref}=0),"",({m_ref}-{s_ref})/{s_ref})',
              bold=True, fill=C_SEP, fmt=FMT_PCT)
    _cell(ws, r_sum, note_col, "不含下方标签行（DCO 等），口径与仿真一致" if n_label else "",
          fill=C_SEP, align="left")
    if r_all:
        _cell(ws, r_all, 1, fill=C_SEP)
        _cell(ws, r_all, 2, "Σ 总合计（含 DCO 等标签行）", bold=True, fill=C_SEP, align="left")
        _cell(ws, r_all, 3, fill=C_SEP)
        _cell(ws, r_all, 4, "µA", bold=True, fill=C_SEP)
        for gi, (mode, chip) in enumerate(col_keys):
            c0 = FIX + 1 + gi * grp_w
            for ti in range(n_t):
                col = c0 + ti
                _cell(ws, r_all, col,
                      f"=SUM({ref(r_mod0, col)}:{ref(r_sum - 1, col)})",
                      bold=True, fill=C_SEP, fmt=FMT_UA)
            _cell(ws, r_all, c0 + n_t, "", fill=C_SEP)
            _cell(ws, r_all, c0 + n_t + 1, "", fill=C_SEP)
        _cell(ws, r_all, note_col, "仿真未覆盖标签行，无对比", fill=C_SEP, align="left")

    # 偏差%列条件格式：|Δ| 超阈值红粗（含条件行与 Σ 行）
    red = Font(name=FONT_NAME, size=10, bold=True, color=C_FLAG)
    for gi in range(len(col_keys)):
        cd = FIX + 1 + gi * grp_w + n_t + 1
        rng = f"{ref(r_base, cd)}:{ref(r_sum, cd)}"
        ws.conditional_formatting.add(rng, CellIsRule(
            operator="notBetween", formula=[str(-thr), str(thr)], font=red))
    ws.freeze_panes = ws.cell(row=r_base, column=FIX + 1)

    # ================= 温度趋势 =================
    n_charts = 0
    if len([t for t in temps if t is not None]) >= 2:
        ws = wb.create_sheet("温度趋势")
        ws.sheet_view.showGridLines = False
        cur_row = 1

        def chart_block(title, row_names, values_of, unit):
            """左侧数据块 + 右侧折线图。values_of(name, temp) -> 数值或 None。"""
            nonlocal cur_row, n_charts
            _cell(ws, cur_row, 1, title, bold=True, size=12).border = Border()
            hdr = cur_row + 1
            _cell(ws, hdr, 1, "系列", bold=True, fill=C_HEADER)
            for ti, t in enumerate(temps):
                _cell(ws, hdr, 2 + ti, _t(t), bold=True, fill=C_HEADER)
            for ri, name in enumerate(row_names):
                _cell(ws, hdr + 1 + ri, 1, name, align="left")
                for ti, t in enumerate(temps):
                    v = values_of(name, t)
                    _cell(ws, hdr + 1 + ri, 2 + ti, rnd(v, 2) if v is not None else "")
            last = hdr + len(row_names)
            ch = LineChart()
            ch.title = _chart_title(title)
            ch.style = 12
            ch.y_axis.title = _chart_title(unit, sz=1000, bold=False)
            ch.x_axis.title = _chart_title("温度", sz=1000, bold=False)
            ch.height, ch.width = max(7.5, 1.2 + 0.42 * len(row_names)), 16
            data = Reference(ws, min_col=1, min_row=hdr + 1, max_col=1 + n_t, max_row=last)
            ch.add_data(data, from_rows=True, titles_from_data=True)
            ch.set_categories(Reference(ws, min_col=2, min_row=hdr, max_col=1 + n_t))
            for s in ch.series:
                s.marker.symbol = "circle"
                s.marker.size = 5
                s.smooth = False
            ws.add_chart(ch, ref(cur_row + 1, n_t + 3))
            n_charts += 1
            cur_row = last + max(3, int(ch.height * 2) - len(row_names) - 1)

        ws.column_dimensions["A"].width = 38
        for ti in range(n_t):
            ws.column_dimensions[get_column_letter(2 + ti)].width = 10
        chart_block("各模式锁定后总电流 vs 温度 (mA)",
                    [col_title(m, c) for m, c in col_keys],
                    lambda name, t: next((base_ma.get((m, c, t)) for m, c in col_keys
                                          if col_title(m, c) == name), None), "mA")
        for mode, chip in col_keys:
            names = [k[1] for k in row_keys
                     if any((mode, chip, t) in matrix[k] for t in temps)]
            keyof = {k[1]: k for k in row_keys}
            chart_block(f"{col_title(mode, chip)} 各模块电流 vs 温度 (µA)", names,
                        lambda name, t, _m=mode, _c=chip, _ko=keyof:
                            matrix[_ko[name]].get((_m, _c, t)), "µA")

    # ================= 对比明细 =================
    ws = wb.create_sheet("对比明细")
    hdrs = ["模式", "芯片", "温度", "编号", "模块 (OFF 步)", "仿真模块名",
            "实测_µA", "仿真pre_µA", "仿真post_µA", "Δ_µA", "Δ%", "备注"]
    widths = [20, 7, 8, 9, 24, 32, 11, 11, 11, 10, 8, 40]
    for c, (h, w) in enumerate(zip(hdrs, widths), 1):
        _cell(ws, 1, c, h, bold=True, fill=C_HEADER)
        ws.column_dimensions[get_column_letter(c)].width = w
    red = Font(name=FONT_NAME, size=10, bold=True, color=C_FLAG)
    rr = 2
    for mode, chip in col_keys:
        for t in temps:
            k3 = (mode, chip, t)
            if k3 not in per_groups:
                continue
            for key in row_keys:
                if k3 not in matrix.get(key, {}):
                    continue
                disp, step_name = key
                _cell(ws, rr, 1, col_title(mode, chip), align="left")
                _cell(ws, rr, 2, chip)
                _cell(ws, rr, 3, t if t is not None else "")
                _cell(ws, rr, 4, disp)
                _cell(ws, rr, 5, step_name, align="left")
                _cell(ws, rr, 6, sim_names.get(key, ""), align="left")
                _cell(ws, rr, 7, rnd(matrix[key][k3], 1), fmt=FMT_UA)
                pv, sv = sim_pre_val.get((key, mode)), sim_val.get((key, mode))
                _cell(ws, rr, 8, rnd(pv, 1) if pv is not None else "", fmt=FMT_UA)
                _cell(ws, rr, 9, rnd(sv, 1) if sv is not None else "", fmt=FMT_UA)
                _cell(ws, rr, 10, f'=IF(OR(G{rr}="",I{rr}=""),"",G{rr}-I{rr})', fmt=FMT_UA)
                _cell(ws, rr, 11, f'=IF(OR(G{rr}="",I{rr}=""),"",(G{rr}-I{rr})/I{rr})',
                      fmt=FMT_PCT)
                _cell(ws, rr, 12, "；".join(sorted(notes.get(key, ()))), align="left")
                rr += 1
    if rr > 2:
        ws.auto_filter.ref = f"A1:L{rr - 1}"
        ws.conditional_formatting.add(f"K2:K{rr - 1}", ColorScaleRule(
            start_type="num", start_value=-0.5, start_color="63BE7B",
            mid_type="num", mid_value=0, mid_color="FFFFFF",
            end_type="num", end_value=0.5, end_color="F8696B"))
        ws.conditional_formatting.add(f"K2:K{rr - 1}", CellIsRule(
            operator="notBetween", formula=[str(-thr), str(thr)], font=red))
    ws.freeze_panes = "A2"

    wb.save(out_path)
    return len(runs), len(row_keys), n_charts


def cmd_summary(args):
    root = os.path.dirname(os.path.abspath(args.db))
    config, _, _ = load_config(root, args.config)
    conn = open_db(args.db)
    n_runs, n_rows, n_charts = cmd_summary_export(conn, args.out, config)
    conn.close()
    print(f"[完成] 功耗汇总簿: {args.out}（{n_runs} 个 run，矩阵 {n_rows} 行，{n_charts} 张趋势图）")


# ---------------------------------------------------------------- 命令

def walk_xlsx(root, skip_dirs, exclude_globs=()):
    """递归列出 root 下所有 xlsx（跳过 skip_dirs 子树、~$ 锁文件和排除名单）。"""
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = [d for d in dirnames if d not in skip_dirs]
        for f in sorted(filenames):
            if not f.lower().endswith((".xlsx", ".xlsm")) or f.startswith("~$"):
                continue
            if any(fnmatch.fnmatch(f, pat) for pat in exclude_globs):
                continue
            yield os.path.join(dirpath, f)


def find_sim_workbook(root, config):
    """返回 (工作簿路径, tab名) 或 (None, None)。不认文件名，按表头内容全递归扫。"""
    if config.get("sim_workbook"):
        p = config["sim_workbook"]
        return (p if os.path.isabs(p) else os.path.join(root, p)), config.get("sim_sheet")
    skip = set(config.get("skip_dirs") or [])
    excl = list(config.get("exclude_globs") or [])
    sim_sheet = norm(config.get("sim_sheet", "Current_data"))
    cands = []
    for path in walk_xlsx(root, skip, excl):
        try:
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            for sn in wb.sheetnames:
                if norm(sn) == sim_sheet or any(
                        match_sim_header(r) for r in
                        wb[sn].iter_rows(min_row=1, max_row=30, values_only=True)):
                    cands.append((os.path.getmtime(path), path, sn))
                    break
            wb.close()
        except Exception:
            continue
    if not cands:
        return None, None
    cands.sort(reverse=True)
    if len(cands) > 1:
        print("[提示] 发现多个疑似仿真长表，取最新修改的；其余可在 config.sim_workbook 里指定：")
        for _mt, p, sn in cands:
            print(f"       {os.path.relpath(p, root)} / {sn}")
    return cands[0][1], cands[0][2]


def cmd_build(args):
    root = os.path.abspath(args.root)
    if not os.path.isdir(root):
        raise SystemExit(f"[错误] 根目录不存在: {root}")
    config, cfg_path, created = load_config(root, args.config)
    if created:
        print(f"[提示] 首次运行，已生成配置 {cfg_path}（LDO 归并/模式映射/标签映射都在里面改）")

    db_path = args.db or os.path.join(root, "current.db")
    if os.path.exists(db_path):
        os.remove(db_path)  # build = 全量重建；增量请用 ingest-* 子命令
    conn = open_db(db_path)

    if args.sim:
        sim_wb, sim_sheet = args.sim, config.get("sim_sheet")
    else:
        sim_wb, sim_sheet = find_sim_workbook(root, config)
    if sim_wb and os.path.exists(sim_wb):
        n = ingest_sim(conn, sim_wb, sim_sheet)
        print(f"[仿真] {os.path.relpath(sim_wb, root)} / {sim_sheet or '自动'} -> {n} 行")
    else:
        print("[警告] 未找到仿真长表（任意工作簿中表头含 ID/Mode/Current+simulation|Unit|Tier 的 tab），"
              "只导入实测；可用 --sim 或 config.sim_workbook 指定")

    skip = set(config.get("skip_dirs") or [])
    excl = list(config.get("exclude_globs") or [])
    out = args.out or os.path.join(root, "Current_compare_pivot.xlsx")
    excl.append(os.path.basename(out))
    result_glob = config.get("result_glob", "Result*.xlsx")
    sim_modes = {r[0] for r in conn.execute("SELECT DISTINCT mode FROM sim_current")}
    n_runs = 0
    mapping = {}  # mode_raw -> (mode, how)
    for f in walk_xlsx(root, skip, excl):
        if not fnmatch.fnmatch(os.path.basename(f), result_glob):
            continue
        folder = os.path.basename(os.path.dirname(f))
        results = ingest_result_file(conn, f, args.chip, config, sim_modes, folder_mode=folder)
        n_runs += len(results)
        rel = os.path.relpath(f, root)
        if len(results) > 1:
            print(f"[实测] {rel}  全模式单文件 -> {len(results)} 个 (模式,温度) 段:")
        for run_id, mode, mode_raw, how, temp, n_steps in results:
            mapping[(mode_raw, mode)] = how
            pre = "        " if len(results) > 1 else f"[实测] {rel}  "
            print(f"{pre}{mode:<22} run#{run_id}  {n_steps} 个模块组  "
                  f"{temp if temp is not None else '?'}°C")
    if n_runs == 0:
        print("[警告] 没有扫到任何 Result 文件")

    if mapping:
        print("[模式映射] 实测段标签 -> 仿真 Mode（config.mode_map 可强制指定）:")
        how_disp = {"config": "config指定", "auto": "自动匹配", "ambig": "⚠多个候选未映射",
                    "none": "⚠仿真表无此模式", "folder": "⚠按文件夹名(段内标签与文件夹不符)"}
        for raw_name, mode in sorted(mapping):
            how = mapping[(raw_name, mode)]
            arrow = "=" if raw_name == mode else "->"
            print(f"    {raw_name:<24} {arrow} {mode:<24} {how_disp.get(how, how)}")

    export_xlsx(conn, out, all_runs=args.all_runs, config=config)
    conn.close()
    print(f"[完成] 数据库: {db_path}")
    print(f"[完成] 导出:   {out}")


def cmd_ingest_sim(args):
    conn = open_db(args.db)
    n = ingest_sim(conn, args.xlsx, args.sheet)
    conn.close()
    print(f"[仿真] {args.xlsx} -> {n} 行")


def cmd_ingest_run(args):
    root = os.path.dirname(os.path.abspath(args.db))
    config, _, _ = load_config(root, args.config)
    conn = open_db(args.db)
    run_id, n_steps, temp, run_ts = ingest_run(conn, args.xlsx, args.mode, args.chip, config,
                                               sheet_name=args.sheet)
    conn.close()
    print(f"[实测] run#{run_id} mode={args.mode} {n_steps} 个模块组 {temp}°C {run_ts}")


def cmd_ingest_probe(args):
    root = os.path.dirname(os.path.abspath(args.db))
    config, _, _ = load_config(root, args.config)
    conn = open_db(args.db)
    results = ingest_probe_json(conn, args.json, args.chip, config)
    conn.close()
    for run_id, mode, mode_raw, how, temp, n_steps in results:
        print(f"[实测] {mode:<22} (段标签 {mode_raw}, {how})  run#{run_id}  "
              f"{n_steps} 个模块组  {temp if temp is not None else '?'}°C")


def cmd_export(args):
    root = os.path.dirname(os.path.abspath(args.db))
    config, _, _ = load_config(root, args.config)
    conn = open_db(args.db)
    export_xlsx(conn, args.out, all_runs=args.all_runs, config=config)
    conn.close()
    print(f"[完成] 导出: {args.out}")


def main():
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    ap = argparse.ArgumentParser(description="电流数据库：仿真+实测 -> SQLite -> pivot 长表")
    sub = ap.add_subparsers(dest="cmd", required=True)

    b = sub.add_parser("build", help="一键：扫描数据根目录，全量重建库并导出")
    b.add_argument("--root", required=True, help="数据根目录，如 D:\\Excel")
    b.add_argument("--chip", default="C1", help="芯片编号（默认 C1）")
    b.add_argument("--sim", help="仿真工作簿路径（默认自动找 Current_all_mode*.xlsx）")
    b.add_argument("--db", help="SQLite 输出路径（默认 root/current.db）")
    b.add_argument("--out", help="Excel 输出路径（默认 root/Current_compare_pivot.xlsx）")
    b.add_argument("--config", help="配置文件路径（默认 root/current_config.json）")
    b.add_argument("--all-runs", action="store_true", help="导出全部 run（默认每模式×芯片取最新）")
    b.set_defaults(func=cmd_build)

    s = sub.add_parser("ingest-sim", help="导入仿真长表")
    s.add_argument("--db", required=True)
    s.add_argument("--xlsx", required=True)
    s.add_argument("--sheet", default="Current_data")
    s.set_defaults(func=cmd_ingest_sim)

    r = sub.add_parser("ingest-run", help="导入单个实测 Result 文件")
    r.add_argument("--db", required=True)
    r.add_argument("--xlsx", required=True)
    r.add_argument("--mode", required=True, help="模式名（需与仿真表 Mode 一致）")
    r.add_argument("--chip", default="C1")
    r.add_argument("--sheet", help="tab 名（默认自动扫描）")
    r.add_argument("--config", help="配置文件路径（默认取 db 同目录 current_config.json）")
    r.set_defaults(func=cmd_ingest_run)

    p = sub.add_parser("ingest-probe", help="导入 probe_allmode_result.py --json 的产物")
    p.add_argument("--db", required=True)
    p.add_argument("--json", required=True)
    p.add_argument("--chip", default="C1")
    p.add_argument("--config", help="配置文件路径（默认取 db 同目录 current_config.json）")
    p.set_defaults(func=cmd_ingest_probe)

    e = sub.add_parser("export", help="从库导出 Excel")
    e.add_argument("--db", required=True)
    e.add_argument("--out", required=True)
    e.add_argument("--all-runs", action="store_true")
    e.add_argument("--config", help="配置文件路径（默认取 db 同目录 current_config.json）")
    e.set_defaults(func=cmd_export)

    m = sub.add_parser("summary", help="导出人直接读的功耗汇总簿（总览矩阵+温度趋势图+对比明细）")
    m.add_argument("--db", required=True)
    m.add_argument("--out", required=True)
    m.add_argument("--config", help="配置文件路径（默认取 db 同目录 current_config.json）")
    m.set_defaults(func=cmd_summary)

    args = ap.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
