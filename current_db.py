#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
current_db.py — 电流数据库（阶段一：SQLite + pivot 长表导出）

把两类数据统一进一个 SQLite 库，并导出可直接做数据透视的 Excel：
  1) 仿真长表：某工作簿的 Current_data 页（ID/Module/Trim/Mode/simulation/Tier/Current/Unit）
  2) 实测结果：各模式文件夹下的 Result*.xlsx（NO./Mode/Temperature/Current_mA 列按表头名自动定位）

实测解析规则：
  - 一个序列从 Init 行开始；基线 = 第一个 OFF 行之前最后一行的电流（通常是最后一个 Lock_step）
  - 模块电流 = 上一行电流 - 本行电流（逐级关断做差），统一换算成 uA
  - 第二个及以后的 Init 段 = 锁定复验，忽略（原始行仍入库审计）
  - "chamber close" 行为终止行，不参与做差
  - NO. 列多个编号（如 "45,46"）= 该步同时关断的一组模块，按组对比（仿真侧求和）
  - LDO 归并（config.ldo_reparent，如 28->26）：子模块不在被测 LDO 下，
    其实测 delta 并入父模块组；对比时仿真侧同样求和（meas(26)+meas(28) vs sim(26)+sim(28)）
  - NO. 列非数字标签（如 "DCO5G"）：可在 config.label_groups 里映射到仿真模块 ID

用法（在数据所在机器上）：
  python current_db.py build --root D:\\Excel --chip C1
    首次运行会在 root 下生成 current_config.json（模式映射/LDO 归并等都在里面改），
    并输出 current.db + Current_compare_pivot.xlsx

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
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.formatting.rule import ColorScaleRule
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
        "ldo_reparent": "子模块ID -> 父模块ID：子模块不在被测 LDO 下，实测/仿真都并入父模块组",
        "label_groups": "NO. 列非数字标签 -> 仿真模块ID列表，如 {\"DCO5G\": [21]}",
    },
    "sim_workbook": None,
    "sim_sheet": "Current_data",
    "result_glob": "Result*.xlsx",
    "result_sheet": None,
    "skip_dirs": ["Simulation", "自动化"],
    "mode_map": {},
    "ldo_reparent": {"8": "6", "28": "26"},
    "label_groups": {},
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
    src_file TEXT, run_ts TEXT, ingested_ts TEXT);
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
    current_ua REAL, note TEXT);
CREATE TABLE IF NOT EXISTS sim_current(
    id INTEGER PRIMARY KEY,
    module_id INTEGER, module_name TEXT, trim TEXT, mode TEXT,
    stage TEXT, tier TEXT, current_ua REAL, unit_raw TEXT, src_file TEXT);
"""


def open_db(path):
    conn = sqlite3.connect(path)
    conn.executescript(SCHEMA)
    return conn


# ---------------------------------------------------------------- 仿真表导入

def match_sim_header(row):
    """仿真长表表头：ID + Mode + Current*，且有 simulation/Unit/Tier 佐证。"""
    names = [norm(c) for c in row]
    return ("id" in names and "mode" in names
            and any(n.startswith("current") for n in names)
            and (any(n.startswith("simulation") for n in names)
                 or "unit" in names or "tier" in names))


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

def find_result_sheet(wb, sheet_name):
    """返回 (worksheet, 表头行号, 列映射)。按表头名定位，不按列字母。"""
    names = [sheet_name] if sheet_name else wb.sheetnames
    for sn in names:
        if sn not in wb.sheetnames:
            continue
        ws = wb[sn]
        for i, row in enumerate(ws.iter_rows(min_row=1, max_row=30, values_only=True), 1):
            no_col = mode_col = cur_col = temp_col = None
            for j, c in enumerate(row):
                n = norm(c)
                if n in ("no.", "no", "no．"):
                    no_col = j
                elif n == "mode" and mode_col is None:
                    mode_col = j
                elif n.startswith("current") and cur_col is None:
                    cur_col = j
                elif "temp" in n and temp_col is None:
                    temp_col = j
            if no_col is not None and cur_col is not None:
                if mode_col is None:
                    mode_col = no_col + 1
                unit = "ma"
                m = re.search(r"current[_\s]*([munµμ]?a)", norm(row[cur_col]))
                if m:
                    unit = m.group(1)
                return ws, i, {"no": no_col, "mode": mode_col, "cur": cur_col,
                               "temp": temp_col, "unit": unit}
    return None, None, None


def classify_rows(ws, hdr, cols):
    """逐行分类并做差。返回 (rows, temp)。
    rows: dict(row_idx, no_raw, label, cur_ma, delta_ma, temp, seq, kind)"""
    factor_to_ma = UNIT_TO_UA.get(cols["unit"], 1000.0) / 1000.0  # 原始单位 -> mA
    out = []
    seq = 0
    prev_cur = None
    temp_first = None
    for i, row in enumerate(ws.iter_rows(min_row=hdr + 1, values_only=True), hdr + 1):
        no_raw = cell(row, cols["no"])
        label = cell(row, cols["mode"])
        cur = as_float(cell(row, cols["cur"]))
        temp = as_float(cell(row, cols["temp"])) if cols["temp"] is not None else None
        if no_raw is None and label is None and cur is None:
            continue
        ln, nn = norm(label), norm(no_raw)
        if "chamber" in ln or "chamber" in nn:
            kind = "chamber"
        elif ln.startswith("init") or ln.startswith("inital") or ln.startswith("initail"):
            kind = "init"
            seq += 1
        elif "lock" in ln:
            kind = "lock"
        elif ln.startswith("off"):
            kind = "off"
        else:
            kind = "other"
        if seq == 0 and cur is not None:
            seq = 1  # 没有显式 Init 行时，第一段视为正式测量
            if kind == "other":
                kind = "init"
        delta = None
        if seq == 1 and kind in ("lock", "off", "other") and cur is not None and prev_cur is not None:
            delta = prev_cur * factor_to_ma - cur * factor_to_ma
        if seq == 1 and cur is not None and kind != "chamber":
            prev_cur = cur
        if temp is not None and temp_first is None:
            temp_first = temp
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
        if ids is None:
            mapped = label_groups.get(disp)
            if mapped:
                sim_ids = [int(x) for x in mapped]
                note = f"标签 {disp} 按 config.label_groups 映射到仿真 ID {sim_ids}"
            else:
                note = "标签未映射仿真模块（可在 current_config.json 的 label_groups 补充）"
        steps.append(dict(row_idx=r["row_idx"], ids=ids, sim_ids=sim_ids, disp=disp,
                          step_name=step_name, delta_ua=r["delta_ma"] * 1000.0, note=note))

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
        parent_step["sim_ids"] = (parent_step["sim_ids"] or []) + [child]
        parent_step["disp"] = parent_step["disp"] + f"+{child}"
        parent_step["note"] = (parent_step["note"] + "；" if parent_step["note"] else "") + \
            f"含模块{child}（{child}不在被测LDO下，实测与仿真均并入）"
        absorbed[child_step["row_idx"]] = parent_step["disp"]
    steps = [s for s in steps if s["row_idx"] not in absorbed]
    for order, s in enumerate(steps, 1):
        s["order"] = order
    return steps, absorbed


def ingest_run(conn, xlsx, mode, chip, config, sheet_name=None):
    wb = openpyxl.load_workbook(xlsx, read_only=True, data_only=True)
    try:
        ws, hdr, cols = find_result_sheet(wb, sheet_name or config.get("result_sheet"))
        if ws is None:
            raise SystemExit(f"[错误] {os.path.basename(xlsx)} 里找不到含 NO./Current 表头的 tab")
        rows, temp = classify_rows(ws, hdr, cols)
        steps, absorbed = build_groups(rows, config)

        m = re.search(r"(\d{4}-\d{2}-\d{2})-(\d{2})-(\d{2})-(\d{2})", os.path.basename(xlsx))
        if m:
            run_ts = f"{m.group(1)} {m.group(2)}:{m.group(3)}:{m.group(4)}"
        else:
            run_ts = datetime.datetime.fromtimestamp(os.path.getmtime(xlsx)).strftime("%Y-%m-%d %H:%M:%S")

        src = os.path.abspath(xlsx)
        old = conn.execute("SELECT run_id FROM runs WHERE src_file=? AND chip=?", (src, chip)).fetchall()
        for (rid,) in old:
            conn.execute("DELETE FROM meas_raw WHERE run_id=?", (rid,))
            conn.execute("DELETE FROM meas_module WHERE run_id=?", (rid,))
            conn.execute("DELETE FROM runs WHERE run_id=?", (rid,))
        cur = conn.execute(
            "INSERT INTO runs(mode,chip,temp_c,src_file,run_ts,ingested_ts) VALUES(?,?,?,?,?,?)",
            (mode, chip, temp, src, run_ts, now_iso()))
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
                "INSERT INTO meas_module(run_id,step_order,group_disp,step_name,module_ids,sim_ids,current_ua,note)"
                " VALUES(?,?,?,?,?,?,?,?)",
                (run_id, s["order"], s["disp"], s["step_name"],
                 json.dumps(s["ids"]) if s["ids"] else None,
                 json.dumps(s["sim_ids"]) if s["sim_ids"] else None,
                 s["delta_ua"], s["note"]))
        conn.commit()
        return run_id, len(steps), temp, run_ts
    finally:
        wb.close()


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


def sim_lookup(conn, mode, ids, stage):
    """返回 (合计uA, 缺失ID列表, trim集合, tier集合)。"""
    total, missing, trims, tiers = 0.0, [], set(), set()
    found_any = False
    for mid in ids:
        rows = conn.execute(
            "SELECT current_ua,trim,tier FROM sim_current WHERE mode=? AND module_id=? AND stage=?",
            (mode, mid, stage)).fetchall()
        if not rows:
            missing.append(mid)
            continue
        found_any = True
        for cur, trim, tier in rows:
            total += cur
            if trim:
                trims.add(trim)
            if tier:
                tiers.add(tier)
    return (total if found_any else None), missing, trims, tiers


def module_names(conn, ids):
    names = []
    for mid in ids or []:
        row = conn.execute(
            "SELECT module_name FROM sim_current WHERE module_id=? AND module_name!='' LIMIT 1",
            (mid,)).fetchone()
        names.append(row[0] if row else f"ID{mid}")
    return " + ".join(names)


def export_xlsx(conn, out_path, all_runs=False):
    runs = conn.execute(
        "SELECT run_id,mode,chip,temp_c,src_file,run_ts FROM runs ORDER BY mode,chip,run_ts").fetchall()
    if not all_runs:
        latest = {}
        for r in runs:
            latest[(r[1], r[2])] = r  # 同 mode+chip 取 run_ts 最新（已按 run_ts 升序）
        runs = sorted(latest.values(), key=lambda r: (r[1], r[2]))

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
        "     其实测 delta 与仿真值都并入父模块组，组名显示为 如 26+28",
        "  5. 第二个及以后的 Init 段 = 锁定复验，忽略（Meas_steps 里有原始行）",
        "  6. 多 run 时默认每个 模式×芯片 取最新一次；--all-runs 可导出全部",
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
    for run_id, mode, chip, temp, _src, run_ts in runs:
        groups = conn.execute(
            "SELECT step_order,group_disp,step_name,module_ids,sim_ids,current_ua,note FROM meas_module"
            " WHERE run_id=? ORDER BY step_order", (run_id,)).fetchall()
        mode_in_sim = mode in sim_modes
        for order, disp, step_name, _mids, sim_ids_j, meas_ua, note in groups:
            sim_ids = json.loads(sim_ids_j) if sim_ids_j else None
            names = module_names(conn, sim_ids) if sim_ids else ""
            notes = [note] if note else []
            sim_pre = sim_post = None
            trims, tiers = set(), set()
            if sim_ids and mode_in_sim:
                sim_pre, miss_pre, t1, r1 = sim_lookup(conn, mode, sim_ids, "pre")
                sim_post, miss_post, t2, r2 = sim_lookup(conn, mode, sim_ids, "post")
                trims, tiers = t1 | t2, r1 | r2
                miss = sorted(set(miss_pre) & set(miss_post))
                if miss:
                    notes.append(f"仿真表缺ID: {miss}")
            elif sim_ids and not mode_in_sim:
                notes.append(f"仿真表无模式 {mode}")
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
    for run_id, mode, chip, _temp, _src, run_ts in runs:
        for r in conn.execute(
                "SELECT row_idx,seq_idx,kind,no_raw,mode_label,current_ma,delta_ma,temp_c,note"
                " FROM meas_raw WHERE run_id=? ORDER BY row_idx", (run_id,)):
            ws.append([mode, chip, run_ts, r[0], r[1], r[2], r[3], r[4],
                       rnd(r[5], 4), rnd(r[6], 4), r[7], r[8]])
    style_sheet(ws, [16, 8, 17, 6, 5, 8, 12, 24, 11, 10, 8, 30])

    # ---- Runs
    ws = wb.create_sheet("Runs")
    ws.append(["Run_ID", "Mode", "Chip", "Temp_C", "Run_TS", "Src_file"])
    for r in runs:
        ws.append([r[0], r[1], r[2], r[3], r[5], r[4]])
    style_sheet(ws, [7, 16, 8, 8, 17, 70])

    wb.save(out_path)


# ---------------------------------------------------------------- 命令

def walk_xlsx(root, skip_dirs):
    """递归列出 root 下所有 xlsx（跳过 skip_dirs 子树和 ~$ 锁文件）。"""
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = [d for d in dirnames if d not in skip_dirs]
        for f in sorted(filenames):
            if f.lower().endswith((".xlsx", ".xlsm")) and not f.startswith("~$"):
                yield os.path.join(dirpath, f)


def find_sim_workbook(root, config):
    """返回 (工作簿路径, tab名) 或 (None, None)。不认文件名，按表头内容全递归扫。"""
    if config.get("sim_workbook"):
        p = config["sim_workbook"]
        return (p if os.path.isabs(p) else os.path.join(root, p)), config.get("sim_sheet")
    skip = set(config.get("skip_dirs") or [])
    sim_sheet = norm(config.get("sim_sheet", "Current_data"))
    cands = []
    for path in walk_xlsx(root, skip):
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

    mode_map = config.get("mode_map") or {}
    skip = set(config.get("skip_dirs") or [])
    result_glob = config.get("result_glob", "Result*.xlsx")
    n_runs = 0
    for f in walk_xlsx(root, skip):
        if not fnmatch.fnmatch(os.path.basename(f), result_glob):
            continue
        d = os.path.basename(os.path.dirname(f))  # 模式 = 所在文件夹名
        mode = mode_map.get(d, d)
        run_id, n_steps, temp, run_ts = ingest_run(conn, f, mode, args.chip, config)
        n_runs += 1
        print(f"[实测] {mode} <- {os.path.relpath(f, root)}  run#{run_id}  {n_steps} 个模块组  "
              f"{temp if temp is not None else '?'}°C  {run_ts}")
    if n_runs == 0:
        print("[警告] 没有扫到任何 Result 文件")

    out = args.out or os.path.join(root, "Current_compare_pivot.xlsx")
    export_xlsx(conn, out, all_runs=args.all_runs)
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


def cmd_export(args):
    conn = open_db(args.db)
    export_xlsx(conn, args.out, all_runs=args.all_runs)
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

    e = sub.add_parser("export", help="从库导出 Excel")
    e.add_argument("--db", required=True)
    e.add_argument("--out", required=True)
    e.add_argument("--all-runs", action="store_true")
    e.set_defaults(func=cmd_export)

    args = ap.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
