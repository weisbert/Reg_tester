# Reg_tester — 中间层 JSON schema & 序列生成算法（权威契约）

本文件是**通用契约**。**所有示例一律用脱敏占位**——地址写 `<BASE>+0xNN`、寄存器名 `REG_A/REG_B`、
节点 id `MOD/inst_x::leaf_y`、信号名 `sig_*`——真实信号名/地址/模块名只存在于 gitignore 的
`private/`、`projects/` 里，绝不进本文件、也绝不进公开仓库。

GUI（`webapp/`）、无头生成器（`gen_testcase.py`）、渲染器都只认这里定义的 schema，
换项目不改代码——只换 `private/` 里的数据 + `project.json` 里的规则覆盖。

四个中间 schema 全部带 `schema_version`，改结构必升版本号。

```
netlist ─(build_flowgraph)→ flowgraph/1 ─┐
excel   ─(build_regmap)   → regmap/1    ─┼→ GUI 编辑 → modes/1 ─(gen_testcase)→ testcase/1 ─(renderer)→ ate.txt / debug.html
                                          └───────────────────────────────────┘
```

约定：`<BASE>` = 目标模块域的块基址（真实值在 regmap 数据里）；`0xNNNN` = 某个 16-bit 值占位；
`0xNN` = 某个字节偏移占位；`sig_*` / `REG_*` / `MOD` / `inst_*` / `leaf_*` = 占位名。

---

## 0. `project/2`（工程包 = 唯一芯片真相，P1）

一颗芯片 = 一个 gitignore 的工程包目录 `projects/<chip>/`。`project.json` 吃下**全部芯片专属配置**
（原先散在 `private/tool_config/` 的 4 个 per-script config + 手工 `aliases.json`），派生物按固定名放同目录。
所有建库引擎 `--project projects/<chip>` 从这里读配置与 IO；**代码零真实字面量**（模块名/基址/前缀/别名全在数据侧）。

```jsonc
{
  "schema_version": "project/2",
  "name": "<chip>",
  "netlist":  { "root_module": "TOP", "target_modules": ["MOD_A","MOD_B"],   // extract_ports
                "expand_submodules": ["MOD_A_BUF"] },                        // build_flowgraph 递归展开
  "regbook":  { "base_address": "<BASE>", "sheet_name": "REGS",             // make_mock + build_regmap
                "column_map": {"reg_name":0,"offset":3,"field_name":9,"bit":10,"default":12,"comment":14,"...":0},
                "reg_group_prefixes": [["G1_","G1"],["G0_","G0"]], "primary_group":"G0", "common_group":"COMMON",
                "name_xform": {"G1": [["sig_","g1_sig_"]]} },
  "flowgraph_rules": { "type_rules": [["(?i)buf","primitive","buf","prim"]], // build_flowgraph
                       "module_bands": {"MOD_A":"A"}, "strip_prefix": ["mod_"],
                       "known_cross_edges": [], "crossmodule_net_blocklist": [] },
  "matching": { "alias": {"sig_x":"REG_field_y"}, "logic_derived": [], "unresolved_pending": [] }, // P2 GUI 写回
  "artifacts": { "conn":"conn.json","expand_conn":["expand_conn.json"],"control_signals":"control_signals.json",
                 "pll_rows":"pll_rows.json","signal_reg_map":"signal_reg_map.json",
                 "regmap":"regmap.json","flowgraph":"flowgraph.json" }
}
```

引擎映射：`extract_ports --project`→`netlist`；`make_mock_regmap --project`→`regbook`(base/sheet/column_map)+`matching`(alias)；
`build_regmap --project`→`regbook`(分组前缀/name_xform)；`build_flowgraph --project`→`flowgraph_rules`+`netlist.expand_submodules`。
GUI（`regtool --project`）只从 `project.json` 取 `name`，其余按固定名读 `flowgraph.json/regmap.json/layout.json/modes/`。
无 `--project` 时各脚本回退旧默认（`private/tool_config/` + `private/adpll/`），两路输出逐字节一致。

---

## 1. `flowgraph/1`（M1 产物，`build_flowgraph.py`）

信号流图。GUI 渲染 + inspector + 序列生成器都读它。关键字段（完整见产物）：

- `reg_base` `reg_groups`（并行寄存器组名数组，如 `["G_A","G_B","G_C"]`）`module_tags` `module_bands`
- `nodes[]`：`id`（唯一，形如 `MOD/inst_buf::leaf_rxbuf`）、`kind`(module|composite|primitive|inferred)、
  `device`(dco|div|buf|mux|inv|logic|route|blackbox|group)、`symbol`(oscillator|box_divN|triangle|
  trapezoid|triangle_bubble|ctrl_block|pass|group)、`parent`、`children[]`、`band`、`reg_group_default`、
  `inferred`、`hidden_default`、`opaque_blackbox`、`expandable`、`control_domain`、
  `pins[]`（含 `signal_ref`/`net`/`role`/`resolved`）、
  **`controls[]`** `{pin, signal_ref, role, shared, lane}`（`role` 取 enable|sel|tune|data_in|…）、
  **`off_controls[]`** `{pin, signal_ref, off_value, active_high, polarity_inferred, lane}` ← 序列生成器逐级关的门、
  `reg_touch[]`。
- `edges[]`：`id`、`scope`、`from{node,pin}`、`to[{node,pin}]`、`differential`、`direction`、
  `cross_module`、`provenance`(net|asserted)、`kind`(lo|clk|data)、可选 `pair`（差分反相脚）。
- `signals{}`：`signal_ref → {drives[], bound_pins[], banks{}, category, shared, ...}`（inspector 单跳引用）。
- `stats` `diagnostics`（`uncovered_off_gates`=真门但无可挂节点、序列不会自动关，需人工补）。

## 2. `regmap/1`（M1 产物，`build_regmap.py`）

信号→寄存器。`gen_testcase` RMW + inspector 读它。

- `base_addr` `reg_groups` `primary_group` `common_group`(如 `"COMMON"`)
- `registers[]`：`reg_name`、`reg_group`、`offset`、`addr`、`reset`(16-bit 上电值)、`width`、`fields[]`。
- `signals[]`：`id`(=`reg_net`=flowgraph 的 `signal_ref`)、`category`、`shared`、`active_high`、`off_value`、
  `single_copy`、`drives[]`、**`variants{<各 reg_group>|COMMON}`**——每 variant `{reg_group,reg_name,offset,addr,
  reset,width,field_name,bit,attr,default,comment}`。
  - `single_copy`/COMMON-only 信号：所有 reg_group 共用 `variants.COMMON`。
  - 取字段：`variant = signal.variants[mode.reg_group] or signal.variants["COMMON"]`。

`bit` 形如 `"4"`（单位）或 `"11:10"`（`hi:lo`）。`default`/`reset` 形如 `"0x0"`/`"0xNNNN"`。

---

## 3. `modes/1`（GUI 产出，`projects/<p>/modes/<id>.json`）

一个"模式" = 一条被激活的 LO 通路 + 选定寄存器组 + 基线字段覆盖 + 关闭顺序。**示例全为占位名。**

```jsonc
{
  "schema_version": "modes/1",
  "id": "MODE_EXAMPLE",             // 文件名去 .json，唯一，且=服务端/GUI 的键（文件名权威）
  "name": "示例模式",
  "description": "某接收通路：DCO → 一级分频 → 输出 buffer → LO",
  "reg_group": "G_A",              // 本模式落哪套并行寄存器组（取自 regmap.reg_groups 之一）
  "targets": ["MOD/leaf_lo_out"],  // 可选：本通路的 LO 输出节点（供高亮/回溯）
  "flow_path": [                   // **描通路**（GUI 主定义方式）：按信号流 源→末端 顺序的节点链。
    "MOD/inst_dco",               //   派生：enabled_nodes = flow_path；order.manual = flow_path 逆序（末端先关）。
    "MOD/inst_bufblk::leaf_div",  //   一条链同时定义"哪些 block 开"+"按什么顺序关"，无需自动拓扑。空=用老方式(手工激活/录制)。
    "MOD/inst_bufblk::leaf_rxbuf"
  ],
  "enabled_nodes": [               // 处于"开"状态的节点 id（描通路时 = flow_path；也可手工激活）
    "MOD/inst_dco",
    "MOD/inst_bufblk::leaf_div",
    "MOD/inst_bufblk::leaf_rxbuf"
  ],
  "mux_sel": { "MOD/inst_bufblk::leaf_mux": 0 },   // MUX 选择（0=上输入/1=下输入约定）；会写进该 mux 的 sel 字段
  "baseline": {                    // 显式基线字段覆盖：signal_id → 整数值（覆盖 variant.default / 门默认 / mux_sel）
    "sig_itune": 8,                //   en 门由 enabled_nodes 自动置开/关；这里放 ictrl/tune/ct/mt 等手改值
    "sig_buf_ictrl": 2
  },
  "order": {
    "mode": "auto",                // auto = 激活通路反向拓扑；manual = 用户录制顺序
    "manual": []                   // mode=manual 时：关闭节点 id 的有序列表（节点粒度）
  },
  "extra_writes": [                // 口子：模式级额外写（LDO/时钟等域外前置，工具默认不管）
    // { "addr": "<BASE>+0xNN", "value": "0xNNNN", "note": "..." }   // VALUE 会被规范化成 0x+大写4位
  ],
  "notes": ""
}
```

## 4. `testcase/1`（`gen_testcase` 产物 = 唯一事实来源，渲染器只读它）

下面示例**列出全部字段**（非节选），值为占位。

```jsonc
{
  "schema_version": "testcase/1",
  "mode": "MODE_EXAMPLE",          // = 模式 id（缺失则 null）
  "mode_name": "示例模式",          // = 模式 name（缺失则 null）
  "reg_group": "G_A",
  "base_addr": "<BASE>",
  "order_mode": "auto",
  "baseline": {
    "note": "建立全开起始态（激活通路开、其余门关、tune/ictrl 取基线值）",
    "writes": [
      { "addr": "<BASE>+0x54", "reg": "REG_A", "value": "0x001D", "reset": "0x000D",
        "fields": [ { "signal": "sig_dco_en", "bit": "4", "value": 1, "role": "enable", "on": true },
                    { "signal": "sig_itune",  "bit": "9:6", "value": 8, "role": "override" },
                    { "signal": "sig_mux_sel","bit": "5", "value": 0, "role": "mux_sel" } ] }
    ]
  },
  "steps": [
    {
      "index": 1,
      "off_node": "MOD/inst_bufblk::leaf_rxbuf",
      "off_label": "leaf_rxbuf",
      "device": "buf",
      "measure": "关此级后测总电流",
      "gates": [ { "signal": "sig_rxbuf_en", "pin": "pin_rxbuf_en",
                   "off_value": 0, "shared": false, "polarity_inferred": false, "collateral_nodes": [] } ],
      "writes": [ { "addr": "<BASE>+0x54", "reg": "REG_A", "value": "0x001C", "prev": "0x001D",
                    "fields": [ { "signal": "sig_rxbuf_en", "bit": "0",
                                  "before": 1, "after": 0, "role": "enable" } ] } ],
      "warnings": [],
      "note": null                 // 非 null 时：本级门已被前面共用位提前关掉（仍是测量点）
    }
  ],
  "extra_writes": [],
  "warnings": [],                  // 顶层告警（未解析信号 / 相邻同级顺序需人工确认 / manual 未覆盖 / MUX 多 sel …）
  "diagnostics": {
    "uncovered_off_gates": [ /* {signal, note} */ ],
    "shared_collateral": [ /* {step, node, signal} */ ]
  },
  "stats": { "baseline_regs": 1, "steps": 1, "gates_off": 1 }
}
```

字段说明补充：`baseline.writes[].reset`、`baseline.writes[].fields[].on`、`steps[].device`、`steps[].note`、
`gates[].polarity_inferred`、顶层 `stats`/`mode_name` 都是渲染器依赖的正式字段。可空字段缺失时值为 `null`
（Python 与 JS 生成器对此**一致**：JS 侧把 undefined 归一成 null）。

---

## 5. 序列生成算法（`gen_testcase.py` 与 GUI `webapp/generator.js` 必须逐字节一致）

**输入**：`flowgraph/1`、`regmap/1`、`modes/1`。**输出**：`testcase/1`。

**术语**：
- `variant(sig) = regmap.signals[sig].variants[mode.reg_group] or [...].variants["COMMON"]`。
- `on_value(gate)`：enable 门的"开"值 = `1 if active_high else 0`（`active_high` 缺失按高有效兜底、标
  `polarity_inferred`）。`off_value` 取 `off_controls` 里给的。
- 位域写入 `set_field(word16, bit, val)`：`bit="hi:lo"` → `width=hi-lo+1`，`mask=((1<<width)-1)<<lo`，
  `word = (word & ~mask) | ((val & (mask>>lo)) << lo)`；`bit="n"` 视作 `n:n`。

**Step A — 收集门**：遍历 `flowgraph.nodes[].off_controls`，建：
- `gate_nodes[signal] = {所有把该 signal 作为 off_control 的节点 id}`；
- `node_gates[node] = [该节点的 off_control 门…]`。

**Step B — 判每个门信号的开/关（共用位：开压倒关）**：
`signal_on(sig) = any(n in mode.enabled_nodes for n in gate_nodes[sig])`。
> 一个信号驱动多个节点时（硅迭代遗留的共用位），只要其中一个节点在激活集里就得开——
> 否则会关掉激活通路。这类信号标 `shared` 并在受影响处出 collateral 警告。

**Step C — 基线寄存器映像**：
1. `touched_addrs` = 所有门信号 variant.addr ∪ `mode.baseline` 信号 ∪ `enabled_nodes` 上任何 control 的信号。
2. 预置：每个 touched 信号（含激活节点的 tune/tail 控制）建立 reset 映像 `int(variant.reset,16)`；
   不解析（无寄存器的频率/模式信号）静默跳过——它们不是电流门。
3. 对每个门信号：`set_field(image, variant.bit, on_value if signal_on(sig) else off_value)`。
4. `mux_sel`：对每个 `node:val`，若该 mux 节点有**唯一** `sel` 控制且能解析，`set_field(image, sel.bit, val)`。
5. 叠加 `mode.baseline`：对每个 `sig:val`，`set_field(image, variant.bit, val)`（手改值压倒门默认与 mux_sel）。
6. baseline.writes = 每个 touched addr 一条 `{addr,reg,value=hex4,reset=hex4,fields:[被本模式显式设置的字段…]}`。

**Step D — 关闭步骤（累积语义）**：
- `active_gate_nodes` = `enabled_nodes` 里、有至少一个"基线为开"门的节点。
- **排序**（确定性，全有兜底键）：`sort_key = (-shutdown_rank, -enabled_index, id)`。
  - `shutdown_rank` = 沿 `edges` 的"距 DCO 源头深度"（越大越靠末端）；边缺失时用器件类基座
    `{dco:0, logic:0, div:1, mux:2, buf:2, inv:2, blackbox:2, route:3, group:0}`，取 `max(基座, 1+max(前驱 rank))`。
  - `enabled_index` = 节点在 `mode.enabled_nodes` 里的位置。**同 rank 时靠它兜底**（设计者按源→末端录入 →
    越靠后越靠末端 → 越先关），避免 composite 内缺 leaf→leaf 边时退化成 id 字典序把上下游关反。
  - 仍有相邻同 rank 时，输出一条 `warnings` 提示"该处先后由拓扑无法判定、按录入序兜底、请人工确认"。
  - `manual` 模式：先按 `order.manual` 排，未覆盖的激活节点按上面 sort_key 追加并告警。
- 维护 `reg_image`(=baseline 副本) 和 `signals_off`。逐节点：对每个"基线为开"门 `g`——
  - 若 `g.signal in signals_off`：跳过（已被共用位提前关）→ 记 `shared_collateral`；
  - 否则 `set_field`→off_value，映像变了则记该 addr 新值到本步 writes；`signals_off.add`；
    该 signal 若还驱动别的激活节点，出 collateral 警告。
  - 一步内落同一寄存器的多门 → 合并成一次写。空写步仍是测量点，`note` 标注。
- 每步 = 一个测量点：**先发本步增量写 → 再测总电流**；相邻步总电流差 = 该级模块功耗。

**确定性**：同输入 → 同输出（排序键全有兜底，无时间/随机）。产物 idempotent。

## 6. 渲染器（`testcase/1` → 文本）

- **`ate.txt`（交付格式）**：数据行 `ADDR VALUE MODULE [; 行内注释]`（`ADDR`=`0x`+大写8位、`VALUE`=`0x`+大写4位，
  **含 extra_writes 的 VALUE 也规范化**；`MODULE`=关闭的节点/模块名）。`MODULE` 后可跟以 ` ; ` 起的行内注释；
  以 `#` 起头的整行为纯注释（baseline 段、步号、字段 before→after、共用位警告都在注释里）。
  精确列样式在 M3 拿首发样例与用户逐行钉定，模板可配。
- **`debug.html`（designer 看）**：每步展开到字段级（信号/寄存器/bit/前后值）+ baseline/extra_writes/未覆盖门表；
  所有字段值经 HTML 转义（无注入）。Python 版最全；GUI bundle 内置精简 JS 版供离线下载。
