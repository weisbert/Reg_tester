# Reg_tester

在真实芯片测试时，生成"电流逐级关闭"测试项的工具。

从"全开"状态出发，按某种顺序把一个个电流源逐级关掉，每关一级生成一个"测试项"
（一串寄存器写操作），供 ATE / 台架在每一级测电流。

## 输入

1. **netlist (veriloga)** —— Cadence schematic TOP 层导出，含 PIN 脚信息与内部连接，
   提供电流相关控制信号及其在结构中的层级。
2. **控制信号 Excel** —— 控制信号名 → 寄存器地址 + bit 位 + 值。

> 真实的寄存器表 / netlist 属机密，默认被 `.gitignore` 挡在仓库外。

## 安装

```powershell
pip install -r requirements.txt
```

## 工具

### `explore_excel.py` —— 探查 Excel 结构

在不知道寄存器表长什么样时，先摸清结构（sheet、表头行、合并单元格、各列类型），
便于本地复刻一份等价表来开发。只依赖 `openpyxl`。支持 `.xlsx` 和 `.xlsm`
（macro-enabled，宏不影响读数据）。

```powershell
# 控制台切 UTF-8，避免中文乱码（每个窗口执行一次）
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

python explore_excel.py "寄存器表.xlsm"                       # 控制台看结构预览
python explore_excel.py "寄存器表.xlsm" --index idx.json        # ① 超精简：整体形状（先发这个）
python explore_excel.py "寄存器表.xlsm" --schema s.json --sheet "某表"  # ② 单个 sheet 详细结构
python explore_excel.py "寄存器表.xlsm" --schema s.json --max-sheets 3  # 前 3 个 sheet 详细结构
python explore_excel.py "寄存器表.xlsm" --schema-dir schemas            # ③ 每个 sheet 各一个小文件（一次收齐）
python explore_excel.py "寄存器表.xlsm" --sheet <某寄存器 sheet> --rowdump 100:235 --dump rows.json  # 抓某表某段完整内容(紧凑,裁空列)
python explore_excel.py "寄存器表.xlsm" --dump reg_dump.json    # 完整内容导出（会很大）
python explore_excel.py "寄存器表.xlsm" --formulas             # 值是宏/公式算的、读成空时改看公式
```

体积从小到大，导出时都会打印字节数：

- `--index`：**超精简**，每个 sheet 只留 名字/尺寸/行列/合并数/表头。sheet 再多也很小，先发这个看整体形状。
- `--schema`：**结构骨架**，每个 sheet 出 每列类型+去重样例+头几行样本；sheet 多时配 `--sheet "名"` 或 `--max-sheets N` 压体积。
- `--dump`：把每个格子都导出，文件会很大。

> `--index` 不给路径则直接打印到控制台；`--schema` 同理。

### `extract_ports.py` —— 抽取指定模块的端口 I/O

从 `vh_extract` 生成的结构化 Verilog(-A) 网表里，只抽出关心的 sub-top block 的
输入/输出端口（名字/方向/位宽），跳过内部 `wire`，用来定位需要哪些控制信号。
只依赖标准库。

```powershell
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

python extract_ports.py "路径\netlist.vh"                 # 默认抽三个目标 DCO/LO 模块
python extract_ports.py "路径\netlist.vh" --json ports.json  # 同时导出 JSON
python extract_ports.py "路径\netlist.vh" --list            # 只列出文件里所有模块名
python extract_ports.py "路径\netlist.vh" --modules A,B,C    # 指定模块
```

默认目标模块名是项目私有配置（见脚本内 `DEFAULT_TARGETS`，或用 `--modules A,B,C` 显式指定）——
一般是若干 DCO/LO 链的顶层 block 加一个时钟 MUX block。
输出把端口分成 控制输入 / 输出 / 模拟 inout / 电源地 四类，控制输入即候选控制信号。

**连接模式 `--connections`**（命名多轮迭代后不可信，靠连接定 ground truth）：抽出
端口 + wire + instance(例化) + 连线，并建 `net_index`（每根网络接到哪些顶层端口/实例引脚），
用于追踪 `EN → buffer实例 → 输出` 的真实通路。带自诊断：模块体里没被识别为
端口/wire/assign/instance 的残留会进 `unparsed`——**`✓ 无残留` 才算可信**。

```powershell
python extract_ports.py "netlist.vh" --connections --json conn.json            # 抽连接(完整,带缩进,较大)
python extract_ports.py "netlist.vh" --connections --compact --json conn.json  # 紧凑无损(约 1/5 体积)
python extract_ports.py "netlist.vh" --body <MODULE> --head 80                 # 打印模块体原文(核对语法)
```

- `--compact`：无损压缩。去掉 `net_index`（可由 instances+ports+assigns 重建）、每连接的 `pos`/`nets`
  （`pos`=数组顺序、`nets`=从 `expr` 解析，均可重建），连接表示为 `[pin, expr]`，紧凑排版。约 5× 减小。
  太大还可加 `--modules <单个模块>` 分模块导出。

**层级/向上追踪**（目标模块怎么延到文件顶层）：

```powershell
python extract_ports.py "netlist.vh" --tree                  # 层级树：根(文件顶)一层层到目标模块
python extract_ports.py "netlist.vh" --uptrace --json up.json # 目标每个端口向上追到文件顶层引脚
```

- `--tree`：打印 根模块(文件顶) → ... → 目标模块 的例化路径，输出很小。
- `--uptrace`：目标模块每个端口，顺着父层例化的连线（并跟随 `assign` 别名/level-shift 透传）
  一直追到文件顶层引脚，或标出在哪层被内部消耗。
- 边界：本文件的"顶"是抽取起点（网表文件顶层模块）；再往上到芯片真 TOP / 寄存器配置位需靠 Excel 或更高层网表桥接。

### `excel_lookup.py` —— 按控制信号 list 反查寄存器 Excel

给一份控制信号名清单（`.json` 取所有 `reg_net` / `.txt` 每行一个 / 逗号串），
在寄存器簿每个 sheet 每个单元格里做子串匹配，命中就把整行（裁剪后）抓出来——
看每个控制信号对应的 地址/bit/值/寄存器 怎么写。输出只含关心的信号，体积小。

```powershell
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
python excel_lookup.py "寄存器.xlsm" --signals terms.txt --json hits.json
python excel_lookup.py "寄存器.xlsm" --signals a,b,c --sheets RegMapDesign,regmap,Topout
```

自动去掉信号名的位宽后缀 `[3:0]`；`--formulas` 读公式原文；`--max-hits` 限每信号命中行数。

### 工程包 `--project`（P1，推荐入口）

一颗芯片 = 一个 gitignore 的工程包 `projects/<chip>/`，`project.json`（`project/2`）吃下**全部芯片专属配置**
（原先散在 `private/tool_config/` 的 4 个 config + 手工 `aliases.json`），派生物按固定名放同目录。四个建库脚本都支持
`--project projects/<chip>`：配置读 `project.json` 对应段、IO 按 `artifacts` 从包内解析——**代码零真实字面量**。
完整字段见 `SCHEMAS.md` 的 `project/2`。

```powershell
python make_mock_regmap.py --project projects/<chip>    # regbook 段 + matching.alias → signal_reg_map.json
python build_regmap.py     --project projects/<chip>    # regbook 分组前缀/name_xform → regmap.json
python build_flowgraph.py  --project projects/<chip>    # flowgraph_rules + netlist.expand_submodules → flowgraph.json
python regtool.py --serve  --project projects/<chip>    # GUI 读同一个包
```

不带 `--project` 时各脚本回退旧默认（`private/tool_config/` + `private/adpll/`），两路输出**逐字节一致**。下面各脚本文档为传统单步用法。

### `make_mock_regmap.py` —— 解析控制信号→寄存器 + 本地复刻寄存器 Excel

用抓回的 目标寄存器 sheet 行(`--rowdump` 结果) + 控制信号 list + 变体映射(alias)，把每个控制信号
解析到 寄存器/绝对地址(base+offset)/bit/默认值/关断值，并生成一个**结构一模一样**的
nManager 布局 `.xlsx`（本地开发用，不再依赖真文件）。脚本本身不含真实信号名，只读 private/ 输入。

```powershell
python make_mock_regmap.py --rows pll_rows.json --signals control_signals.json ^
  --aliases aliases.json --schema <REG>.schema.json ^
  --out-xlsx <REG>_mock.xlsx --out-map signal_reg_map.json
```

### `build_regmap.py` —— 寄存器侧 adapter：signal_reg_map.json → 规范 regmap.json

把阶段一的 `signal_reg_map.json` 规范化成带版本号的中间层 `regmap.json`（GUI inspector +
序列生成器 RMW 都读它）。核心增量：给每个控制信号补 **WL / WLT 平行寄存器字段**（BT/WL
两套并行寄存器组都要测）。孪生匹配抗命名漂移：寄存器按基名分组（去 `BT_/WL_/WL1_` 前缀），
信号字段 → 孪生寄存器里**同 bit 位**的字段 = 平行字段（bit 位是主键，名字变换只做校验）。
只依赖标准库；脚本不含真实信号名/地址，只读 `private/`、写 `private/`。

```powershell
python build_regmap.py                         # 默认 private/adpll/signal_reg_map.json → regmap.json
python build_regmap.py --print                 # 额外打印 reg_net → BT/WL/WLT addr@bit 核对表
python build_regmap.py --config regmap_rules.json   # 覆盖默认分组/前缀规则（换项目）
```

输出 `regmap.json`：`signals[]` 每条含 `id`(=reg_net，即 flowgraph 控制脚的引用键)、`match`、
`category`、`shared`、`warn`、`active_high`、`off_value`、`variants{BT|WL|WLT|COMMON}`
（每 variant 带 `addr/offset/bit/default/reset`）。`single_copy` 信号只有 `COMMON`。

### `build_flowgraph.py` —— 网表侧 adapter：conn.json → 规范 flowgraph.json

把 `extract_ports.py --connections` 抽出的 sub-top 连接转成带版本号的 `flowgraph.json`
（GUI 信号流图 + inspector + 序列生成器都读它）。要点：

- **节点分层**：module 分组框 → composite 黑盒(buffer bank) → 子节点。子节点两种来源：
  ①**真实展开**（有该黑盒内部连接时，`--expand` 递归展开真实叶子实例，`inferred:false`，节点用真实实例名）；
  ②**推断合成**（无内部连接时回退 channel synthesis，`inferred/provisional:true`）。
  不透明 divider/route primitive（内部无数据，不合成）；logic = 控制域(默认折叠)。
- **子模块递归展开**(`--expand <conn.json>`)：黑盒有真实内部连接时用**同一套 type_rules** 展开——
  边界端口并入父命名域（控制脚经父网解析寄存器，与推断态结果一致）、纯内部网留子命名域产内部边、
  叶子按类型分 buf/mux/div/inv/glue。glue 标准单元(ND/NR/INV)即使输入网连寄存器也只记 `gate_in`
  不产 off_control（非耗电器件）。派生使能（经内部逻辑生成，如低功耗 rx 使能）以内部边呈现。
  缺省自动加载 `private/adpll/expand_conn.json`（若存在）。
- **控制脚挂寄存器**：控制脚的**驱动网**(`ls_` 网)经 Logic 追回 sub-top 原始端口，再经
  `regmap.drives` 反查信号。**只信连接不信名字**——寄存器位由驱动网决定，不由引脚名决定。
- **off_controls**：类别属“电流门”的 enable 脚，`active_high` 缺失时按“高有效/关=0”兜底并标
  `polarity_inferred`，供序列生成器逐级关。
- **glue 门控前推**(`glue_gate_propagation`)：经内部 glue 逻辑门控的寄存器使能（如某低功耗 buffer
  的使能经内部 ND2+INV 逻辑门链门控该 buffer），沿数据边前推到它最终门控的耗电器件当
  `off_control`（标 `via_glue`）。用寄存器**自身 off_value**——写寄存器到 OFF 即关断整条通路，与 glue
  内部极性无关（AND 型门控）。这样 shutdown 序列不会漏关经组合逻辑门控的支路。
- **差分合并**：仅当 p/n 两相同一驱动节点才合成一条边（ADC 真实反相器/两个 buf 出各自保留）。
- **DCO 源头边**：DCO 核（`device=dco`）的裸差分输出网（`diff_bare_pn`：tank/gm/buf p/n）标为 output，
  这样"DCO 核→取该振荡网的源头 buffer"边能正确生成（否则多端点无驱动会被判方向不明而丢弃）。
- **规则全配置化**(`--config`)：换项目=改配置(前缀/后缀/asserted_edges)不改代码。

```powershell
python build_regmap.py            # 先出 regmap.json（flowgraph 要读它做 drives 反查）
python build_flowgraph.py         # 默认 conn.json + regmap.json → flowgraph.json
python build_flowgraph.py --print # 打印节点树 + off_controls + 未解析诊断
```

`flowgraph.json`：`nodes[]`（含 `pins/controls/off_controls/reg_touch`）、`edges[]`（差分合并 +
方向 + 跨模块）、内嵌 `signals{}`（引用式，inspector 单文件一跳）、`stats` + `diagnostics`
（未解析控制脚 / 未配对输出 / 隐藏计数，供人工核对）。**buffer bank 子节点**：有内部连接时为
真实展开（`expanded:true`，节点=真实实例名）；无则为推断态（`provisional_children:true`）。
⚠ 推断态→真实展开后**节点 id 变化**（真实实例名 ≠ 推断 token），引用旧 id 的 modes/layout 需在
GUI 里对真实图重指——一个推断 token 常对应 0、1 或多个真实器件，不做自动映射（避免错关）。

### `gen_testcase.py` —— 无头序列生成器 + 渲染器（阶段二 M3 核心）

读 `flowgraph.json` + `regmap.json` + 一个模式定义（`modes/1`），生成"电流逐级关闭"测试序列
（唯一事实来源 `testcase/1`），再渲染 `ate.txt`（交付格式）/ `debug.html`（designer 看）。
**累积逐级关闭**语义：从"全开基线"出发，按激活通路**反向拓扑**（末端 buffer → DCO 源头）逐级 read-modify-write
关 enable，每步先发增量写、再测总电流，相邻步电流差 = 该级模块功耗。共用位/极性/未覆盖门都有告警。
只依赖标准库；脚本零真实信号名/地址（全从 JSON 读）。算法权威定义见 `SCHEMAS.md`。

```powershell
python gen_testcase.py --project projects/<name> --mode <MODE_ID>          # 写 testcases/<mode>.{json,ate.txt,debug.html}
python gen_testcase.py --project projects/<name> --mode <MODE_ID> --print  # 打印 ate.txt 到控制台
python gen_testcase.py --flowgraph fg.json --regmap rm.json --mode-file m.json --out-ate a.txt --out-html d.html
```

### `regtool.py` —— GUI 信号流工具启动器（阶段二 M2）

纯客户端 GUI（`webapp/`），两种启动方式，前端同一套代码：

```powershell
python regtool.py --serve  --project projects/<name>              # http.server + 工程 REST（主用），默认 :8765
python regtool.py --serve  --project projects/<name> --open       # 顺便开浏览器
python regtool.py --bundle --project projects/<name> --out out.html  # 打成自包含单 HTML（离线/黄区应急）
```

GUI 能力：分层信号流图（buf=三角 / mux=梯形 / div=方框 / DCO=振荡器；三 sub-top 成可折叠分组框，
双击黑盒展开/折叠）、缩放平移搜索小地图、拖拽/框选/隐藏/翻边/备注 + 撤销重做（落 `layout.json`）、
inspector 侧栏（点节点看控制信号→寄存器/地址/bit/默认/关断值，可改基线字段）、模式编辑、
序列生成预览 + 导出 `ate.txt`/`debug.html`。

**描通路（模式主定义方式）**：按**信号流顺序**点节点（DCO 源头→输出末端）= 一条链,同时给出
**激活集**（哪些 block 开）和**关断顺序**（末端→源头,自动逆序,写进 `mode.flow_path`,派生 `enabled_nodes`
与 `order.manual`）。不必依赖自动拓扑/边方向——**用户的通路知识直接定义模式**（模拟电路里信号流方向
网表推不可靠,故由人描）。空 flow_path 时回退老方式（手工激活 + 录制/自动反向拓扑）。
纯原生 JS + SVG + 自研分层布局（无 npm 依赖，符合红区轻依赖纪律）。中间层 schema 见 `SCHEMAS.md`。

**匹配标签（P2，serve 模式）**：网表控制信号名与寄存器字段名经硅迭代漂移，自动匹配只覆盖 exact/case。
"匹配"侧栏列出每个信号的状态（已匹配 / 待处理 / 逻辑推导），对不上的用字段选择器点选真实字段（或标为
逻辑推导）；「应用 & 重建」把确认结果写回 `project.json` 的 `matching.alias`，后端 `POST /api/matching`
按序重跑 `make_mock_regmap → build_regmap → build_flowgraph`（stdlib subprocess），刷新图与寄存器。
**人工确认关卡**保留：对真硅写寄存器的模糊匹配必须人点头，不做自动猜。bundle 模式只读。

**工程标签 / 建库向导（P2.2，serve 模式）**：从零建库不必手写 `project.json` 或先跑一串 CLI——
`--project` 指向一个**不存在的目录**即自动脚手架空工程起 GUI，`bootstrap.needs_setup=true` 时前端自动切到
"工程"标签，引导走 air-gap 建库六步：

1. **芯片配置**：填 `name`/`root_module`/`base_address`/`sheet_name`、目标模块 `target_modules`、展开子模块
   `expand_submodules`、`column_map`（字段名→Excel 列号）、可选高级 `flowgraph_rules`（留空用通用默认）
   → `POST /api/project/config` 写 `project.json`（`artifacts` 补默认固定名）。
2. **导入 netlist**（填本机绝对路径）→ `POST /api/import/netlist`：后端跑 `extract_ports --connections --compact`
   出 `conn.json`（+ `--expand` 出 `expand_conn.json`），再抽 `--uptrace` **自动生成 `control_signals` 候选**
   （top_pin 输入控制脚按后缀启发式分 primary/secondary；category/desc 需人工确认，经内部逻辑门控的漏网脚需手工补）。
3. **导入 Excel**（路径 + sheet + 行区间）→ `POST /api/import/excel`：后端跑 `explore_excel --rowdump` 出 `pll_rows.json`。
4. **控制信号**：候选填入文本框，人工确认/补全后 `POST /api/control_signals` 写 `control_signals.json`。
5. **建库** → `POST /api/build`：跑 `make_mock_regmap → build_regmap → build_flowgraph`（stdlib subprocess），载入图。
6. **导出工程配置** → `GET /api/project/export`：`project.json` + `control_signals` + 全部 `modes` 打成一段文本，
   复制发回本地归档（air-gap 出口）。

抽取脚本读 `.xlsm` 需 `openpyxl`（黄区可装；红区 stdlib-only 时用离线包）；`extract_ports`/建库三件套均 stdlib。
路径参数只在本机（`--serve` 绑 `127.0.0.1`）读用户指定的本地文件，subprocess 以列表参数调用（无 shell 注入）。
**人工确认关卡**同样保留：control_signals 候选与匹配都要人过目。bundle 模式只读（建库需 `--serve`）。

> **工程数据在 `projects/<name>/`（gitignore）**：`project.json`/`flowgraph.json`/`regmap.json`/`layout.json`/
> `modes/*.json`/`testcases/` + 输入 `conn.json`/`expand_conn.json`/`control_signals.json`/`pll_rows.json`。
> 含真实信号名与地址，绝不进公开仓库。

## 状态

阶段一（需求对齐 + 收集 + 本地复刻）完成。阶段二全部完成：
- **M1**：`build_regmap.py`（补 WL/WLT 平行字段）+ `build_flowgraph.py`（conn.json → 分层信号流图）。
- **M2/M3**：`regtool.py`（serve/bundle 双启动 + 工程 REST）+ `webapp/`（图编辑 + inspector + 模式编辑 +
  序列生成）+ `gen_testcase.py`（无头生成器 + `ate.txt`/`debug.html` 渲染器）。
- **M4**：多路对抗式验证——Python↔JS 生成器 24/24 组合逐字节一致、RMW 数值独立复算金标准（BT/WL/WLT）、
  关闭顺序端→源头、REST 路径穿越/坏输入加固、前端 XSS 面收敛、jsdom 端到端渲染/交互全过。

阶段三：**P0** 通用递归子模块展开（黑盒真连接取代命名配对推断）；**P1** `project/2` 工程包=唯一芯片真相
（配置沉数据侧、代码零真名，`--project` 读它）；**P2.1** GUI 半自动信号→寄存器匹配；**P2.2** 建库向导
（GUI 里 netlist+Excel 导入→自动抽取→控制信号候选→建库→导出配置，见上"工程标签"）。
