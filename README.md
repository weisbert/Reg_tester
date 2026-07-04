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
python explore_excel.py "寄存器表.xlsm" --sheet REG_SHEET --rowdump 100:235 --dump rows.json  # 抓某表某段完整内容(紧凑,裁空列)
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

默认目标模块：`CHAIN_TOP_A`、`CHAIN_TOP_B`、`CLK_MUX_C`。
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
- 边界：本文件的"顶"是抽取起点（如 `top_core`）；再往上到芯片真 TOP / 寄存器 `cfg_*` 需靠 Excel 或更高层网表桥接。

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

### `make_mock_regmap.py` —— 解析控制信号→寄存器 + 本地复刻寄存器 Excel

用抓回的 REG_SHEET 行(`--rowdump` 结果) + 控制信号 list + 变体映射(alias)，把每个控制信号
解析到 寄存器/绝对地址(base+offset)/bit/默认值/关断值，并生成一个**结构一模一样**的
nManager 布局 `.xlsx`（本地开发用，不再依赖真文件）。脚本本身不含真实信号名，只读 private/ 输入。

```powershell
python make_mock_regmap.py --rows pll_rows.json --signals control_signals.json ^
  --aliases aliases.json --schema REG_SHEET.schema.json ^
  --out-xlsx REG_SHEET_mock.xlsx --out-map signal_reg_map.json
```

## 状态

需求对齐 + 收集文件基本完成。PLL/LO 控制信号已反查到 REG_SHEET 并解析出 地址/bit/默认值；
本地已复刻结构一致的寄存器 Excel，后续开发不依赖黄区真文件。
