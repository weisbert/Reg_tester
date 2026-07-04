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

默认目标模块：`CHAIN_TOP_A`、`CHAIN_TOP_B`、`CLK_MUX`。
输出把端口分成 控制输入 / 输出 / 模拟 inout / 电源地 四类，控制输入即候选控制信号。

## 状态

需求对齐 + 收集文件中。
