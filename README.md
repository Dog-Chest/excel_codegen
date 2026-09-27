# excel_codegen

**Excel 模板参数填写 + Jinja2 代码生成器**：把「YAML 定义变量与模板」→「Excel 填参数」→「渲染出代码」三段式流程固化下来。
参数在 Excel 里填写最直观（工程师熟悉、可批量右拉、能直接分享），而模板与生成规则留在 YAML / Git 里可评审、可版本管理。

```
example.yaml ──init──► template.xlsx ──(人工填写 Global / Local)──► render ──► Output 表 / 生成的代码文件
     ▲                                                                             │
     └─────────────────── Jinja2 模板（支持 Prefix+Value+Suffix）◄──────────────────┘
```

## 特性

- **一个 YAML 描述一切**：全局变量、局部变量、一个或多个输出模板（内联 `code` 或外部 `template_file`）。
- **Excel 就是表单**：`Global Parameter` 一行一个全局变量，`Local Parameter` 从 E 列开始一个 Case 一列，右拉即可增加 Case。
- **Prefix + Value + Suffix**：`{{ port }}` 得到 `GPIOA_PORT`，`{{ port.value }}` 得到 `A`，`{{ port.prefix }}` / `{{ port.suffix }}` 单独取值；另有 `pvs` / `wrap` 过滤器。空格有意义（`suffix: " m"` 就是 `" m"`）。
- **两种输出引擎**：`engine: snapshot`（默认）把渲染结果写成**文本快照**；`engine: excel` 把模板编译成 **Excel 公式** —— 改了参数 Excel 打开即重算，**不用再跑脚本**（`examples/example_formula.yaml`）。
- **两种输出布局**：`horizontal`（一个 Case 一列）/ `vertical`（一个 Case 一行），直接写回 Excel 的 Output 表。
- **派生参数（参数引用参数）**：`derived: "rho * g"` 让一个参数引用同 Case 的 local 与 global，参数表里写成 Excel 公式、改输入自动重算，模板里照常 `{{ rho_g }}`。
- **一本工作簿放多套规则**：`case_filter: "kind == 'EXT'"` 让某个模板只作用于匹配的 Case（共用同一张 Global 表）。
- **可导出代码文件**：文件名支持 Jinja2（`uart_init_{{ case_name }}.c`），一次生成多份代码。
- **工作簿自带说明与指纹**：`init` 生成 `HOWTO` 表写清"下一步跑什么"；每次写回记录**时间 / 参数指纹 / 输出指纹**，`excel-codegen check` 可直接判定"表里的代码是否已过期"（可放进 CI）。
- **清晰的错误提示**：YAML 语法（含重复键）/ 配置校验 / Excel 表缺失 / 变量缺失 / 模板与 `case_filter` 语法错误 / 公式模式越界写法（带行内容），都给出人话提示（中文）。
- **自带测试**：`pytest` **107 项**用例覆盖生成、读取、写回、渲染、导出、公式编译与求值、派生参数、`check` 与 CLI 全流程。

## 目录结构

```
excel_codegen/
├── pyproject.toml              # 依赖与 excel-codegen 命令入口（含 PEP 735 的 dev 依赖组）
├── uv.lock                     # 跨平台锁文件：Windows / macOS / Linux 装到完全相同的版本
├── .python-version             # 3.12 —— uv 据此挑解释器，缺了就自己下载
├── README.md                   # 本文档
├── CHANGELOG.md                # 0.1.0 → 0.5.2 变更历史
├── setup.sh                    # Linux / macOS 一键建环境（有 uv 用 uv，自动避开 NTFS 盘）
├── .gitattributes              # 行尾统一（LF），避免 Windows / Linux 混用产生整文件 diff
├── .github/workflows/ci.yml    # 可选：三系统 × 两 Python 版本的自动回归（本地等价命令就是 uv run pytest）
├── .gitignore
├── docs/
│   └── template_guide.md       # 模板库扩展指南（变量 / Prefix-Suffix / Excel 结构 / case_filter / 公式模式 / FAQ）
├── excel_codegen/              # ← 工具本体（可 pip install）
│   ├── __init__.py             # 对外 Python API
│   ├── __main__.py             # python -m excel_codegen
│   ├── cli.py                  # init / render / validate / check（typer + rich）
│   ├── models.py               # Pydantic 配置模型 + CaseData / RenderResult
│   ├── excel_io.py             # 生成模板、读取填写内容、写回结果、HOWTO 表与指纹
│   ├── derived.py              # 派生参数：参数引用参数（Python 求值 + 翻译成 Excel 公式）
│   ├── formula.py              # 公式引擎：把"纯替换"模板编译成 Excel 公式
│   ├── formula_eval.py         # 公式求值器：把公式算一遍，验证"值"与 Python 渲染一致
│   ├── jinja_env.py            # Jinja2 环境与 pvs / wrap（模板与派生参数共用） 
│   ├── renderer.py             # 上下文组装、case_filter、渲染与导出
│   └── utils.py                # 错误类型、单元格地址运算、VarValue
├── examples/                   # ← 最小示例（开箱可跑）
│   ├── example.yaml            # 快照模式：2 个模板 / 2 种布局
│   ├── example_formula.yaml    # 公式模式：改参数免重跑
│   ├── template.xlsx           # 上面两个 YAML 生成的 Excel 示例
│   ├── template_formula.xlsx
│   ├── generated/              # 渲染输出示例（.c / .md）
│   └── generated_formula/      # 公式模式示例的导出文件
├── tests/                      # ← 107 项 pytest 用例
│   ├── conftest.py
│   ├── test_excel_io.py
│   ├── test_renderer.py
│   ├── test_formula.py
│   ├── test_formula_eval.py
│   ├── test_derived.py
│   └── test_cli.py
└── abs_fpi/                    # ← 现场实测项目（ABS FPI → GeniE 加载代码，与工具本体相互独立）
    ├── README.md               # 怎么用、边界、四层校验链
    ├── FINDINGS.md             # 对 excel_codegen 的实测报告 + 修复复测记录（本项目的"测试文档"）
    ├── TEMPLATES.md            # 模板库索引 + 加新规范的清单
    ├── abs_fpi_external.yaml   # 规则集真源（外压 / 内压）
    ├── abs_fpi_internal.yaml
    ├── compose.yaml + compose.py     # 规则集 → 项目工作簿 YAML（合并 + 冲突检查）
    ├── abs_fpi.yaml                  # 合并后的项目配置（两套规则共用一张 Global 表）
    ├── *.xlsx                        # 三本工作簿（产物，Output 表是公式）
    ├── build.py                      # 一条命令：compose + 生成 + 导出 + 两层验证
    ├── verify_excel_engine.py        # 独立的公式求值器（项目侧回归工具）
    ├── compare_with_rules.js         # 与 GeniE/Rules 的产物逐行比对（21 项）
    ├── fill_cases.py                 # 样例工况取值（只在首次 --init 用）
    ├── templates/*.j2                # 有语法高亮的模板文件
    ├── generated/                    # 导出的 GeniE 代码（7 段 .js + 汇总 .md）
    └── probes/                       # 最小复现与探针（含 run_probes.py 与 _probe_run.txt）
```

> **怎么读这个仓库**：`excel_codegen/` 是工具，`examples/` 是最小用法，
> `abs_fpi/` 是"拿去干真活"的现场项目 —— 它同时是这套工具的**实测证据**：
> `FINDINGS.md` 记录每一轮试用发现的问题与修复复测，`build.py` 的验证链
> （`validate` → `check`（含公式求值）→ `verify_excel_engine.py` → `compare_with_rules.js`）
> 每跑一次都会重新证明"生成的代码是对的"。
> `abs_fpi` 不参与工具打包（`pyproject.toml` 只收 `excel_codegen*`），可以单独删掉。

## 安装

需要 **Python 3.11+**（Windows / Linux / macOS 均可）。实测通过 **3.11 / 3.12 / 3.13 / 3.14**
四个版本，见下面的「Python 版本策略」。

```bash
cd excel_codegen

# 推荐：虚拟环境（放哪儿有讲究，见下面「Ubuntu」第 3 条）
python3 -m venv .venv
# Windows
.venv\Scripts\activate
# Linux / macOS
source .venv/bin/activate

# 安装（含测试依赖）
pip install -e ".[dev]"
```

安装后会得到命令 `excel-codegen`（也可用 `python -m excel_codegen`）。

### Python 版本策略（以后加功能时怎么判断）

| 角色 | 版本 | 定义在哪 |
| --- | --- | --- |
| **支持的底线** | **3.11** | `pyproject.toml` 的 `requires-python` |
| **CI 真正测的** | 3.11（底线）+ 3.13 | `.github/workflows/ci.yml` 的矩阵 |
| **开发默认** | **3.12** | `.python-version`（uv 据此挑解释器，本机没有会自动下载） |

三条规矩：

1. **`requires-python` 只写"CI 真的跑过"的最老版本。** 现在写 `>=3.11`，因为 **3.10 已在
   2026-10 结束支持**（[Python 版本支持周期](https://devguide.python.org/versions/)），
   继续兼容等于替一个没人维护的解释器背锅。
2. **想用更新的语法，就同时改三处**：`requires-python`、CI 矩阵、README 这一节；
   然后跑一次 `uv lock`（依赖解析会跟着变）与 `uv run pytest`。CI 里有一条
   `uv lock --check`，改了 `pyproject.toml` 忘了重新锁会直接失败。
3. **平时不用惦记版本**：为 3.11 写的代码在 3.12 / 3.13 / 3.14 上必然也能跑，
   反方向才会出事 —— 而拦住反方向的正是"CI 测底线"。

> 以后新增功能时，只有两类东西需要先想一下版本：
> ① 3.12+ 才有的标准库 API（`itertools.batched`、`typing.override`、`tomllib` 是 3.11 就有）；
> ② 3.12+ 的新语法（PEP 695 的 `type X = ...`、`def f[T]()`）。
> 想用就把底线抬到 3.12（一行 + CI 矩阵），不想用就继续待在 3.11 —— **没有任何东西强制你选**。

### Ubuntu / 从 Windows 拷过来的仓库：三个坑

本仓库原本在 Windows 上开发，整目录拷到 Ubuntu 后有三件事必须处理，否则"装不上"或"跑得极慢"：

1. **不要复用 `.venv/`**。虚拟环境里记着绝对路径（`pyvenv.cfg` 的 `home=`、`Scripts/` 目录），
   从 Windows 拷过来在 Linux 上完全不可用 —— 直接删掉重建（它本来就在 `.gitignore` 里）。
2. **`python3 -m venv` 报 `ensurepip is not available`**。Ubuntu 把虚拟环境拆成了独立包：

   ```bash
   sudo apt install python3-venv python3-pip git
   ```

3. **外置 NTFS / exFAT 盘上，环境放到原生文件系统，并且别中断安装**。
   实测结论（比"NTFS 不能用"更准确的说法）：
   * 干净地写入、干净地删除一个 36MB / 1080 个文件的 venv，在 ntfs3 上是**正常且很快**的；
   * 真正会把盘卡住的是**安装写到一半被强杀**：一次被打断的 `pip install --target` 留下了一个
     读不动的 `__pycache__`，之后连 `rm -rf` / `rmdir` 都停在 `D`（不可中断睡眠）状态，
     必须重启才能清掉。慢只是次要问题（碎文件多）。

   所以：环境放原生盘更省心，且**不要在装依赖时 Ctrl-C**。

   ```bash
   cd /path/to/excel_codegen
   python3 -m venv ~/.venvs/excel_codegen
   source ~/.venvs/excel_codegen/bin/activate
   pip install -e ".[dev]"
   pytest
   ```

   仓库自带的 `./setup.sh` 就是按这条规则做的（自动检测文件系统、选位置、装依赖、跑测试）：

   ```bash
   ./setup.sh                    # 环境默认建在 ~/.venvs/excel_codegen
   EXCEL_CODEGEN_VENV=/data/venvs/ecg ./setup.sh   # 也可以自己指定
   ```

### 想在三个系统上零配置跑起来：用 uv（推荐）

仓库里的 `pyproject.toml` + `uv.lock` + `.python-version` 已经配好了。装了 uv 之后，
**不需要 venv、不需要 pip、不需要 apt、也不用管本机是哪个 Python 版本**：

```bash
# 每台机器装一次 uv（三选一）
curl -LsSf https://astral.sh/uv/install.sh | sh                 # Linux / macOS
powershell -c "irm https://astral.sh/uv/install.ps1 | iex"      # Windows PowerShell
pipx install uv                                                 # 任何系统

# 之后在任何系统、任何目录，命令完全一样
uv run pytest                                   # 自动建环境 + 按 uv.lock 装依赖 + 跑测试
uv run excel-codegen --version                  # 跑 CLI，不用 activate
uv run excel-codegen render -c examples/example.yaml -x examples/template.xlsx --write-excel
uv run python abs_fpi/build.py --check
```

它替你做的三件事，正好就是"额外配置"的来源：

| 麻烦 | uv 怎么办 |
| --- | --- |
| 本机 Python 版本不对 / 没有 | `.python-version` 写着 `3.12`，**没有就自动下载**（不需要管理员权限） |
| 各机器装到的依赖版本不一致 | `uv.lock` 是**跨平台锁文件**，同一份锁同时含 `win_amd64` / `macosx_*` / `manylinux` / `musllinux` 的 wheel，三个系统装出来完全一致 |
| 测试依赖要额外加参数 | `[dependency-groups] dev` 是 uv 的**默认**依赖组，`uv run` 会一起装上 —— 所以 `uv run pytest` 不需要 `--extra dev` |

这个仓库只需要记住两条命令：**`uv run pytest`**（验证）和 **`uv run excel-codegen ...`**（干活）。

> **环境放哪**：uv 默认在仓库里建 `.venv/`（在原生盘上这是最优的）。仓库在外置 NTFS / exFAT 盘上、
> 想把环境挪走时用环境变量（这是 uv 唯一支持的方式，`pyproject.toml` 里没有对应开关）：
>
> ```bash
> export UV_PROJECT_ENVIRONMENT=~/.venvs/excel_codegen                                  # Linux / macOS
> $env:UV_PROJECT_ENVIRONMENT = "$HOME\.venvs\excel_codegen"                            # Windows PowerShell
> ```
>
> 或者直接跑 `./setup.sh`，它会自动探测文件系统并替你设好。
>
> **没装 uv 也能用**：`./setup.sh`（Linux/macOS）或 `pip install -e ".[dev]"`（任何系统 + venv）。
> 两条路都保留：`pyproject.toml` 里 `[project.optional-dependencies] dev` 与
> `[dependency-groups] dev` 内容一致，改一处记得改另一处。

> **`abs_fpi/compare_with_rules.js` 需要一个外部仓库**：它是"与 `GeniE/Rules` 逐行交叉比对"
> 那条独立证据，依赖**同级目录**下的 `../../GeniE/Rules`。该仓库不在本工作区时，脚本会打印
> `SKIP` 并以退出码 **2** 结束（`--allow-missing` 可视为跳过、退出码 0）—— **不会**把
> "没验证"伪装成"通过"。其余三层验证（`check` 值校验、`verify_excel_engine.py`、`build.py --check`）
> 不依赖外部仓库，可以照常跑。

## 快速开始

> 下面用的是 `excel-codegen`。用 uv 的话把它换成 `uv run excel-codegen` 即可
> （`uv run` 会先保证环境与 `uv.lock` 一致，所以不用先 activate）。

```bash
# 1) 生成 Excel 模板（--cases 指定初始 Case 列数，默认 2）
excel-codegen init --config examples/example.yaml --output examples/template.xlsx --cases 2

# 2) 在 Excel 里填写：Global Parameter 的 B 列、Local Parameter 的 E/F 列（需要更多 Case 就右拉）

# 3) 渲染：写回 Excel 的 Output 表
excel-codegen render --config examples/example.yaml --excel examples/template.xlsx --write-excel

#    或：导出为代码文件
excel-codegen render --config examples/example.yaml --excel examples/template.xlsx --outdir generated/

# 4) 只校验配置（可附带校验 Excel 结构），不产生输出
excel-codegen validate --config examples/example.yaml --excel examples/template.xlsx

# 5) 想知道"表里的代码是不是已经过期"（CI 可用，过期退出码 1）
excel-codegen check --config examples/example.yaml --excel examples/template.xlsx
```

`init` 输出：

```
                Excel 模板已生成
┌───────────┬──────────────────────────────────┐
│ 文件      │ examples\template.xlsx           │
│ Global 表 │ Global Parameter                 │
│ Local 表  │ Local Parameter（Case 列：E 起） │
│ Output 表 │ Output, Output Vertical          │
│ Case 列   │ Case1, Case2                     │
│ 说明表    │ HOWTO                            │
│ 模板数    │ 2                                │
└───────────┴──────────────────────────────────┘
```

`render` 输出（节选）——**"写回 Excel"这一行永远存在**，所以"成功"和"成功但没动文件"能一眼分开：

```
│ Case       │ Case1, Case2                             │
│ 渲染结果   │ 4 个（2 模板 × 2 Case）                  │
│ 写回 Excel │ 是                                       │
│ 导出文件   │ examples\generated\uart_init_Case1.c     │
│            │ examples\generated\uart_summary_Case1.md │
OK 渲染完成
```

不带 `--write-excel` / `--outdir` 时：

```
│ 写回 Excel │ 否（需要 --write-excel） │
│ 导出文件   │ 无                       │
! 本次只预览：没有写回 Excel，也没有导出文件。加 --write-excel / --outdir 才会落盘。
OK 渲染完成
```

`check` 输出（节选）：先比指纹，再逐行比内容，指出第几行不同：

```
│ 上次渲染时间 │ 2026-09-24 15:23:26                        │
│ 参数指纹     │ 记录 c1273da2acb5 / 当前 8f1c0e77a913   ← 参数改过了 │
│ 输出指纹     │ 记录 80ecfe5dc6e7 / 当前 91ac2f0b7d44   │
ERROR 输出表已过期，共 1 处不一致：
    Output 第 B 列（Case1）：第 2 行不同：表里 'UART_Init(115200, …)'，应为 'UART_Init(9600, …)'
    → 跑一次 `excel-codegen render -c … -x … --write-excel` 刷新
```

仓库中已附带示例产物：`examples/template.xlsx`、`examples/generated/uart_init_Case1.c`、`examples/generated/uart_summary_Case1.md`。

生成的文件示例（`examples/generated/uart_init_Case1.c`）：

```c
// Case: Case1
UART_Init(115200, GPIOA_PORT, MODE_TX_RX);
// 纯值: A, 前缀: GPIO, 后缀: _PORT
// 过滤器: GPIOA_PORT
// MCU: STM32F103
```

## 命令说明

全局：`excel-codegen --version`

### `excel-codegen init`

| 选项 | 说明 | 默认 |
| --- | --- | --- |
| `-c, --config PATH` | YAML 配置文件（必填，必须存在） | — |
| `-o, --output PATH` | 生成的 Excel 路径 | 配置中的 `excel.output` |
| `--cases VALUE` | 初始 Case 列：数量（`3`）**或逗号分隔的名字**（`EXT-T20.559,INT-T15`） | `2` |
| `-f, --force` | 目标文件已存在时覆盖 | 否 |
| `--template-sheet / --no-template-sheet` | 是否生成隐藏的 `Template` 表（保存模板原文与元信息） | 生成 |
| `--howto / --no-howto` | 是否生成 `HOWTO` 说明表（放在第一张） | 生成 |

### `excel-codegen render`

| 选项 | 说明 | 默认 |
| --- | --- | --- |
| `-c, --config PATH` | YAML 配置文件（必填） | — |
| `-x, --excel PATH` | 已填写的 Excel | 配置中的 `excel.output` |
| `-w, --write-excel` | 渲染结果写回 Excel 的 Output 表（并刷新 HOWTO / 指纹） | 否 |
| `-d, --outdir PATH` | 渲染结果导出为代码文件到该目录 | 否 |
| `--case NAME` | 只渲染指定 Case（可重复传入） | 全部 Case |
| `--show / --no-show` | 在终端打印渲染结果 | `--show` |
| `--overwrite / --no-overwrite` | 导出文件已存在时是否覆盖 | `--overwrite` |

不传 `--write-excel` 和 `--outdir` 时只在终端预览渲染结果（安全模式）；摘要里会明确写出"否（需要 --write-excel）"。

### `excel-codegen validate`

| 选项 | 说明 |
| --- | --- |
| `-c, --config PATH` | YAML 配置文件（必填） |
| `-x, --excel PATH` | 顺带校验 Excel 结构：表是否存在、表头是否正确、每个 Case 的取值、上次渲染时间与指纹 |

校验内容：YAML 结构与语义（含重复键）、模板与 `case_filter` 的 Jinja2 语法、模板引用的变量是否已定义、**定义了却没有被引用的变量**、`output_sheet` 是否已声明、变量是否重名、Case 取值一览。

### `excel-codegen check`

| 选项 | 说明 |
| --- | --- |
| `-c, --config PATH` | YAML 配置文件（必填） |
| `-x, --excel PATH` | 要检查的 Excel（默认取配置中的 `excel.output`） |
| `--values / --no-values` | 公式模式：把 Output 表的公式**在 Python 里算一遍**再与 Python 渲染比对（默认开） |

* 快照模式：拿当前参数重新渲染，与表内内容逐行比对（含 Case 表头）。
* 公式模式：① 比公式文本（与当前 YAML 是否一致）；② 把公式求值，与 Python 渲染比对
  —— 后者能抓到"列标指错 / 空单元格处理错"这类**公式本身**的问题，以及"插删过 Case 列"这种结构变化。
* **一致 → 退出码 0；过期 → 退出码 1** 并列出第一处差异（公式模式会指出是**模板第几行**）。
  适合放进 CI 断言"提交的工作簿与代码是同步的"。

## Excel 模板结构

生成的 `template.xlsx` 包含以下工作表（名称全部可在 YAML 中配置）：

| 工作表 | 作用 | 结构 |
| --- | --- | --- |
| `HOWTO` | **打开工作簿第一眼看到的东西**：三步怎么用、要跑哪条命令、本次生成时间与指纹、每个模板的输出位置 | A 列文本 |
| `Global Parameter` | 全局变量，对所有 Case 生效 | A=Variable，**B=Value（填写）**，C=Description，D=Prefix，E=Suffix |
| `Local Parameter` | 局部变量，每个 Case 一列 | A=Variable，B=Description，C=Prefix，D=Suffix，**E 起为 Case1、Case2 …** |
| `Output` / `Output Vertical` | 渲染结果写入区 | 由模板的 `output_sheet`、`start_cell`、`direction` 决定 |
| `Template`（隐藏） | 模板原文 + 机器可读元信息（`## excel-codegen-meta`） | A 列键/行号，B 列内容 |

填写规则：

1. **Global 表**：B 列留空 → 自动回落到 YAML 的 `default`；D/E 列留空 → 回落到 YAML 的 `prefix` / `suffix`（**非空则原样使用，首尾空格保留**）。
2. **Local 表**：E 列是 `Case1`、F 列是 `Case2`…… 选中 E:H 之类的整块区域**右拉**即可增加 Case；新增列的表头按 `Case3`、`Case4` 命名（也可直接改成 `EXT-T20.559` 这样的工况名）。
3. Case 列必须是**连续**的：从 E 列开始遇到空表头即认为 Case 结束。
4. 单个 Case 内的单元格留空 → 回落该变量的 YAML `default`。
5. `HOWTO` 与 `Template` 表由工具维护：改它们不影响渲染结果（`Template` 表首行也写了这句提示）。
6. 想让工作簿自己说明"上次是什么时候、用哪套参数生成的"：看 `HOWTO` 表的"本次生成"，或跑 `excel-codegen check`。

## 变量与 Prefix / Value / Suffix

YAML 中每个变量可以带 `prefix` / `suffix`，渲染时包装成 `VarValue`：

```python
VarValue(value="A", prefix="GPIO", suffix="_PORT")
str(v)        # "GPIOA_PORT"   <- 模板里写 {{ port }}
v.value       # "A"            <- {{ port.value }}
v.prefix      # "GPIO"         <- {{ port.prefix }}
v.suffix      # "_PORT"        <- {{ port.suffix }}
```

模板与结果对照：

```jinja
UART_Init({{ baud }}, {{ port }}, {{ mode }});
// 纯值: {{ port.value }}, 前缀: {{ port.prefix }}, 后缀: {{ port.suffix }}
// 过滤器: {{ port.value | pvs("GPIO", "_PORT") }}
// 等价的 wrap: {{ port.value | wrap("GPIO", "_PORT") }}
```

```c
UART_Init(115200, GPIOA_PORT, MODE_TX_RX);
// 纯值: A, 前缀: GPIO, 后缀: _PORT
// 过滤器: GPIOA_PORT
// 等价的 wrap: GPIOA_PORT
```

说明：

- 变量在**渲染上下文中始终是 `VarValue`**，所以 `{{ x }}` 是组合值，纯值请显式写 `{{ x.value }}`。
- 过滤器 `pvs(value, prefix, suffix)` / `wrap(value, prefix, suffix)`（两者等价）用于对**任意表达式**做一次组合，例如 `{{ 115200 | pvs("BAUD_", "U") }}` → `BAUD_115200U`。
- **数值形态一致**：整数浮点（`340.0`）在输出时规范化为 `340`，这条规则同时作用于 `{{ x }}`、`{{ x.value }}` 和 `{{ 340.0 }}`（Jinja2 `finalize`），所以两条路径不会产生假差异；`bool` 输出 `true`/`false`，`None` 输出空串。`|float`、`|int` 等过滤器不受影响。
- **空格有意义**：`prefix` / `suffix` 只在"Excel 单元格真的为空"时才回落到 YAML，非空值原样使用 —— `suffix: " m"` 得到 `340 m` 而不是 `340m`。
- 模板默认使用 `StrictUndefined`：引用了不存在的变量会**立即报错**并指出变量名，而不是静默渲染为空。

## 一本工作簿放多套规则（`case_filter`）

`render` 默认是"模板 × Case"的交叉积。给模板加一行 `case_filter` 就能只渲染匹配的 Case，
于是"共用一张 Global 表 + 内外压两套模板"可以放进同一本工作簿，主尺度只填一次：

```yaml
templates:
  - name: ext_code
    output_sheet: "Code EXT"
    case_filter: "kind == 'EXT'"      # 只对本列 kind=EXT 的 Case 生效
    code: |
      // EXT {{ case_name }} (L={{ L }})
  - name: int_code
    output_sheet: "Code INT"
    case_filter: "kind == 'INT'"
    code: |
      // INT {{ case_name }} (L={{ L }})
```

* 过滤器里变量**当字符串用时等于组合值**（和 `{{ x }}` 一致），所以要拿纯值或做数值比较请用 `.value`：
  `case_filter: "draft.value > 20"`。
* 被跳过的 Case 不占输出列/行，`render` 的摘要里会写明"跳过 N 个"。
* 语法错误或引用了取不到的变量会直接报 `RenderError`，不会静默变成"什么都没生成"。

## 公式模式：改参数不用再跑脚本（`engine: excel`）

默认（`snapshot`）输出表是**文本快照**：改了参数必须重跑 `render --write-excel`。
把模板改成 `engine: excel`，输出表里写的就是**公式** —— 改参数后 Excel / WPS 打开即重算。

```yaml
templates:
  - name: uart_init
    output_sheet: "Output"
    start_cell: "B2"
    direction: "horizontal"
    engine: "excel"                   # ← 只要这一行
    code: |
      UART_Init({{ baud }}, {{ port }}, {{ mode }});
      // 纯值: {{ port.value }}, 前缀: {{ port.prefix }}, 后缀: {{ port.suffix }}
```

```bash
# 生成 + 写一次公式；之后改参数就完全不用碰脚本了
excel-codegen init   --config examples/example_formula.yaml --output examples/template_formula.xlsx
excel-codegen render --config examples/example_formula.yaml --excel examples/template_formula.xlsx --write-excel
```

**支持范围（只做"纯替换"）**：`{{ x }}` / `{{ x.value }}` / `{{ x.text }}` / `{{ x.prefix }}` /
`{{ x.suffix }}` / `{{ case_name }}` / `{{ template_name }}`。
出现 **Jinja** 的 `{% if %}`、`{% for %}`、过滤器、算式会**直接报错并给出行内容**，让该模板改回 `snapshot`。

> 判断标准是"有没有 `{%`"，**不是**"代码里有没有 `if`"：生成目标语言自己的 `if/else`
> （GeniE / C / Python 语句）只是字面文本，不影响公式模式。`abs_fpi` 的内外压模板就是这样
> —— 那些 C_1 / k_lo / Girth 插值的 `if/else` 是字面文本，只有开头几行坐标映射是 Jinja，
> 改写成纯替换之后整份模板就进公式模式了。

**三条代价（务必知悉）**：

1. **值只活在 Excel 里** —— `openpyxl` 不算公式，所以 `--outdir` 导出文件仍由 Python 渲染。
   但 `check` 默认会**把公式在 Python 里算一遍**（内置求值器 `formula_eval.py`）再与 Python 渲染
   逐行比对，所以"Excel 里看到的"与"导出的"是否同一段代码是**有验证的**（`--no-values` 可关）。
2. **改参数免脚本，改模板 / 改结构仍需重跑**（`--write-excel` 刷新公式）—— YAML 仍是唯一真源。
   插删**变量行**不用重跑（按变量名定位）；**增删 Case 列要重跑**（列标是相对字母）。见指南 §14.5。
3. 两个已知差异：`type: bool` 在 Excel 里是 `TRUE`/`FALSE`；`TEXT()` 格式串受区域设置影响。
   公式长度：一个 `{{ x }}` ≈ 300–400 字符，一行 8 个以上占位符就该拆行（会告警），硬上限 8192 字符。

同一个工作簿里可以两种引擎混用（`engine` 是 per-template 的）。
细节与取舍：见 [`docs/template_guide.md` §14](docs/template_guide.md)。

## 输出布局

`horizontal`（默认）——一个 Case 一列，结果行向下展开；Case 名写在起始单元格的上一行：

```
        B        C          ← start_cell = "B2"
1     Case1    Case2
2   // Case: Case1   // Case: Case2
3   UART_Init(...)   UART_Init(...)
```

`vertical` —— 一个 Case 一行，结果行向右展开；Case 名写在起始单元格左侧一列：

```
        A        B                  C
2     Case1   | case_name | ... |  | --- | ... |
3     Case2   | case_name | ... |  | --- | ... |
```

重新渲染时会按**表内原有的真实边界**清理旧区域，所以"改短模板 / 减少 Case"之后也不会残留上一次的内容。

## 作为 Python 库使用

```python
from excel_codegen import (
    load_config, create_template, render_all, write_results, export_files,
    build_environment, input_fingerprint, output_fingerprint, read_metadata,
)

config = load_config("examples/example.yaml")          # 读取并校验 YAML
create_template(config, "template.xlsx", cases=3, overwrite=True)

output = render_all(config, "template.xlsx")           # 读取填写的参数并渲染
write_results("template.xlsx", config, output.results) # 写回 Output 表 + 指纹
files = export_files(config, output.results, "generated/")   # 导出代码文件

print(output.results["uart_init"][0].text)             # 取某个模板某个 Case 的文本
print(output.skipped)                                  # case_filter 跳过了哪些 Case
print(input_fingerprint(output.global_values, output.cases))
```

## 错误处理

所有面向用户的错误都继承自 `CodeGenError`，CLI 会以 `ERROR 提示` 形式打印并返回退出码 1：

| 异常 | 触发场景 | 提示示例 |
| --- | --- | --- |
| `ConfigError` | YAML 语法错误、同一映射里键重复、字段非法、重名、`output_sheet` 未声明 | `YAML 键重复: 'variables'（第 20 行）…` |
| `ExcelError` | 文件不存在、工作表缺失、表头不对、无 Case 列、变量名重复 | `Excel 缺少工作表: 'Local Parameter'。当前工作表: ...` |
| `RenderError` | 模板 / `case_filter` 语法错误、变量缺失、导出文件名非法 | `模板 'uart_init' 语法错误：第 3 行: Unexpected end of template ...` |

> 控制台兼容：CLI 只用 ASCII 标记（`OK` / `ERROR` / `!`），并在启动时把 stdout/stderr 的
> 错误处理设为 `backslashreplace`。中文 Windows（GBK 控制台）下**不会**再出现
> "活干完了、最后一行字打不出来、退出码 1" 的情况。

## 测试

```bash
uv run pytest                  # 107 passed（推荐：零配置，见「安装」里的 uv 一节）
pytest                         # 已 activate 环境时等价
pytest --cov=excel_codegen
```

新机器上先跑 `./setup.sh` 会更省事：它建环境、装依赖，然后就把上面这条测试跑一遍。

## 0.5.2 变更：多平台零配置（uv）

| 变化 | 说明 |
| --- | --- |
| 新增 `uv.lock` | **跨平台锁文件**：同一份锁同时含 `win_amd64` / `win32` / `macosx_*` / `manylinux*` / `musllinux*` 的 wheel，三个系统装到的版本完全一致 |
| 新增 `.python-version` | `3.12`；本机没有时 uv 自动下载解释器，不需要管理员权限 |
| `[dependency-groups] dev` | PEP 735 依赖组，uv 默认装 —— 于是 `uv run pytest` 不需要任何额外参数 |
| `setup.sh` | 有 uv 走 `uv sync`，没有 uv 走原来的 venv + pip；两条路都实测过 |

现在跨平台只有两条命令：`uv run pytest` 与 `uv run excel-codegen ...`。详见「安装」。

## 0.5.1 变更：Windows → Ubuntu 迁移（只动环境与文档）

仓库从 Windows 整目录拷到 Ubuntu 之后，**代码不用改**（没有盘符、没有 Windows API、路径全是相对的），
要重建的是环境。本轮把踩到的坑固化进仓库：

| 变化 | 说明 |
| --- | --- |
| 新增 `./setup.sh` | 建 venv + 装依赖 + 跑测试；**先探测仓库所在文件系统**，NTFS / exFAT 上自动把 venv 放到 `~/.venvs/excel_codegen`（`EXCEL_CODEGEN_VENV` 可覆盖） |
| 新增 `.gitattributes` | `* text=auto eol=lf` —— Windows 时期生成的文件是 CRLF，不声明的话与 Linux 侧会整文件 diff |
| README「安装」 | 补上 Ubuntu 的三个坑：不复用旧的 `.venv/`、要装 `python3-venv`、环境别放外置 NTFS 盘（当时判断为"NTFS 写碎文件会卡死"；**0.5.2 已更正**：真正原因是安装被中途强杀，见上一节） |
| `compare_with_rules.js` | 缺外部 `GeniE/Rules` 仓库时打印 `SKIP` + 原因并以退出码 **2** 结束（`--allow-missing` → 0）：不崩溃，也不把"没验证"当成"通过" |
| 文档对齐 | 测试数量统一为 **107 项**；`tests/` 目录树补上漏写的 `test_derived.py`；更正 `build.py` / `verify_excel_engine.py` 里"`check` 在公式模式下比不了值"的过时说法（0.4.0 起已默认值校验） |

**工具行为零改动**：`excel_codegen/` 的生成、读取、写回、公式引擎、派生参数逻辑一行未动，
107 项测试与 `abs_fpi/build.py --check`（公式求值 28 项）在 Ubuntu 上复跑全过。

## 0.5.0 变更：派生参数（参数引用参数）

| 变化 | 说明 |
| --- | --- |
| **`derived: "表达式"`** | 变量可以不填值，而是引用同 Case 的 local 与 global 算出来；不跨 Case |
| **表里是活公式** | 派生格写成 Excel 公式（如 `=MAX((draft - z),0)`），改输入自动重算；描述里标「自动计算」、格子淡绿 |
| **一致性检查** | `validate` 提前编译（语法 / 越界 / 循环 / 能否翻译）；`check` 会递归求值派生格并与 Python 渲染比对；手工改派生格会告警并报错 |
| **数值形态统一** | 去掉 `TEXT()`（它会把 `20.559000000000001` 这种二进制尾巴打出来），两边都按 **15 位有效数字**（Excel General / Python `%.15g`） |
| **不翻译的写法** | `round` / `ceil` / `floor` / `//` 等两边语义不同的运算**故意不翻译**，降级为"写入算好的值"并告警 —— 宁可不自动，也不要"Excel 里看到的"和导出的不一致 |
| **示例** | `examples/example.yaml`（快照）与 `example_formula.yaml`（公式）都加了派生参数 |

## 0.4.0 变更：把"公式算出来的值"也验了

| 变化 | 说明 |
| --- | --- |
| **公式求值器（内置）** | 新增 `formula_eval.py`：解析 Output 表里**真实的公式文本**并求值（`&` / `IF` / `ISBLANK` / `TEXT` / `INDEX`+`MATCH` / 相对列），与 Python 渲染逐行比对 |
| **`check` 默认开值校验** | 公式模式下 `check` = 比公式文本 **+** 比公式算出的值（`--no-values` 可关）；参数单元格本身是公式时自动降级为"只比公式"并提示 |
| **差异信息给人看** | 公式模式不一致时贴**模板第几行**，不再贴几百字符的公式 |
| **空 Case 列告警** | 整列都空的 Case（新插的空列）会提示"它可能悄悄落进某个 `case_filter` 规则集" |
| **超长公式告警** | 一行公式超过 3000 字符时提醒拆行（一个 `{{ x }}` ≈ 300–400 字符，一行 8 个以上就该留意） |
| **文档更正** | 0.3.0 说"有 `if/else` 就不能用公式模式"是**错的**：要按"有没有 `{%`"判断，生成语言自己的 `if/else` 只是字面文本。新增 §14.5「改结构什么要重跑」 |

## 0.3.0 变更：公式模式（改参数免跑脚本）

| 变化 | 说明 |
| --- | --- |
| **`engine: excel`** | 新增公式引擎：Output 表写 Excel 公式而不是文本，改参数后 Excel / WPS 打开即重算 |
| **只做"纯替换"** | 支持 `{{ x }}` / `.value` / `.text` / `.prefix` / `.suffix` / `{{ case_name }}` / `{{ template_name }}`；控制流与过滤器报错并给出行内容 |
| **与 Python 侧对齐** | `INDEX/MATCH` 按变量名定位（插行不指错）、`ISBLANK` 处理空单元格回落、`TEXT()` 规范化数值、横向相对列 / 纵向绝对列 |
| **`check` 语义分模式** | 公式模式比"公式是否与当前 YAML 一致"（改参数不算过期）；快照模式比内容 |
| **`validate` 显示引擎** | 模板清单新增"引擎"列，公式模式在校验阶段就编译一遍（越界写法提前报错） |
| **`HOWTO` 表自适应** | 公式模板标"公式·自动重算"，快照模板标"快照·需重跑" |
| **新示例** | `examples/example_formula.yaml` → `examples/template_formula.xlsx` |

## 0.2.0 变更（实测报告修复会话）

0.2.0 来自一次真实移植的实测报告（`abs_fpi/FINDINGS.md`：用本工具重写 ABS FPI 内外压
GeniE 加载代码，7 个工况 2 个规范集）。报告列出的 12 条问题中 11 条已修，1 条判为能力边界：

| 变化 | 说明 |
| --- | --- |
| **Prefix/Suffix 不再被 strip** | `suffix: " m"` 现在得到 `340 m`（原来静默变成 `340m`） |
| **`render` 不再静默** | 摘要永远打印"写回 Excel │ 是 / 否（需要 --write-excel）"；只预览时补一条 `!` 提示 |
| **GC 友好** | `OK` / `ERROR` / `!` 全是 ASCII，stdout/stderr 加 `errors="backslashreplace"`；GBK 控制台退出码不再撒谎 |
| **清理旧结果按真实底边** | 改短模板 / 减少 Case 后重渲染不再残留上一版代码 |
| **工作簿自带 HOWTO 表** | 三步说明 + 命令 + 快照提醒 + 输出位置，放在第一张（`--no-howto` 可关） |
| **指纹与 `check`** | 每次写回记录时间 / 参数指纹 / 输出指纹；`excel-codegen check` 判定是否过期（CI 可用） |
| **`case_filter`** | per-template 的 Case 过滤，一本工作簿放多套规则、共用一张 Global 表 |
| **`{{ x.value }}` 形态一致** | 整数浮点规范化对 `{{ x }}` 与 `{{ x.value }}` 同时生效 |
| **YAML 重复键报错** | 配置阶段就给出"第 N 行键重复"，不再静默丢一份 |
| **`--cases` 支持名字** | `--cases EXT-T20.559,INT-T15` 直接建出有意义的工况名 |
| **`Template` 表首行提示** | 明说"只读参考：模板真源是 YAML / template_file" |
| **文档补齐** | 能力边界（没有第三层作用域）、`safe_filename` 规则、"空行各占一行"的噪音警告 |

## 未做 / 路线图

| 项 | 为什么现在不做 | 现在的替代 |
| --- | --- | --- |
| **`extends` / `!include`**（跨文件复用模板库） | 会改到"一个 YAML = 一本工作簿"的核心约定、`source_dir` 解析与 `template_file` 相对路径，属产品级改动；合并语义还要定"字段冲突报错还是告警" | 项目侧 `abs_fpi/compose.py` + `compose.yaml`：把多个规则集 YAML 合并成项目 YAML，并检查同名变量的 `prefix`/`suffix`/`type` 是否一致 |
| **第三层作用域**（船 → 工况 → 舱/设备） | 变量解析、Excel 表结构、`render` 循环语义都要动 | 按工况摊平（见指南 §13）；`case_filter` 只解决"模板 × Case"的交叉积 |
| **公式模式支持 Jinja 控制流** | 单元格引用表达不了"行数不定 / 条件包含" | 把模板改写成纯替换（分支搬进参数），或该模板留在 `snapshot` |
| **在 Excel 之外求值公式** | `openpyxl` 不算公式，这是上游限制 | `check` 内置求值器（`formula_eval.py`）已在 Python 侧算一遍做验证；导出文件仍由 Python 渲染 |

## 文档

- 扩展模板库、变量定义、Prefix/Suffix、`case_filter`、Jinja2 语法、Excel 表结构、`check`、常见问题与能力边界：见 [`docs/template_guide.md`](docs/template_guide.md)。
- 版本变更历史：[`CHANGELOG.md`](CHANGELOG.md)。
- 现场实测项目（ABS FPI 内外压 → GeniE 加载代码，含公式模式与四层校验链）：[`abs_fpi/README.md`](abs_fpi/README.md) 与 [`abs_fpi/FINDINGS.md`](abs_fpi/FINDINGS.md)（对工具本身的实测报告 + 修复复测记录）。

## License

MIT
