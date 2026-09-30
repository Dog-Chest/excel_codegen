# excel_codegen

**在 Excel 里填参数，用 YAML + Jinja2 出代码。**

[![ci](https://github.com/Dog-Chest/excel_codegen/actions/workflows/ci.yml/badge.svg)](https://github.com/Dog-Chest/excel_codegen/actions/workflows/ci.yml)
![python](https://img.shields.io/badge/python-3.11%2B-blue)
![license](https://img.shields.io/badge/license-MIT-green)

> 🤖 **本项目由 DeepSeek 自动生成**：代码、测试与文档都在 DeepSeek Harness 里由 AI 编写，
> 人工负责提需求、定方向与验收。

工程师最熟悉的填参数界面是 Excel，最该被评审和版本管理的生成规则却常常散在一次性脚本里。
这个工具把两头接起来：**变量与模板写在 YAML**（进 Git、可 diff、可评审、可攒成模板库），
**生成一本 Excel 之后就脱离命令行干活** —— 参数在 Excel 里填，输出表里的代码由 Excel 自己算，
改一格就变，不需要再跑任何命令。

```
   YAML 模板库                     一本独立可用的工作簿
┌──────────────────┐        ┌───────────────────────────────────────┐
│ variables        │        │ Global / Local Parameter：填参数       │
│ templates(Jinja2)│──init─►│ Output：**Excel 公式**，改参数自己重算  │
└──────────────────┘        └───────────────────────────────────────┘
         │                                  │
         │ 下次复用 / 微调再改 YAML          │ 要把代码导成文件时才用命令行
         └──────────────────────────────────┴──► render --outdir
```

## 快速开始

```bash
# 1) 把内置示例拷出来 —— 示例随包发布，装了 pip 包就够，不用克隆仓库
excel-codegen examples                      # 看有哪些、各自演示什么
excel-codegen examples --copy ./examples

# 2) 打开示例里那本已经填好样例参数的工作簿就能改：
#    Global Parameter 的 B 列、Local Parameter 的 E/F 列
#    Output 表里是活公式 —— 改完参数**它自己就重算了**，不需要任何命令

# 3) 只有「要把代码导成文件」时才回到命令行
excel-codegen render -c examples/basic/example_formula.yaml \
    -x examples/basic/template_formula.xlsx --outdir out
```

想核对「Excel 里算出来的」与「Python 渲染的」是否一致（CI 里很有用）：`excel-codegen check`。

没装 uv 就去掉 `uv run`，先按下面「安装」把环境准备好。

内置示例（`excel-codegen examples`，源码在 [`excel_codegen/examples/`](excel_codegen/examples/)）：

| 示例 | 演示什么 |
| --- | --- |
| [`basic`](excel_codegen/examples/basic/) | 入门：`example_formula.yaml` 是**默认的公式模式**（改参数后打开 Excel 就重算）；`example.yaml` 显式 `engine: snapshot`，演示过滤器 / 循环 / 导出文件 |
| [`nastran`](excel_codegen/examples/nastran/) | **一行一个工况**（`local_direction: vertical`）+ 公式模式生成 NASTRAN 工况控制语句：语句留空就不输出，在 Excel 里改一格 Code 列立刻跟着变（指南 §19 / §20） |
| [`abs_fpi`](excel_codegen/examples/abs_fpi/) | **现场用例**：ABS FPI 内外压 → GeniE。两个规则集可各自单用，也可合成一本项目工作簿共用一张 Global 表；内压用成员表 `Tank Data` 把舱参数只写一遍 |

## 特性

| | |
| --- | --- |
| **一个 YAML 描述一切** | 全局变量、局部变量、多个输出模板（内联 `code` 或外部 `template_file`） |
| **Excel 就是表单** | 一行一个全局变量；局部参数默认一个工况一列（右拉即增列），也可切成**一行一个工况**（`local_direction: vertical`，下拉即增行）—— 输入与输出布局各自独立 |
| **Prefix + Value + Suffix** | `{{ port }}` → `GPIOA_PORT`，`{{ port.value }}` → `A`；空格有意义（`suffix: " m"` 就是 `" m"`） |
| **默认就是公式模式** | 输出表里写 **Excel 公式**：改参数后打开 Excel 即重算，**不用再跑脚本**，工作簿脱离命令行也独立可用。支持**行内 `{% if %}`**（编译成 `IF()`）；模板要用 `{% for %}` / 过滤器 / `{% include %}` 时才显式写 `engine: snapshot` |
| **派生参数** | `derived: "rho * g"` 让参数引用参数（同工况的 local + global），参数表里是活公式 |
| **取值约束** | `min` / `max` / `choices` / `pattern` 声明合法取值：Excel 里变下拉列表与数值范围，`render` / `validate` / `check` 每次读表都再查一遍 —— 挡住"手滑把 20.559 打成 205.59 却照样生成代码" |
| **一本工作簿放多套规则** | `case_filter: "kind == 'EXT'"` 让模板只作用于匹配的工况，共用一张 Global 表 |
| **跨文件复用** | `extends: [rules/a.yaml, rules/b.yaml]` 把几套规范合并进一份项目配置 —— 变量与模板不用手抄，`template_file` 相对各自文件解析 |
| **三层作用域** | 船（global）→ 工况（local）→ **舱 / 设备（成员表）**：被多个工况引用的舱参数只写一遍，改一处就够 |
| **可导出代码文件** | 文件名支持 Jinja2，**参数也能用**（`cc_{{ seq }}_{{ case_name }}.inc`），一次生成多份 |
| **建表时就把错误拦住** | 模板语法、`{% include %}` 片段、以及公式模式表达不了的写法（过滤器 / 循环 / 跨行 `{% if %}`）在 `init` 就报，并告诉你怎么改 |
| **自带说明与指纹** | 工作簿里有 `HOWTO` 表与生成指纹；`excel-codegen check` 判定"表里的代码是否已过期"，**CI 可用** |

## 现场用例：ABS FPI 内外压 → GeniE

[`excel_codegen/examples/abs_fpi/`](excel_codegen/examples/abs_fpi/) 是拿这个工具干真活的现场用例
（`excel-codegen examples --copy .` 就能拿到）—— 把 **ABS FPI**（5A-3-2/5.5 外压、
5.7 内压）的面载荷计算，生成可直接粘进 GeniE 的 JavaScript 函数体。

| 工作簿 | 内容 | 工况 |
| --- | --- | --- |
| `abs_fpi_external.xlsx` | 外压单用 | 3 |
| `abs_fpi_internal.xlsx` | 内压单用 | 4 |
| `ABS_FPI_load_cases.xlsx` | **项目工作簿**：两套规则共用一张 Global 表，主尺度只填一次 | 7 |

三本都是 `engine: excel` —— 在 Excel 里改吃水 / 舱容，代码列立刻重算。
模板里那些 C_1 区间、Girth 插值的 `if/else` 是**生成目标语言自己的字面文本**，
不是 Jinja 控制流，所以整份模板能进公式模式。

它同时是这个工具的**实测证据**：四层校验链（`validate` → `check`（含公式求值）→
`verify_excel_engine.py` → `compare_with_rules.js`）每跑一次都重新证明"生成的代码是对的"；
逐轮试用发现的问题与修复复测记录在 [`abs_fpi/FINDINGS.md`](abs_fpi/FINDINGS.md)
（`abs_fpi/` 是开发侧的现场脚本与实测报告，**不随包发布**；随包发布的是上面那个示例目录）。

> 用例里的舱容与工况组合只是**示意数值**（随手取的几个数），不代表任何真实船舶。

## 文档

| 文档 | 内容 |
| --- | --- |
| [`docs/template_guide.md`](docs/template_guide.md) | **模板库扩展指南**：YAML 字段、Prefix/Suffix、Jinja2 速查、Excel 表结构、`case_filter`、排错、FAQ、能力边界、公式模式与派生参数 |
| [`docs/cli.md`](docs/cli.md) | 命令参考（六个命令的全部选项）、作为 Python 库使用、错误类型 |
| [`docs/setup.md`](docs/setup.md) | 安装与跨平台环境：uv 零配置、venv + pip、Windows / Ubuntu 的坑、Python 版本策略 |
| [`excel_codegen/examples/README.md`](excel_codegen/examples/README.md) | **内置示例**：三个示例各自演示什么、怎么拷出来、怎么改 |
| [`CHANGELOG.md`](CHANGELOG.md) | 版本变更历史 |
| [`abs_fpi/README.md`](abs_fpi/README.md) | 现场项目：怎么用、边界、四层校验链 |
| [`docs/publishing.md`](docs/publishing.md) | 维护者向：发布到 GitHub / PyPI、密钥与隐私 |

## 安装

需要 **Python 3.11+**（Windows / Linux / macOS 均可）。

```bash
# 推荐：uv —— 不用建 venv、不用 pip、不用管本机 Python 版本
curl -LsSf https://astral.sh/uv/install.sh | sh     # Windows 见 docs/setup.md

# 或经典方式
python3 -m venv ~/.venvs/excel_codegen && source ~/.venvs/excel_codegen/bin/activate
pip install -e ".[dev]"
```

装好后得到命令 `excel-codegen`（也可用 `python -m excel_codegen`）。
跨平台细节、外置盘注意事项与 Python 版本策略见 [`docs/setup.md`](docs/setup.md)。

> **PyPI 发布流程已就绪、尚未首次发布**（[`docs/publishing.md`](docs/publishing.md) 里有一次性登记步骤）。
> 发布之后就可以不克隆仓库、直接 `uv tool install excel-codegen`（或 `pipx install excel-codegen`）。

## 开发

```bash
uv run pytest --cov          # 341 项用例 + 覆盖率门槛（85%）
uv run ruff check .          # lint
uv run ruff format --check . # 格式
uv run mypy                  # 类型检查
```

CI（[配置](.github/workflows/ci.yml)）在 **Windows / macOS / Linux × Python 3.11 / 3.13**
上把上面四条各跑一遍。支持下限是 **3.11**，开发默认 3.12；改下限的规矩见 [`docs/setup.md`](docs/setup.md)。

> 质量护栏都在 `pyproject.toml` 里：`[tool.ruff]` 的 `target-version = "py311"` 会拦住
> "3.12 才支持、但下限写的是 3.11"的语法（真发生过一次：f-string 表达式里写引号在 3.11 是
> 语法错误，而在本地 3.12 上跑得好好的），`[tool.coverage.report]` 的 `fail_under = 85`
> 让覆盖率掉了就直接红。

## License

MIT —— 见 [`LICENSE`](LICENSE)。
