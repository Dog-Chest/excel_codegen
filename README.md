# excel_codegen

**在 Excel 里填参数，用 YAML + Jinja2 出代码。**

[![ci](https://github.com/Dog-Chest/excel_codegen/actions/workflows/ci.yml/badge.svg)](https://github.com/Dog-Chest/excel_codegen/actions/workflows/ci.yml)
![python](https://img.shields.io/badge/python-3.11%2B-blue)
![license](https://img.shields.io/badge/license-MIT-green)

> 🤖 **本项目由 DeepSeek 自动生成**：代码、测试与文档都在 DeepSeek Harness 里由 AI 编写，
> 人工负责提需求、定方向与验收。

工程师最熟悉的填参数界面是 Excel，最该被评审和版本管理的生成规则却常常散在一次性脚本里。
这个工具把两头接起来：**变量与模板写在 YAML**（进 Git、可 diff、可评审），
**参数在 Excel 里填**（可右拉、可分享），一条命令渲染出代码。

```
example.yaml ──init──► template.xlsx ──(人工填参数)──► render ──► Output 表 / 代码文件
     ▲                                                              │
     └──────────────── Jinja2 模板（Prefix + Value + Suffix）◄───────┘
```

## 快速开始

```bash
# 1) 生成 Excel 表单
uv run excel-codegen init   -c examples/example.yaml -o examples/template.xlsx --cases 2

# 2) 在 Excel 里填：Global Parameter 的 B 列、Local Parameter 的 E/F 列（要更多工况就整块右拉）

# 3) 渲染回 Excel 的 Output 表
uv run excel-codegen render -c examples/example.yaml -x examples/template.xlsx --write-excel
```

没装 uv 就去掉 `uv run`，先按下面「安装」把环境准备好。

## 特性

| | |
| --- | --- |
| **一个 YAML 描述一切** | 全局变量、局部变量、多个输出模板（内联 `code` 或外部 `template_file`） |
| **Excel 就是表单** | 一行一个全局变量；从 E 列起一个工况一列，右拉即增列 |
| **Prefix + Value + Suffix** | `{{ port }}` → `GPIOA_PORT`，`{{ port.value }}` → `A`；空格有意义（`suffix: " m"` 就是 `" m"`） |
| **两种输出引擎** | `snapshot` 写文本快照；`excel` 写 **Excel 公式** —— 改参数后打开 Excel 即重算，**不用再跑脚本** |
| **派生参数** | `derived: "rho * g"` 让参数引用参数（同工况的 local + global），参数表里是活公式 |
| **一本工作簿放多套规则** | `case_filter: "kind == 'EXT'"` 让模板只作用于匹配的工况，共用一张 Global 表 |
| **可导出代码文件** | 文件名支持 Jinja2（`uart_init_{{ case_name }}.c`），一次生成多份 |
| **自带说明与指纹** | 工作簿里有 `HOWTO` 表与生成指纹；`excel-codegen check` 判定"表里的代码是否已过期"，**CI 可用** |

## 现场用例：ABS FPI 内外压 → GeniE

[`abs_fpi/`](abs_fpi/) 是拿这个工具干真活的现场项目 —— 把 **ABS FPI**（5A-3-2/5.5 外压、
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
逐轮试用发现的问题与修复复测记录在 [`abs_fpi/FINDINGS.md`](abs_fpi/FINDINGS.md)。

> 用例里的舱容与工况组合只是**示意数值**（随手取的几个数），不代表任何真实船舶。

## 文档

| 文档 | 内容 |
| --- | --- |
| [`docs/template_guide.md`](docs/template_guide.md) | **模板库扩展指南**：YAML 字段、Prefix/Suffix、Jinja2 速查、Excel 表结构、`case_filter`、排错、FAQ、能力边界、公式模式与派生参数 |
| [`docs/cli.md`](docs/cli.md) | 命令参考（四个命令的全部选项）、作为 Python 库使用、错误类型 |
| [`docs/setup.md`](docs/setup.md) | 安装与跨平台环境：uv 零配置、venv + pip、Ubuntu 与外置盘的坑、Python 版本策略 |
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

## 开发

```bash
uv run pytest                      # 107 项用例
uv run pytest --cov=excel_codegen
```

CI（[配置](.github/workflows/ci.yml)）在 **Windows / macOS / Linux × Python 3.11 / 3.13**
上各跑一遍。支持下限是 **3.11**，开发默认 3.12；改下限的规矩见 [`docs/setup.md`](docs/setup.md)。

## License

MIT —— 见 [`LICENSE`](LICENSE)。
