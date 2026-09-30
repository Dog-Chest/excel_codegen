# 内置示例

示例**随包发布**（wheel 与 sdist 里都有），所以装了 `spreadsheet-codegen` 就够了，
不需要克隆仓库：

```bash
spreadsheet-codegen examples                     # 列出示例、各自演示什么
spreadsheet-codegen examples --copy ./examples   # 拷到 ./examples（--only abs_fpi 只拷一个）
```

拷出来之后**打开里面的 `.xlsx` 就能改** —— 每一本都已经填好样例数值，而且默认是
**公式模式**：改一格参数，`Output` 表里的代码自己就重算，不需要跑任何命令。
要让「导出的代码文件」跟着更新，才需要回到命令行（或双击同目录的 `*_render.bat` / 跑 `*_render.sh`）。

| 示例 | 入口 YAML | 工作簿 | 演示什么 |
| --- | --- | --- | --- |
| [`basic/`](basic/) | `example_formula.yaml` | `template_formula.xlsx` | **默认的公式模式**：改参数在 Excel 里自动重算 |
| | `example.yaml` | `template.xlsx` | 显式 `engine: snapshot`：过滤器 / 循环 / 导出文件 |
| [`nastran/`](nastran/) | `nastran_case_control.yaml` | `nastran_case_control.xlsx` | **一行一个工况**（`local_direction: vertical`）生成工况控制语句（指南 §19 / §20） |
| [`abs_fpi/`](abs_fpi/) | `abs_fpi_internal.yaml` | `abs_fpi_internal.xlsx` | **现场用例**：ABS FPI 内外压 → GeniE；多规则集 / 成员表（指南 §18） |

> `generated*/` 里的文件是**提交进仓库的实测产物**（脚本导出一次的结果），
> 拿它跟你自己跑出来的对一下，就知道有没有改坏。

---

## basic —— 入门：公式模式与快照模式

```bash
cd basic

# 工作簿已经生成好了，直接打开 template_formula.xlsx 改参数即可。
# 要把代码导成文件：
spreadsheet-codegen render -c example_formula.yaml -x template_formula.xlsx --outdir generated_formula

# 核对「Excel 里算出来的」与「Python 渲染的」是否一致（CI 里很有用）：
spreadsheet-codegen check -c example_formula.yaml -x template_formula.xlsx
```

`example.yaml` 是**对照示例**：它显式写 `engine: snapshot`，用到了公式模式表达不了的
`{% for %}` 与过滤器，同时演示 `filename` 导出多个文件。两条路的取舍见指南 §14。

## nastran —— 一行一个工况

```bash
cd nastran
spreadsheet-codegen render -c nastran_case_control.yaml -x nastran_case_control.xlsx --outdir generated
```

Local 表是**竖着**的（`local_direction: vertical`）：一行一个工况，下拉即增行，
整块从别处粘进来也方便。每条语句**留空就不输出**，而它在表里占的那一行仍然是空的
（方便整列复制）。输出是 `.inc` / `.deck` 工况控制语句，外加每块的摘要。

## abs_fpi —— 现场用例：ABS FPI 内外压 → GeniE

把 **ABS FPI**（5A-3-2/5.5 外压、5.7 内压）的面载荷计算，生成可直接粘进 GeniE 的
JavaScript 函数体。详见 [`abs_fpi/README.md`](abs_fpi/README.md)。

```bash
cd abs_fpi

# 三本工作簿，挑一本打开就能改：
#   abs_fpi_external.xlsx     外压单用（3 个工况）
#   abs_fpi_internal.xlsx     内压单用（4 个工况）
#   ABS_FPI_load_cases.xlsx   项目工作簿：两套规则共用一张 Global 表（7 个工况）
spreadsheet-codegen render -c abs_fpi_internal.yaml -x abs_fpi_internal.xlsx --outdir generated
```

> 用例里的舱容与工况组合只是**示意数值**（随手取的几个数），不代表任何真实船舶。
