# 命令参考 / Python 库 / 错误类型

五个命令：`init` → 填写 → `render` → `validate` / `check`；另有 `doctor` 体检。
全局选项：`excel-codegen --version`。

---

## `excel-codegen init`

生成 Excel 表单。

默认（`--prerender`）建完骨架就**顺手把输出也写一遍** —— 公式模式下输出表里是活公式，
所以"`init` 完就能打开 Excel 干活"，不需要再跑一条 `render --write-excel`。
参数还是 YAML 默认值、暂时过不了取值约束或 `asserts` 时会**跳过预填并说明原因**
（骨架照常生成，`init` 不会因此失败）。

| 选项 | 说明 | 默认 |
| --- | --- | --- |
| `-c, --config PATH` | YAML 配置文件（必填，必须存在） | — |
| `-o, --output PATH` | 生成的 Excel 路径 | 配置中的 `excel.output` |
| `--cases VALUE` | 初始工况：数量（`3`）**或逗号分隔的名字**（`EXT-T20.559,INT-T15`）——横向布局是列、纵向是行（`excel.local_direction`） | `2` |
| `-f, --force` | 目标文件已存在时覆盖 | 否 |
| `--template-sheet / --no-template-sheet` | 是否生成隐藏的 `Template` 表（保存模板原文与元信息） | 生成 |
| `--howto / --no-howto` | 是否生成 `HOWTO` 说明表（放在第一张） | 生成 |
| `--comments / --no-comments` | 是否给「变量名」那一格加 **Excel 批注**（描述 / 单位 / 类型 / 约束 / 前缀后缀 / 派生表达式 / 模板里怎么引用） | 生成 |
| `--scripts / --no-scripts` | 是否在工作簿旁边生成**一键刷新脚本** `<工作簿名>_render.bat` / `.sh`（指南 §6.6） | 生成 |
| `--prerender / --no-prerender` | 建完表就把输出写进去（公式模式下工作簿因此开箱可用）；`--no-prerender` 只要骨架 | 预填 |

## `excel-codegen render`

| 选项 | 说明 | 默认 |
| --- | --- | --- |
| `-c, --config PATH` | YAML 配置文件（必填） | — |
| `-x, --excel PATH` | 已填写的 Excel | 配置中的 `excel.output` |
| `-w, --write-excel` | 渲染结果写回 Excel 的 Output 表（并刷新 HOWTO / 指纹） | 否 |
| `-d, --outdir PATH` | 渲染结果导出为代码文件到该目录 | 否 |
| `--case NAME` | 只渲染指定 Case（可重复传入） | 全部 Case |
| `--show / --no-show` | 在终端打印渲染结果 | `--show` |
| `--overwrite / --no-overwrite` | 导出文件已存在时是否覆盖 | `--overwrite` |

不传 `--write-excel` 和 `--outdir` 时只在终端预览（安全模式）；摘要里会明确写出
「写回 Excel │ 否（需要 --write-excel）」，不会让人误以为已经落盘。

## `excel-codegen validate`

| 选项 | 说明 |
| --- | --- |
| `-c, --config PATH` | YAML 配置文件（必填） |
| `-x, --excel PATH` | 顺带校验 Excel 结构：表是否存在、表头是否正确、每个 Case 的取值、上次渲染时间与指纹 |

校验内容：YAML 结构与语义（含重复键）、模板与 `case_filter` 的 Jinja2 语法、
`{% include %}` 片段是否存在（指南 §17）、
模板引用的变量是否已定义、**定义了却没有被引用的变量**、`output_sheet` 是否已声明、
变量是否重名、Case 取值一览。

带 `-x` 时还会校验**参数取值是否满足声明的约束**（`min` / `max` / `choices` / `pattern`，
见指南 §3.5）—— 一次列出全部越界处，指出是哪张表、哪一列、哪个变量。
另外还会把根级 **`asserts`（跨变量校验，指南 §3.6）** 编译一遍（带 `-x` 时连每个 Case
一起求值）—— 一次列出全部越界处，指出是哪张表、哪一列、哪个变量。
配了 `variables.group`（指南 §18）时，成员表的取值约束也会一起查。

## `excel-codegen check`

| 选项 | 说明 |
| --- | --- |
| `-c, --config PATH` | YAML 配置文件（必填） |
| `-x, --excel PATH` | 要检查的 Excel（默认取配置中的 `excel.output`） |
| `--values / --no-values` | 公式模式：把 Output 表的公式**在 Python 里算一遍**再与 Python 渲染比对（默认开） |
| `--json` | 把结果打成 **JSON**（stdout 上只有 JSON，方便 `jq` / CI 看板消费）；退出码与表格模式一致 |

* 快照模式（`engine: snapshot`）：拿当前参数重新渲染，与表内内容逐行比对（含 Case 表头）。
* 公式模式：① 比公式文本（与当前 YAML 是否一致）；② 把公式求值，与 Python 渲染比对
  —— 后者能抓到"列标指错 / 空单元格处理错"这类**公式本身**的问题，以及"插删过 Case 列"这种结构变化。
* **一致 → 退出码 0；过期 → 退出码 1**。差异**按行列出**（每个 Case 最多 5 行，其余折成
  一句"还有 N 行不同"），公式模式还会指出是**模板第几行**。适合放进 CI 断言
  "提交的工作簿与代码是同步的"。
* `--json` 输出的结构：`ok` / `config` / `excel` / `recorded{time,input_fingerprint,output_fingerprint}`
  / `current{...}` / `drift` / `warnings` / `problems`。例如：

  ```bash
  excel-codegen check -c project.yaml --json | jq -r '.problems[]'
  ```
* `render` 与 `check` 都会先校验取值约束（指南 §3.5），越界时以退出码 1 结束 ——
  所以"参数填错"与"输出过期"是两类不同的失败，报错信息里能直接分辨。

---

## 作为 Python 库使用

```python
from excel_codegen import (
    load_config,
    create_template,
    render_all,
    write_results,
    export_files,
    build_environment,
    input_fingerprint,
    output_fingerprint,
    read_metadata,
)

config = load_config("examples/example.yaml")  # 读取并校验 YAML
create_template(config, "template.xlsx", cases=3, overwrite=True)

output = render_all(config, "template.xlsx")  # 读取填写的参数并渲染
write_results("template.xlsx", config, output.results)  # 写回 Output 表 + 指纹
files = export_files(config, output.results, "generated/")  # 导出代码文件

print(output.results["uart_init"][0].text)  # 取某个模板某个 Case 的文本
print(output.skipped)  # case_filter 跳过了哪些 Case
print(input_fingerprint(output.global_values, output.cases))
```

---

## 错误类型

所有面向用户的错误都继承自 `CodeGenError`，CLI 会以 `ERROR 提示` 形式打印并返回退出码 1：

| 异常 | 触发场景 | 提示示例 |
| --- | --- | --- |
| `ConfigError` | YAML 语法错误、同一映射里键重复、字段非法、重名、`output_sheet` 未声明 | `YAML 键重复: 'variables'（第 20 行）…` |
| `ExcelError` | 文件不存在、工作表缺失、表头不对、没有工况、变量名重复 | `Excel 缺少工作表: 'Local Parameter'。当前工作表: ...` |
| `RenderError` | 模板 / `case_filter` 语法错误、变量缺失、导出文件名非法 | `模板 'uart_init' 语法错误：第 3 行: Unexpected end of template ...` |

> **控制台兼容**：CLI 只用 ASCII 标记（`OK` / `ERROR` / `!`），并在启动时把 stdout/stderr 的
> 错误处理设为 `backslashreplace`。中文 Windows（GBK 控制台）下**不会**再出现
> "活干完了、最后一行字打不出来、退出码 1" 的情况。

---

## `excel-codegen doctor`

体检**环境 / 配置 / 工作簿**三层，把常见坑一次说清。**有 ERROR 时退出码 1**（没 ERROR 时
只有 `!` 提示也算通过）。

| 选项 | 说明 |
| --- | --- |
| `-c, --config PATH` | YAML 配置文件（必填） |
| `-x, --excel PATH` | 顺带体检这个工作簿（默认取配置中的 `excel.output`） |

三层各看什么：

| 层 | 检查项 |
| --- | --- |
| 环境 | Python 版本是否 ≥ `requires-python`、六个运行时依赖在不在、有没有 uv / git、是不是在虚拟环境里跑 |
| 配置 | YAML 能否加载（含 `extends`）、模板/变量规模、模板语法与 `{% include %}` 片段、派生参数、`asserts` 语法、**有没有一条约束都没有**、有没有定义了却没人用的变量 |
| 工作簿 | 工作表是否齐全、Case 列 / 行、取值约束与 `asserts` 是否满足、**整行/整列都空的 Case**、渲染记录（指纹是否一致）、一键脚本在不在 |

```bash
excel-codegen doctor -c project.yaml
```

典型用途：**新同事拿到仓库的第一条命令**（"我这儿跑不起来"）、发布前自检、
以及定期确认"这本工作簿还受哪些约束管着"。
