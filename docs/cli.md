# 命令参考 / Python 库 / 错误类型

四个命令：`init` → 填写 → `render` → `validate` / `check`。
全局选项：`excel-codegen --version`。

---

## `excel-codegen init`

生成 Excel 表单骨架。

| 选项 | 说明 | 默认 |
| --- | --- | --- |
| `-c, --config PATH` | YAML 配置文件（必填，必须存在） | — |
| `-o, --output PATH` | 生成的 Excel 路径 | 配置中的 `excel.output` |
| `--cases VALUE` | 初始 Case 列：数量（`3`）**或逗号分隔的名字**（`EXT-T20.559,INT-T15`） | `2` |
| `-f, --force` | 目标文件已存在时覆盖 | 否 |
| `--template-sheet / --no-template-sheet` | 是否生成隐藏的 `Template` 表（保存模板原文与元信息） | 生成 |
| `--howto / --no-howto` | 是否生成 `HOWTO` 说明表（放在第一张） | 生成 |

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
模板引用的变量是否已定义、**定义了却没有被引用的变量**、`output_sheet` 是否已声明、
变量是否重名、Case 取值一览。

## `excel-codegen check`

| 选项 | 说明 |
| --- | --- |
| `-c, --config PATH` | YAML 配置文件（必填） |
| `-x, --excel PATH` | 要检查的 Excel（默认取配置中的 `excel.output`） |
| `--values / --no-values` | 公式模式：把 Output 表的公式**在 Python 里算一遍**再与 Python 渲染比对（默认开） |

* 快照模式：拿当前参数重新渲染，与表内内容逐行比对（含 Case 表头）。
* 公式模式：① 比公式文本（与当前 YAML 是否一致）；② 把公式求值，与 Python 渲染比对
  —— 后者能抓到"列标指错 / 空单元格处理错"这类**公式本身**的问题，以及"插删过 Case 列"这种结构变化。
* **一致 → 退出码 0；过期 → 退出码 1**，并列出第一处差异（公式模式会指出是**模板第几行**）。
  适合放进 CI 断言"提交的工作簿与代码是同步的"。

---

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

---

## 错误类型

所有面向用户的错误都继承自 `CodeGenError`，CLI 会以 `ERROR 提示` 形式打印并返回退出码 1：

| 异常 | 触发场景 | 提示示例 |
| --- | --- | --- |
| `ConfigError` | YAML 语法错误、同一映射里键重复、字段非法、重名、`output_sheet` 未声明 | `YAML 键重复: 'variables'（第 20 行）…` |
| `ExcelError` | 文件不存在、工作表缺失、表头不对、无 Case 列、变量名重复 | `Excel 缺少工作表: 'Local Parameter'。当前工作表: ...` |
| `RenderError` | 模板 / `case_filter` 语法错误、变量缺失、导出文件名非法 | `模板 'uart_init' 语法错误：第 3 行: Unexpected end of template ...` |

> **控制台兼容**：CLI 只用 ASCII 标记（`OK` / `ERROR` / `!`），并在启动时把 stdout/stderr 的
> 错误处理设为 `backslashreplace`。中文 Windows（GBK 控制台）下**不会**再出现
> "活干完了、最后一行字打不出来、退出码 1" 的情况。
