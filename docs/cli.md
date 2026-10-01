# 命令参考 / Python 库 / 错误类型

五个命令：`init` → 填写 → `render` → `validate` / `check`；另有 `doctor` 体检。
全局选项：`spreadsheet-codegen --version`。

---

## `spreadsheet-codegen init`

生成 Excel 表单。

默认（`--prerender`）建完骨架就**顺手把输出也写一遍** —— 公式模式下输出表里是活公式，
所以"`init` 完就能打开 Excel 干活"，不需要再跑一条 `render --write-excel`。
参数还是 YAML 默认值、暂时过不了取值约束或 `asserts` 时会**跳过预填并说明原因**
（骨架照常生成，`init` 不会因此失败）。

| 选项 | 说明 | 默认 |
| --- | --- | --- |
| `-c, --config PATH` | YAML 配置文件（必填，必须存在） | — |
| `-o, --output PATH` | 生成的 Excel 路径（相对**当前工作目录**） | 配置中的 `excel.output`（相对**配置文件所在目录**，见下） |
| `--cases VALUE` | 初始工况：数量（`3`）**或逗号分隔的名字**（`EXT-T20.559,INT-T15`）——横向布局是列、纵向是行（`excel.local_direction`） | `2` |
| `-f, --force` | 目标文件已存在时覆盖（**会丢掉已填的参数**，见下） | 否 |
| `-y, --yes` | 配合 `--force`：确认"这本工作簿里填过参数也要重建" | 否 |
| `--template-sheet / --no-template-sheet` | 是否生成隐藏的 `Template` 表（保存模板原文与元信息） | 生成 |
| `--howto / --no-howto` | 是否生成 `HOWTO` 说明表（放在第一张） | 生成 |
| `--comments / --no-comments` | 是否给「变量名」那一格加 **Excel 批注**（描述 / 单位 / 类型 / 约束 / 前缀后缀 / 派生表达式 / 模板里怎么引用） | 生成 |
| `--scripts / --no-scripts` | 是否在工作簿旁边生成**一键刷新脚本** `<工作簿名>_render.bat` / `.sh`（指南 §6.6） | 生成 |
| `--prerender / --no-prerender` | 建完表就把输出写进去（公式模式下工作簿因此开箱可用）；`--no-prerender` 只要骨架 | 预填 |

**`--force` 是数据丢失点**：重建会把已填的参数换成 YAML 默认值。所以工具会先看一眼这本
工作簿里的参数**是否已经不是默认值**（比参数指纹）：

```text
ERROR 这本工作簿里 2 个 Case 的参数已经不是 YAML 默认值（当前指纹 b7463bafa13c，默认值应为
      914186c24e51）—— 看起来人填过东西。
  重建会把这些参数换成 YAML 默认值（原文件不会自动备份）。
  → 确认要丢掉就加 --yes；只想改骨架不想丢数据，就别加 --force
```

刚生成、没人动过的工作簿**不会被拦**（老用法 `init --force` 照常一条命令跑完）。

**路径规则**：`excel.output` 相对**配置文件所在目录**解析（与 `template_file` 同一条规则，
`..` 也会被折掉）；命令行给的 `-o` / `-x` 相对**当前工作目录**。摘要里一律打印规范化后的
**绝对路径**，所以"文件到底写哪去了"一眼能看到（指南 §2.2）。

## `spreadsheet-codegen render`

| 选项 | 说明 | 默认 |
| --- | --- | --- |
| `-c, --config PATH` | YAML 配置文件（必填） | — |
| `-x, --excel PATH` | 已填写的 Excel | 配置中的 `excel.output` |
| `-w, --write-excel` | 渲染结果写回 Excel 的 Output 表（并刷新 HOWTO / 指纹） | 否 |
| `-d, --outdir PATH` | 渲染结果导出为代码文件到该目录 | 否 |
| `--case NAME` | 只渲染指定 Case（可重复传入） | 全部 Case |
| `--show / --no-show` | 在终端打印渲染结果 | `--show` |
| `--overwrite / --no-overwrite` | 导出文件已存在时是否覆盖 | `--overwrite` |
| `--allow-overwrite-filename` | 放行"多份结果写进同一个文件名"（默认拦住，见指南 §8.1） | 否 |

`--write-excel` 与 `--outdir` **可以一起给**：一次渲染既刷新 Excel 又导出代码文件 ——
"改了参数想把两边都同步"不用跑两遍。

不传 `--write-excel` 和 `--outdir` 时只在终端预览（安全模式）；摘要里会明确写出
「写回 Excel │ 否（需要 --write-excel）」与导出目录的**绝对路径**，不会让人误以为已经落盘。

## `spreadsheet-codegen validate`

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

## `spreadsheet-codegen check`

| 选项 | 说明 |
| --- | --- |
| `-c, --config PATH` | YAML 配置文件（必填） |
| `-x, --excel PATH` | 要检查的 Excel（默认取配置中的 `excel.output`） |
| `--values / --no-values` | 公式模式：把 Output 表的公式**在 Python 里算一遍**再与 Python 渲染比对（默认开） |
| `--json` | 把结果打成 **JSON**（stdout 上只有 JSON，方便 `jq` / CI 看板消费）；退出码与表格模式一致 |

* 快照模式（`engine: snapshot`）：拿当前参数重新渲染，与表内内容逐行比对（含 Case 表头）。
* 公式模式：① 比公式文本（与当前 YAML 是否一致）；② 把公式求值，与 Python 渲染比对
  —— 后者能抓到"列标指错 / 空单元格处理错"这类**公式本身**的问题，以及"插删过 Case 列"这种结构变化。
* **一致 → 退出码 0；过期 → 退出码 3**（见下面的退出码表）。差异**按行列出**（每个 Case 最多 5 行，
  其余折成一句"还有 N 行不同"），公式模式还会指出是**模板第几行**。适合放进 CI 断言
  "提交的工作簿与代码是同步的"。
* `--json` 输出的结构：`ok` / `config` / `excel` / `recorded{time,input_fingerprint,output_fingerprint}`
  / `current{...}` / `drift` / `warnings` / `problems` / `problem_kinds`。例如：

  ```bash
  spreadsheet-codegen check -c project.yaml --json | jq -r '.problems[]'
  spreadsheet-codegen check -c project.yaml --json | jq -r '.problem_kinds[]'   # 按性质分流
  ```
* `problems` 是**人读的话术**，`problem_kinds` 是**稳定枚举**（同一个下标一一对应）：

  | `kind` | 含义 |
  | --- | --- |
  | `output_stale` | 输出表内容与当前 YAML / 参数不一致 → 重跑 `--write-excel` |
  | `value_mismatch` | 公式文本一致，但"公式算出来的文本"与 Python 渲染不同（结构改过？） |
  | `case_header` | 输出表的 Case 表头与当前参数表对不上 |
  | `missing_sheet` | 输出表不存在 |
  | `no_matching_case` | `case_filter` 把所有 Case 都跳过了，旧内容无法核对 |

  CI 里可以据此分流（例如只在 `output_stale` 时自动重渲染，`value_mismatch` 直接失败）。

---

## 退出码（CI 里怎么分流）

0.11.0 起不同性质的失败给出不同退出码 —— 以前全是 1，CI 里想区分"我的参数填错了"
与"表里的输出过期了"只能去 grep 文本，而这两件事的处置完全不同：

| 退出码 | 含义 | 谁会给 |
| --- | --- | --- |
| `0` | 成功 | 全部命令 |
| `1` | 通用失败（模板语法、变量缺失、工作簿结构不对） | 全部命令 |
| `2` | **配置 / 取值错误**：YAML 写错、取值越界、`asserts` 不满足、模板写法公式模式表达不了 | `init` / `render` / `validate` / `check` / `doctor` |
| `3` | **输出过期**：输出表内容与当前参数不一致 | `check` |
| `4` | **环境问题**：工作簿打不开 / 写不进（被 Excel 占用等） | `init` / `render` / `validate` / `check` / `doctor` |

```bash
spreadsheet-codegen check -c project.yaml || case $? in
  3) echo "输出过期，自动重渲染"; spreadsheet-codegen render -c project.yaml -x book.xlsx --write-excel ;;
  2) echo "参数或配置不对，得人来改"; exit 1 ;;
  *) echo "其他失败"; exit 1 ;;
esac
```

> 退出码 `1` 仍然覆盖"其他一切失败"，所以老脚本里的 `if errorlevel 1` 照常有效。

**`doctor` 是多层汇总，一次可能会报出好几类问题**：它取**数值最大**的那一类
（`4` 环境 > `2` 配置 > `1` 其他）—— 工作簿都读不了的时候，先解决那个才有意义。

| `doctor` 遇到的情况 | 退出码 |
| --- | --- |
| YAML 加载失败（语法错、字段非法、重名……） | `2` |
| 工作簿打不开 / 读不了（被占用、文件损坏） | `4` |
| 其余有 ERROR（缺表、表头不对、取值越界、`asserts` 不满足……） | `1` 或 `2`（取值类算 `2`） |

> 坏掉的工作簿**只会变成报告里的一行 ERROR**，不会甩 traceback —— 跑 `doctor` 的时候，
> 工作簿往往正好就是坏的。

---

## 作为 Python 库使用

```python
from spreadsheet_codegen import (
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

config = load_config("examples/basic/example.yaml")  # 读取并校验 YAML（示例：spreadsheet-codegen examples --copy .）
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

所有面向用户的错误都继承自 `CodeGenError`，CLI 会以 `ERROR 提示` 形式打印并按
[退出码表](#退出码ci-里怎么分流)返回（0.11.0 起不再全是 1）：

| 异常 | 触发场景 | 退出码 | 提示示例 |
| --- | --- | --- | --- |
| `ConfigError` | YAML 语法错误、同一映射里键重复、字段非法、重名、`output_sheet` 未声明 | 2 | `YAML 键重复: 'variables'（第 20 行）…` |
| `InputError` | **表里填的取值不对**：取值越界、`choices` 不匹配、`asserts` 不满足、成员名不存在 | 2 | `参数取值不满足变量声明的约束，共 2 处：…` |
| `ExcelError` | 文件不存在、工作表缺失、表头不对、没有工况、变量名重复 | 1 | `Excel 缺少工作表: 'Local Parameter'。当前工作表: ...` |
| `WorkbookIOError` | **工作簿读不出来 / 写不进去**：文件损坏、被 Excel 占用、没有权限 | 4 | `无法读取 Excel 文件 book.xlsx: File is not a zip file` |
| `RenderError` | 模板 / `case_filter` 语法错误、变量缺失、导出文件名非法 | 1 | `模板 'uart_init' 语法错误：第 3 行: Unexpected end of template ...` |
| `FormulaError` | 模板写法公式模式表达不了（`{% for %}`、过滤器、跨行分支） | 2 | `公式模式不支持 {% for %}：…请把该模板改回 engine: snapshot` |

`InputError` 与 `WorkbookIOError` 都是 `ExcelError` 的**子类**（0.11.0 起），所以库调用方
既有的 `except ExcelError` 照常工作 —— 它们只是把"取值不对"（2）与"工作簿读写不了"（4）
单独标出来，好让 CLI 给出可区分的退出码。

> **文件不存在**（还没跑 `init`）不算环境问题，走普通的退出码 `1` —— 那是"忘了生成"，
> 不是"环境坏了"，混进 4 只会让 CI 分流失效。

> **控制台兼容**：CLI 只用 ASCII 标记（`OK` / `ERROR` / `!`），并在启动时把 stdout/stderr 的
> 错误处理设为 `backslashreplace`。中文 Windows（GBK 控制台）下**不会**再出现
> "活干完了、最后一行字打不出来、退出码 1" 的情况。

---

## `spreadsheet-codegen doctor`

体检**环境 / 配置 / 工作簿**三层，把常见坑一次说清。有 ERROR 时按
[退出码表](#退出码ci-里怎么分流)返回（YAML 加载失败 `2`、工作簿打不开 `4`、其余 `1`）；
没 ERROR 时只有 `!` 提示也算通过。

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
spreadsheet-codegen doctor -c project.yaml
```

典型用途：**新同事拿到仓库的第一条命令**（"我这儿跑不起来"）、发布前自检、
以及定期确认"这本工作簿还受哪些约束管着"。

---

## `spreadsheet-codegen examples`

列出**随包发布的内置示例**，或把它们拷出来直接用。示例是包数据（wheel 与 sdist 里都有），
所以 `pip install spreadsheet-codegen` / `uv tool install spreadsheet-codegen` 的人**不需要克隆仓库**。

| 选项 | 说明 |
| --- | --- |
| `-o, --copy DIR` | 把示例拷到 DIR（每个示例一个子目录）；不带就是只列出 |
| `--only NAME` | 只处理某一个示例（`basic` / `nastran` / `abs_fpi`） |
| `--force` | 目标目录已存在且非空时覆盖（默认拒绝，避免覆盖你改过的文件） |

```bash
spreadsheet-codegen examples                       # 看有哪些、各自演示什么
spreadsheet-codegen examples --copy ./examples     # 全部拷出来
spreadsheet-codegen examples --copy . --only abs_fpi   # 只要现场用例那个
```

拷出来的每个示例都带一本**已经填好样例参数**的工作簿，打开就能改；
公式模式下改完自动重算，不需要跑命令。示例清单与说明见
[`spreadsheet_codegen/examples/README.md`](../spreadsheet_codegen/examples/README.md)。

> 拷出来的目录是**你的**：`--force` 之外不会覆盖已有内容，改坏了再拷一份就行。
