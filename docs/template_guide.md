# 模板库扩展指南（template_guide）

本文档面向**维护模板库 / 新增生成目标**的工程师，覆盖：YAML 字段参考、变量定义、Prefix/Value/Suffix 语义、
Jinja2 语法速查、Excel 表结构与填写规则、输出布局、导出代码文件、
**公式模式（改参数免重跑脚本）**、常见问题与排错。

> 只想跑通流程？先看 [README 的快速开始](../README.md#快速开始)。
> 想让"改参数不用跑脚本"？直接看 [§14 公式模式](#14-公式模式engine-excel改参数免重跑脚本)。

---

## 1. 心智模型

```
        YAML（模板与变量定义，进 Git）                        Excel（参数填写，给人用）
┌────────────────────────────────────────┐        ┌───────────────────────────────────┐
│ variables: global / local              │        │ Global Parameter : 全局变量取值    │
│ templates: code + output_sheet + 布局  │  init  │ Local  Parameter : 每个 Case 一列  │
│            engine: snapshot | excel    │        │                                   │
└────────────────────────────────────────┘ ─────► └───────────────────────────────────┘
                     │                                         │ 人工填写
                     │                                         ▼
                     └──────────────► render ◄─────────────────┘
                                       │
                     ┌─────────────────┴──────────────────┐
                     ▼                                    ▼
            Output 工作表：文本快照（snapshot）      生成的代码文件（--outdir）
                          或公式（excel，Excel 自己重算）
```

三条不变式：

1. **变量只有两种作用域**：`global`（所有 Case 共用一份取值）、`local`（每个 Case 一列取值）。
2. **渲染上下文 = global + local + `case_name`**，其中 local 覆盖同名 global。
3. **模板只在渲染时求值**：模板里引用但上下文没有的变量会直接报错（`StrictUndefined`），不会静默变空串。

---

## 2. YAML 字段完整参考

### 2.1 根节点

| 字段 | 类型 | 必填 | 默认 | 说明 |
| --- | --- | --- | --- | --- |
| `version` | int | 否 | `1` | 配置版本，必须 ≥ 1 |
| `extends` | list[str] | 否 | `[]` | 先合并这些 YAML（路径相对本文件），见 §16 |
| `excel` | mapping | 否 | 见下 | Excel 相关配置 |
| `variables` | mapping | 否 | 空 | 变量定义：`global` / `local` |
| `templates` | list | **是** | — | 至少一个模板 |

### 2.2 `excel`

| 字段 | 类型 | 默认 | 说明 |
| --- | --- | --- | --- |
| `output` | str | `template.xlsx` | `init` 默认输出路径（相对当前工作目录） |
| `template_sheet` | str \| null | `Template` | 隐藏参考表的名字（模板原文 + `## excel-codegen-meta` 元信息）；设为 `null` 不生成 |
| `howto_sheet` | str \| null | `HOWTO` | 使用说明表，`create_template` 生成、`write_results` 刷新，**放第一张**；设为 `null` 不生成 |
| `sheets.global` | str | `Global Parameter` | 全局表名（≤31 字符，不含 `[]:*?/\`） |
| `sheets.local` | str | `Local Parameter` | 局部表名 |
| `sheets.outputs` | list[str] | `["Output"]` | 所有输出表名；模板的 `output_sheet` 必须在此声明 |

`template_sheet` / `howto_sheet` 不能与 Global / Local / Output 表名冲突（配置阶段就会报错）。

### 2.3 `variables.global[]` / `variables.local[]`

| 字段 | 类型 | 默认 | 说明 |
| --- | --- | --- | --- |
| `name` | str | — | **必填**。必须是合法标识符 `[A-Za-z_][A-Za-z0-9_]*`，不能是保留名 `case_name` / `template_name` |
| `description` | str | `""` | 写在 Excel 的描述列，纯注释 |
| `default` | any | `""` | 默认值：Excel 单元格留空时使用；`init` 时预填进单元格。**与 `derived` 互斥** |
| `prefix` | str | `""` | 组合值前缀（写入 Excel 的 Prefix 列，可在 Excel 中覆盖） |
| `suffix` | str | `""` | 组合值后缀 |
| `type` | enum | `auto` | `auto` / `string` / `int` / `float` / `bool` / `raw`，见 3.2 |
| `derived` | str \| null | `null` | **派生参数**：一段表达式，引用同 Case 的 local 与 global（见 §15）。写了它就不用填值 |
| `min` / `max` | number \| null | `null` | **取值约束**：数值上下限（`type` 为 `string` / `bool` / `raw` 时不允许）。见 3.5 |
| `choices` | list \| null | `null` | **取值约束**：允许的取值集合，按文本比较，会写成 Excel 下拉列表。见 3.5 |
| `pattern` | str \| null | `null` | **取值约束**：整串匹配的正则（`re.fullmatch`）。见 3.5 |

同名变量在 `global` 与 `local` 中重复是允许的（local 覆盖 global），但 `validate` 会给出警告；
同一作用域内重名会直接报错。

### 2.4 `templates[]`

| 字段 | 类型 | 默认 | 说明 |
| --- | --- | --- | --- |
| `name` | str | — | **必填**，模板唯一名（也用于文件名默认值和日志） |
| `description` | str | `""` | 注释 |
| `output_sheet` | str | `Output` | 写回哪张表；必须在 `excel.sheets.outputs` 中声明 |
| `start_cell` | str | `B2` | 结果起始单元格，如 `B2`、`A1`、`H5` |
| `direction` | `horizontal` \| `vertical` | `horizontal` | 输出布局，见第 6 节 |
| `write_case_headers` | bool | `true` | 是否写出 Case 名表头（横向写在上一行，纵向写在左一列） |
| `filename` | str \| null | `null` | 导出文件名，支持 Jinja2（可用 `{{ case_name }}` / `{{ template_name }}`） |
| `extension` | str | `.txt` | 未配置 `filename` 时的默认扩展名（自动补 `.`） |
| `code` | str | — | 内联模板源码（与 `template_file` 二选一） |
| `template_file` | str | — | 外部模板文件路径，相对 YAML 所在目录解析 |
| `engine` | `snapshot` \| `excel` | `snapshot` | **输出引擎**：`snapshot` 写文本快照；`excel` 写 Excel 公式（改参数免重跑脚本）。见 §14 |
| `case_filter` | str \| null | `null` | 可选：Jinja2 **表达式**，对每个 Case 求值；为假则本模板跳过该 Case。见 §9.2 |

`code` 与 `template_file` **必须提供其一**，两者都写时以 `template_file` 为准。

`case_filter` 写的是表达式本身（不要加 `{{ }}`），例如 `kind == 'EXT'`、`draft.value > 20`。

---

## 3. 变量定义进阶

### 3.1 global 还是 local

| 场景 | 用哪种 |
| --- | --- |
| 所有 Case 共用的值（波特率、目标芯片、项目名、版本号） | `global` |
| 每个 Case 不同的值（端口、模式、通道号、外设名） | `local` |
| 想按 Case 变化的前缀（每个 Case 不同前缀） | 见 4.4：再加一个局部变量承载前缀 |

### 3.2 `type` 与取值转换

Excel 单元格的值按 `type` 转换后再进入上下文：

| `type` | 行为 | 示例（Excel 里写） | 上下文中 |
| --- | --- | --- | --- |
| `auto`（默认） | 保留 openpyxl 读到的原始类型 | `115200` | `int 115200` |
| `string` | 一律转字符串 | `115200` | `"115200"` |
| `int` | 转 int（允许 `"9600"`、`9600.0`） | `9600` | `9600` |
| `float` | 转 float | `1.5` | `1.5` |
| `bool` | `1/true/yes/y/on` → True；`0/false/no/n/off/空` → False | `true` | `True` |
| `raw` | 与 `auto` 相同，语义上表示"不要动它" | 任意 | 原值 |

转换失败会抛出明确错误，例如：`变量 Case1.baud 的值 'abc' 无法转换为 int`。

渲染成文本时，整数浮点会被规范化成整数字符串：`115200.0` → `115200`（超过 `1e16` 量级保留科学计数法）；
布尔值输出 `true` / `false`；`None` 输出空串。这条规范化由 Jinja2 的 `finalize` 统一施加，
所以 `{{ x }}`、`{{ x.value }}`、`{{ 340.0 }}` 三条路径的形态一致 —— 做 diff / golden file 时不会出现
"同一数值两种写法"的假差异（`|float`、`|int`、`"%.2f"|format` 等过滤器不受影响）。

### 3.3 变量名规则（重要）

变量名会**直接作为 Jinja2 变量名**使用，所以必须是合法标识符：

```yaml
name: uart_port      # ✅
name: uart-port      # ❌ 报错：不是合法标识符
name: case_name      # ❌ 报错：保留名（渲染时由程序注入）
```

### 3.4 从 Excel 新增"计划外"变量

`Global Parameter` / `Local Parameter` 表里多出来的变量行，即使 YAML 未定义也会被读进上下文
（按 `auto` 类型）。这让临时参数不用改 YAML 就能用；但 `validate` 会对"模板引用了但 YAML 未定义"的名字给出警告，
建议确认无误后补回 YAML，避免参数长期游离在配置之外。

### 3.5 取值约束：`min` / `max` / `choices` / `pattern`

工具会检查"变量有没有定义、类型对不对"，但**不看值** —— 把吃水 `20.559` 手滑打成 `205.59`，
模板照样渲染，`check` 也只会说"与参数一致"。取值约束就是为了堵这一类**静默的错误交付物**。

```yaml
variables:
  global:
    - name: draft
      type: float
      default: 20.559
      min: 0
      max: 50                 # 数值范围
    - name: mcu
      default: "STM32F103"
      pattern: "STM32[A-Z][0-9]{3}[A-Z]{0,2}"     # 整串匹配的正则
  local:
    - name: kind
      choices: ["EXT", "INT"] # 枚举；Excel 里变成下拉列表
    - name: n_cyl
      type: int
      min: 1
      max: 8
```

**两道闸，判据是第二道**：

| 闸 | 在哪 | 挡什么 |
| --- | --- | --- |
| Excel 数据有效性 | `init` 时写进工作簿 | 下拉列表 / 数值范围，**帮人填对**。`choices` → 下拉；`min` / `max` → 数值范围（`type: int` 用整数校验） |
| `check_value_constraints` | `render` / `validate` / `check` 每次读表后 | **判据**。挡住粘贴、脚本写入、别人发来的老文件，以及 Excel 校验表达不了的情况（正则） |

几个要点：

* **空值也算不合格**。声明了约束就意味着"必须给一个合法取值"。这正好治了"新插一个空 Case 列
  悄悄落进某个规则集"的老问题（见 §9.2）：`kind` 一旦有 `choices` 又没给 `default`，
  新列在渲染时就会报"取值（空）不在允许列表里"。
* **按文本比较**：`choices: [1, 2]` 里填 `1.0` 也算命中（整数浮点会规范化）。
* **`min` / `max` 只能给数值类型**（`int` / `float` / `auto`）。写成 `type: string` 又加范围会在配置阶段报错。
* **正则只由工具检查**：Excel 的数据有效性没有正则，所以 `pattern` 不会写进工作簿（写一个乱报错的校验不如不写）。
* **`default` 也要满足约束**（YAML 自己写的值写错了一样是配置错误）。反过来，**不给 `default` 是允许的** ——
  那表示"这一格必须去表里填"。
* **派生参数不能加约束**：它的值是算出来的，请在表达式里约束（`derived: "max(x, 0)"`）。
* **下拉列表的两个硬限制**（超出会直接报错，不会静默失效）：选项里不能有逗号；所有选项拼起来不超过 255 字符。

报错长这样，一次列出全部问题并指名道姓：

```
ERROR 参数取值不满足变量声明的约束，共 2 处：
  Global Parameter 第 2 行 'draft'：取值 205.59 大于上限 50
  Local Parameter 第 F 列 'EXT-T15' 第 2 行 'kind'：取值 （空） 不在允许列表 EXT/INT 里
  → 改 Excel 里的取值，或放宽 YAML 里的 min / max / choices / pattern
```

---

## 4. Prefix / Value / Suffix 详解

### 4.1 数据模型

每个变量在上下文中都是一个 `VarValue`：

```python
VarValue(value="A", prefix="GPIO", suffix="_PORT")
str(v)        # "GPIOA_PORT"   组合值 = prefix + value + suffix
v.value       # "A"            纯值（保留原始类型）
v.prefix      # "GPIO"
v.suffix      # "_PORT"
v.text        # "A"            纯值的字符串形式
v.is_empty    # False          值/前缀/后缀全空时为 True
```

### 4.2 模板里的四种写法

```jinja
{{ port }}                  {# GPIOA_PORT —— 组合值 #}
{{ port.value }}            {# A          —— 纯值 #}
{{ port.prefix }}           {# GPIO       #}
{{ port.suffix }}           {# _PORT      #}
```

### 4.3 过滤器 `pvs` / `wrap`

对任意表达式做一次组合（两者完全等价，`pvs` = prefix-value-suffix，`wrap` 更语义化）：

```jinja
{{ "A" | pvs("GPIO", "_PORT") }}           {# GPIOA_PORT #}
{{ 115200 | pvs("BAUD_", "U") }}           {# BAUD_115200U #}
{{ port.value | wrap("GPIO", "_PORT") }}   {# GPIOA_PORT #}
{{ mode.value | pvs("MODE_") }}            {# MODE_TX_RX（suffix 可省略） #}
```

`pvs` / `wrap` 也接受变量本身：`{{ port | pvs("X", "Y") }}` → `XGPIOA_PORTY`（组合值再包一层）。
一般对 `xxx.value` 使用更符合直觉。

### 4.4 Prefix/Suffix 的作用范围

**前缀/后缀是按"变量"配置的，不是按 Case。** Local 表的 C/D 列在变量行上，所有 Case 共用。

如果确实需要"不同 Case 不同前缀"，推荐做法是**再加一个局部变量承载前缀**：

```yaml
variables:
  local:
    - name: port_prefix
      description: "端口前缀"
      default: "GPIO"
    - name: port_num
      description: "端口号"
      default: "A"
templates:
  - name: uart_init
    code: |
      UART_Init({{ port_prefix }}{{ port_num }}, {{ mode }});
```

或者直接把完整名字填进单元格（`GPIOB_PORT`），把 `prefix` / `suffix` 留空。

### 4.5 空值语义与回落顺序

回落只看**单元格是不是真的空**，非空值原样使用 —— **首尾空格是有意义的**：

| 情况 | 结果 |
| --- | --- |
| Excel 单元格留空（真空白，`None`） | 使用 YAML `default` |
| Excel 单元格只敲了空格 | 视为未填写 → 使用 YAML `default` |
| Excel 单元格填了值 | 使用单元格值（原样，不 strip） |
| Prefix/Suffix 单元格留空 | 使用 YAML `prefix` / `suffix` |
| Prefix/Suffix 单元格非空 | 原样使用，**保留空格** |

所以"值 + 单位"可以直接用后缀表达：

```yaml
- name: L
  default: 340
  suffix: " m"        # 注意前导空格
```

```jinja
var L = {{ L }};       {# -> var L = 340 m;  而不是 340m #}
```

> 注意两点：
> 1. 想"清掉" YAML 里的前缀/后缀，需要改 YAML（把 `prefix` / `suffix` 删掉或置空），
>    因为 Excel 里的空格子被视为"未填写"。
> 2. 反过来，如果你**不想要**那个空格，就别在 YAML 里写 —— 空格不会再被静默吃掉。

---

## 5. Jinja2 语法速查

默认环境：`StrictUndefined`（变量缺失即报错）、不自动转义（生成代码/文本，不需要 HTML 转义）、保留结尾换行。

### 5.1 基础

```jinja
{{ baud }}                          {# 输出（VarValue → 组合值） #}
{{ baud.value }}                    {# 属性访问：纯值 #}
{{ case_name }}                     {# 当前 Case 名，如 Case1 #}
{{ template_name }}                 {# 当前模板名（也可用于 filename） #}
{{ "text" }}                        {# 字面量 #}
{{ 115200 + 1 }}                    {# 表达式：115201 #}
{{ "abc" | upper }}                 {# 内置过滤器：ABC #}
```

### 5.2 控制结构

```jinja
{% if baud.value | int > 9600 %}
// 高速模式
{% elif baud.value | int > 0 %}
// 低速模式
{% else %}
// 未配置
{% endif %}

{% for i in range(1, mode.value | int + 1) %}
  CH{{ i }}_Init();
{% endfor %}

{% set name = "UART" ~ port.value %}   {# 局部变量，不会污染上下文 #}
// {{ name }}

{# 这是注释，不会出现在输出里 #}
```

### 5.3 空白控制

YAML 的块标量（`code: |`）会保留换行。若控制标签产生多余空行，用 `-` 去除：

```jinja
{%- for i in range(3) %}
line {{ i }}
{%- endfor %}
```

> ⚠ **空行会各占一行**。渲染结果按"一行一个单元格"写进 Excel（横向布局是一列一个 Case），
> 所以 `{% if %}` / `{% for %}` 不加空白控制时多出来的空行，会原样出现在表里 ——
> 对"整列复制粘进编辑器"的用法是实打实的噪音。写控制结构时优先用 `{%- ... -%}`，
> 渲染完先 `--show` 看一眼再 `--write-excel`。

需要输出**字面量** `{{` / `{%`（例如在生成 Jinja 模板的模板）时，用 `raw`：

```jinja
{% raw %}{{ this_is_not_rendered }}{% endraw %}
```

### 5.4 常用内置过滤器

| 过滤器 | 示例 | 结果 |
| --- | --- | --- |
| `upper` / `lower` | `{{ port.value \| upper }}` | `A` → `A` |
| `int` / `float` | `{{ baud.value \| int + 1 }}` | `115201` |
| `default` | `{{ mode.value \| default("NONE") }}` | 空值时兜底 |
| `replace` | `{{ port.value \| replace("A", "B") }}` | `B` |
| `join` | `{{ ["a","b"] \| join(",") }}` | `a,b` |
| `trim` | `{{ mode.value \| trim }}` | 去首尾空白 |
| `indent` | `{{ block \| indent(4) }}` | 整体缩进 |
| **`pvs` / `wrap`** | `{{ port.value \| pvs("GPIO", "_PORT") }}` | `GPIOA_PORT` |

---

## 6. Excel 表结构与填写规则

### 6.1 Global Parameter

|  | A | B | C | D | E |
| --- | --- | --- | --- | --- | --- |
| **1** | Variable | Value | Description | Prefix | Suffix |
| **2** | baud | `115200` | 波特率 |  |  |
| **3** | mcu | `STM32F103` | MCU 型号 |  |  |

- **B 列是唯一需要填写的列**（黄色底纹），留空即用默认值。
- D/E 列可覆盖 YAML 的 Prefix/Suffix。

### 6.2 Local Parameter

|  | A | B | C | D | E | F | G |
| --- | --- | --- | --- | --- | --- | --- | --- |
| **1** | Variable | Description | Prefix | Suffix | **Case1** | **Case2** | **Case3** |
| **2** | port | 端口 | GPIO | _PORT | `A` | `C` | `A` |
| **3** | mode | 模式 | MODE_ |  | `TX_RX` | `RX` | `TX_RX` |

- **E1 起每个非空表头就是一列 Case**，名字直接成为 `{{ case_name }}`。
- **右拉增加 Case**：选中 `E1:F3`（含表头的整块区域）→ 拖动右下角填充柄向右 → 把新表头改成 `Case3`，或写成 `UART1`、`BoardB` 等任意名字。
- Case 列必须**连续**：解析时从 E 列开始，遇到第一个空表头即停止。中间不要留空列。
- Case 表头不可重复。

### 6.3 Output 表

- 内容由 `render --write-excel` 写入，位置由模板的 `start_cell` + `direction` 决定。
- **写进去的是什么取决于 `engine`**：`snapshot` 写渲染出来的文本；`excel` 写 Excel 公式（§14）——
  后者改参数后由 Excel 自己重算，不需要重跑脚本。
- 每次渲染前会清理旧区域：清理范围按**表内原有的真实边界**（扫到最后一个有内容的行/列）算，
  不是按"本次渲染有多少行"算 —— 所以**改短模板、减少 Case 之后重渲染也不会残留**上一版的内容。
- 表头（Case 名）也会被写出：横向布局写在 `start_cell` 的**上一行**，纵向布局写在 `start_cell` 的**左侧一列**；
  若 `start_cell` 已在第 1 行 / 第 A 列，则不写表头。
- 同一张输出表放多个模板时请把 `start_cell` 错开：清理是按区域做的，重叠会互相覆盖。

### 6.4 Template 表（隐藏）

由 `init` 生成（`--no-template-sheet` 可关闭）。**首行是红字提示**，写明它只是只读参考：

```
A1: !! 只读参考：模板真源是 YAML / template_file；在这里改模板不会影响渲染结果
A2: ### uart_init
A3: output_sheet     B3: Output
A4: start_cell       B4: B2
...
A9: code ↓
A10: 1               B10: // Case: {{ case_name }}
A11: 2               B11: UART_Init({{ baud }}, {{ port }}, {{ mode }});
```

底部还有一段机器可读的元信息（`render --write-excel` 写、`check` 读）：

```
A20: ## excel-codegen-meta
A21: 时间          B21: 2026-09-24 15:23:26
A22: 输出指纹      B22: 80ecfe5dc6e7
A23: 参数指纹      B23: c1273da2acb5
A24: case 数       B24: 1
A25: cases         B25: Case1
```

它是**只读参考**，修改它不会影响渲染结果（真正的模板来源始终是 YAML / `template_file`）。

### 6.5 HOWTO 表（第一张）

`init` 生成、`render --write-excel` 刷新（`--no-howto` / `excel.howto_sheet: null` 可关闭）。
里面写的是"打开这个文件的人需要知道的事"：

- 三步怎么用（Global 填 B 列 / Local 从 E 列起填 Case / 改完回命令行跑哪条命令）；
- 那句命令本身，以及"只想导出文件换成 `--outdir`"、"想查过期用 `check`"；
- ⚠ 输出表是**快照**不是活公式，以及"不带 `--write-excel` 的 render 不会改这个文件"；
- 每个模板的输出位置（表 + 起始单元格 + 布局 + `case_filter`）；
- **本次生成**：时间 / 参数指纹 / 输出指纹 / case 列表 / 实际执行的命令。

> 这张表是给人看的，所以内容随每次写回重生成 —— 不要在里面写你自己的笔记。

---

## 7. 输出布局与 `start_cell`

### 7.1 horizontal（一个 Case 一列）

```yaml
output_sheet: "Output"
start_cell: "B2"
direction: "horizontal"
```

```
      A          B                     C
1                Case1                 Case2          <- Case 表头（start_cell 上一行）
2   （空闲）     // Case: Case1        // Case: Case2  <- 第 1 行结果
3                UART_Init(...);       UART_Init(...);<- 第 2 行结果
4                // 纯值: A ...        // 纯值: A ...
```

适合：把不同 Case 的代码并排比较。第 n 行的内容来自渲染结果的第 n 行——**模板里一行对 Excel 里一行**。

### 7.2 vertical（一个 Case 一行）

```yaml
output_sheet: "Output Vertical"
start_cell: "B2"
direction: "vertical"
```

```
      A        B       C       D
2     Case1    line1   line2   line3     <- 一行放完该 Case 的所有行
3     Case2    line1   line2   line3
```

适合：Markdown 表格、CSV、汇总信息。此时**模板里的换行会横向展开**，所以纵向布局的模板通常写单行或短块。

---

## 8. 导出代码文件（`--outdir`）

```bash
excel-codegen render --config examples/example.yaml --excel examples/template.xlsx --outdir generated/
```

- 文件名优先级：`template.filename`（Jinja2 渲染，可用 `{{ case_name }}`、`{{ template_name }}`）→ 否则 `<模板名>_<Case名><extension>`。
- 文件名中的 `<>:"/\|?*` 与控制字符会被替换为 `_`（Windows 兼容）。**注意这意味着文件名里不能带目录**：
  `+`、`.`、`-`、空格、中文都是安全的（工况名可以直接当文件名用），要分目录请用不同的 `--outdir`。
- 文件内容以 UTF-8 写入，并保证以换行符结尾。
- 目标文件已存在时默认覆盖；加 `--no-overwrite` 可在冲突时报错。

```yaml
templates:
  - name: uart_init
    filename: "uart_init_{{ case_name }}.c"      # → generated/uart_init_Case1.c
  - name: ext_code
    filename: "{{ case_name }}.js"               # → generated/EXT-T20.559-mu90-kf+1.js（+ 与 . 都安全）
  - name: plain
    extension: ".md"                              # 未写 filename → plain_Case1.md
```

---

## 9. 新增一个模板（扩展模板库 checklist）

1. **想清楚输入**：需要哪些参数？哪些全局、哪些逐 Case？名字用合法标识符。
2. **写变量定义**：加到 `variables.global` / `variables.local`，给出 `description` 与 `default`（默认值决定 `init` 预填内容）。
3. **选输出目标**：确定 `output_sheet`（必须在 `excel.sheets.outputs` 里声明，否则配置校验失败）、`start_cell`、`direction`。
4. **写模板**：
   - 组合值 `{{ x }}`、纯值 `{{ x.value }}`、过滤器 `{{ x.value | pvs("P","S") }}`；
   - 不要引用未定义的变量（`StrictUndefined` 会报错）；
   - 控制结构用 `{% ... %}`，注意空白控制（空行会各占一行）。
5. **需要"一本工作簿两套规则"时**加 `case_filter`（见 9.2），并确认 Global 表只填一次。
6. **校验**：`excel-codegen validate --config xxx.yaml`（只看配置）→ 加 `--excel` 校验填写内容。
7. **重新生成模板**：`excel-codegen init --config xxx.yaml --output template.xlsx --force`
   （已有填写内容时先备份，`init` 会重建工作簿）。
8. **渲染验证**：`excel-codegen render --config xxx.yaml --excel template.xlsx --show`（先预览再写回）。
9. **写回并确认不过期**：`render … --write-excel` 之后跑 `excel-codegen check --config xxx.yaml`（应当 exit 0）。
10. **补测试**：在 `tests/` 中按现有风格加一条端到端用例，保证模板库后续可回归。

### 9.1 外部模板文件示例

```yaml
templates:
  - name: uart_init
    output_sheet: "Output"
    template_file: "templates/uart_init.c.j2"   # 相对 YAML 所在目录
```

`templates/uart_init.c.j2`：

```jinja
/* {{ case_name }}: UART{{ port.value }} @ {{ baud }} */
void UART{{ port.value }}_Init(void) {
    uart_cfg.baud = {{ baud }};
    uart_cfg.mode = {{ mode }};
    uart_cfg.pin  = {{ port | upper }};
}
```

外部模板的好处：语法高亮、可与 IDE 插件配合、便于复用；`Template` 隐藏表仍会保存其内容快照（`init` 时）。

### 9.2 `case_filter`：一个模板只作用于部分 Case

默认 `render` 是"模板 × Case"的交叉积：每个模板都会给每个 Case 出一份结果。
加一行 `case_filter`（一段 Jinja2 **表达式**，对每个 Case 的上下文求值）就能把模板限制在匹配的 Case 上。

典型场景：**一个项目两套/多套规则，共用同一张 Global 表**（例如 ABS FPI 内外压共用船的主尺度）。

```yaml
excel:
  sheets:
    global: "Global Parameter"          # 主尺度只填一次
    local: "Local Parameter"
    outputs: ["Code EXT", "Code INT"]   # 两套模板各自的输出表

variables:
  global:
    - name: L
      default: 340
  local:
    - name: kind                        # 这个 Case 属于哪一套规则
      type: string
      default: "EXT"
    - name: draft
      default: 20.5

templates:
  - name: ext_code
    output_sheet: "Code EXT"
    case_filter: "kind == 'EXT'"
    code: |
      // EXT {{ case_name }} (L={{ L }})
  - name: int_code
    output_sheet: "Code INT"
    case_filter: "kind == 'INT'"
    code: |
      // INT {{ case_name }} (L={{ L }})
```

语义与注意事项：

| 主题 | 说明 |
| --- | --- |
| 变量当字符串用 | 等于**组合值**（和 `{{ x }}` 一致），所以 `kind == 'EXT'`、`port == 'GPIOA_PORT'` 都能写 |
| 纯值 / 数值比较 | 用 `.value`：`draft.value > 20`、`kind.value == 'EXT'`（直接拿字符串和数字比会报错，工具会提示） |
| 被跳过的 Case | 不渲染、不占输出列/行；`render` 摘要里会写"跳过 N 个" |
| 与 `--case` 的关系 | `--case` 是**全局**过滤（先筛 Case 列表），`case_filter` 是**per-template** 过滤，两者可叠加 |
| 报错方式 | 语法错误、引用了取不到的变量、错误的类型比较 → `RenderError`，不会静默变成"什么都没生成" |
| 新增 Case 列 | **先把归属变量（`kind` 之类）填上**：留空会回落 YAML `default`，这一列会**悄悄进了某个规则集**；`render` / `check` 会对"整列都空的 Case"给出提醒（`explicit_values`） |
| `validate` | 会把 `case_filter` 引用的变量算进"被引用"，并检查表达式语法 |

> 代价提示：`case_filter` 解决的是"模板 × Case"的过度生成，**不解决**"多一层作用域"的问题（见 §13）。

---

## 10. 排错手册

| 现象 / 报错 | 原因 | 处理 |
| --- | --- | --- |
| `YAML 解析失败 ...（第 N 行，第 M 列）` | 缩进/引号/块标量写错 | 按提示行号检查；中文建议加引号 |
| `YAML 键重复: 'x'（第 N 行）` | 同一个映射里写了两次同一个键（PyYAML 默认会静默覆盖） | 按行号合并成一份（常见于写了两块 `variables:`） |
| `配置校验失败 ... - templates.0.start_cell: 非法单元格引用` | `start_cell` 写成了 `2B`、`B0` 之类 | 改为 `B2` 形式 |
| `... output_sheet 'X' 未在 excel.sheets.outputs 中声明` | 模板输出表没声明 | 在 `excel.sheets.outputs` 里加上 |
| `Excel 缺少工作表: 'Local Parameter'` | Excel 与 YAML 的表名不一致，或用了旧模板 | 名字对齐，或 `init --force` 重新生成 |
| `工作表 'Local Parameter' 从 E1 开始没有 Case 列` | Case 表头被删/被改到其他列 | 在 E1 起恢复 `Case1`… |
| `模板 'x' 变量缺失: 'y' is undefined` | 模板引用了上下文没有的变量 | 在 YAML 里加变量，或在 Excel 表中加同名变量行 |
| `模板 'x' 语法错误：第 N 行 ...` | Jinja2 标签未闭合、`{% for %}` 缺目标 | 按行号修正；成对检查 `if/endif`、`for/endfor` |
| `模板 'x' 的 case_filter 求值失败: ...` | 表达式里做了非法比较（如 `draft > 20`）、或变量取不到 | 数值比较写 `.value`：`draft.value > 20` |
| `ERROR 输出表已过期，共 N 处不一致` | 改了参数但没跑 `render --write-excel` | 跑一次写回；或按提示的行/列核对 |
| `变量名 'a-b' 不是合法标识符` | 变量名含 `-`、空格、中文 | 改名或用下划线 |
| `无法写回 Excel ... 文件被 Excel 占用？` | 文件正在 Excel 中打开 | 关闭 Excel 后重试 |
| 输出里出现 `None` | 变量既没有值也没有 `default` | 在 YAML 给 `default: ""` |
| 数值变成 `115200.0` | 该变量 `type: float` 或值是浮点 | 用 `type: int`；整数浮点默认已规范化 |

---

## 11. FAQ

**Q1：为什么 `{{ port }}` 带前缀，我只想要 `A`？**
前缀/后缀是变量定义的一部分。三种做法：① 写 `{{ port.value }}`；② 把该变量的 `prefix`/`suffix` 留空；
③ 用过滤器显式控制：`{{ port.value | pvs("", "") }}`。

**Q2：不同 Case 想用不同前缀怎么办？**
Prefix/Suffix 是按变量（不是按 Case）生效的。见 4.4：再加一个局部变量承载前缀，或把完整名字直接填进单元格。

**Q3：`init --force` 会丢掉我填好的参数吗？**
会。`init` 是"重建模板"。建议：填写前先 `git`/复制备份；或只用手工方式给 Excel 增加变量行（解析器允许 YAML 之外的变量行）。

**Q4：能生成多个文件 / 多种语言吗？**
可以。为每种目标加一个 `templates` 条目，配置不同的 `output_sheet`（或同一张表用不同的 `start_cell`）、`filename`、`extension`。一次 `render --outdir` 全部产出。

**Q5：一个 Output 表能放多个模板吗？**
可以，但请让它们的 `start_cell` 错开（写入前会清理区域，重叠会互相覆盖）。稳妥做法是一模板一表。

**Q6：Case 名可以是中文或带空格吗？**
可以（表头任意非空字符串，如 `主板A`）。`{{ case_name }}` 会原样输出；导出文件名中的非法字符会被替换为 `_`。

**Q7：变量值里需要包含空格或特殊字符？**
直接在单元格里写即可，渲染会原样输出（不转义）。注意这通常是生成代码的合法性问题，而不是本工具的问题。

**Q8：怎么在模板里引用变量本身的名字？**
把 `template_name` 用于文件名，`case_name` 用于正文。变量名目前不注入上下文（避免与用户变量冲突）。

**Q9：渲染顺序是怎样的？**
外层遍历 `templates`，内层遍历 Case（按 Excel 中从左到右的列顺序）；`--case` 可以只渲染指定 Case 并保持你传入的顺序。

**Q10：能否在 CI 中校验模板库？**
可以，而且分两层：

1. `excel-codegen validate --config xxx.yaml --excel xxx.xlsx`：配置 / 模板语法 / 变量引用 / Excel 结构；
2. **`excel-codegen check --config xxx.yaml --excel xxx.xlsx`**：表里的代码与当前参数是否一致（过期 → exit 1）。
   `check` 用"重新渲染 + 逐行比对 + 指纹比对"，所以它同时能抓到"改了参数没重跑"和"被手工改过的输出表"。

**Q11：怎么知道一份工作簿是什么时候、用哪套参数生成的？**
看 `HOWTO` 表的"本次生成"（时间 / 参数指纹 / 输出指纹 / case 列表 / 命令），或跑 `validate -x` / `check`。
指纹是**参数**的哈希：只要 Global / Local 的取值、前缀、后缀中任何一项变了，指纹就会变。
（`template_sheet` 被关掉时，这些信息仍在 `HOWTO` 表里。）

**Q12：`render` 不带 `--write-excel` 会不会悄悄改我的文件？**
不会，而且现在也不会再"骗你"：摘要里永远有 `写回 Excel │ 是 / 否（需要 --write-excel）` 这一行，
没写回时还会多一条 `!` 提示。默认行为是**只预览**。

**Q13：`--cases` 只能给数量吗？**
现在也可以给名字：`--cases EXT-T20.559,INT-T15`（逗号分隔），这样工况名一次到位，不用 init 之后再改表头。

**Q14：能不能不用 `HOWTO` / `Template` 表？**
可以：`excel.howto_sheet: null`、`excel.template_sheet: null`，或 CLI 上 `--no-howto` / `--no-template-sheet`。

---

## 12. 版本与变更速览

| 版本 | 要点 |
| --- | --- |
| 0.1.0 | 首个可用版本：YAML 定义 + Excel 填写 + Jinja2 渲染 + 两种布局 + 导出 |
| **0.2.0** | 真实项目移植后的修复会话：Prefix/Suffix 保留空格、`render` 不再静默、GBK 控制台不崩、清理旧结果按真实底边、`HOWTO` 表、指纹 + `check` 命令、`case_filter`、`{{ x.value }}` 形态一致、YAML 重复键报错、`--cases` 支持名字、文档补齐能力边界 |
| **0.3.0** | **公式模式（`engine: excel`）**：输出表写 Excel 公式，改参数由 Excel 自己重算，不用再跑脚本；配 `examples/example_formula.yaml`。见 §14 |
| **0.4.0** | 公式求值器（`check` 默认把公式算一遍再与 Python 渲染比对）、公式模式差异信息改为贴模板行、空 Case 列 / 超长公式告警、§14.5「改结构什么要重跑」、**更正 0.3.0 文档里"有 if/else 就不能用公式模式"的错误判断** |
| **0.5.0** | **派生参数（`derived:`）**：参数引用参数（同 Case 的 local + global），Excel 侧写成公式自动重算；数值形态改成"两边都按 15 位有效数字"（不再用 `TEXT()`）。见 §15 |
| 0.5.1 | 只动环境与文档（Windows → Ubuntu 迁移）：新增 `setup.sh`（探测文件系统后选 venv 位置）与 `.gitattributes`（行尾统一），README「安装」补 Ubuntu 三个坑，`compare_with_rules.js` 缺外部依赖时明确 `SKIP`（退出码 2）。**工具行为无变化** |
| 0.5.2 | **多平台零配置**：新增跨平台锁文件 `uv.lock` 与 `.python-version`，`pyproject.toml` 加 `[dependency-groups] dev`（PEP 735）—— 装上 uv 之后 `uv run pytest` / `uv run excel-codegen …` 在 Windows / macOS / Linux 完全一致，不需要 venv、pip、apt。`setup.sh` 同步支持 uv 路径。**支持下限 3.10 → 3.11**（3.10 于 2026-10 结束支持），CI 改为测「底线 3.11 + 3.13」。**工具行为无变化** |

---

## 13. 能力边界与不适用场景

工具很小，边界要写清楚，免得把"表达不了"误当成"工具坏了"：

| 做不到 | 说明 / 应对 |
| --- | --- |
| **第三层作用域**（船 → 工况 → 舱/设备） | 只有 `global` / `local` 两层。一个舱被多个工况引用时，舱的参数只能**按工况摊平**（同名舱在每个 Case 列里各写一遍），改一个舱的尺寸要改 N 列。应对：把"每 Case 重复"的参数集中放在同一个变量块，或用 §9.2 的 `case_filter` 把不同规则分列；真正的多层引用需要下游脚本（例如本项目 `abs_fpi` 里用 openpyxl 预填）。 |
| 快照模式下的自动重算 | `engine: snapshot`（默认）的输出表是**文本快照**，改了参数必须重跑 `render --write-excel`；用 `check` 判过期。想要"改参数自动重算"就把模板改成 `engine: excel`（§14）。 |
| 公式模式下的**跨行**控制流 | 公式只做占位符替换 + **行内** `{% if %}`（§14.1.1）。`{% for %}`、跨行的 `{% if %}`、过滤器一律报错（带行号）—— 一行模板对应一个单元格，行数一变就没法映射。复杂模板留在快照模式。 |
| 语义检查（量纲、语法、业务规则） | 本工具只管"参数怎么填、模板怎么复用、结果怎么写回"，它不认识 GeniE/C 的语法，也不会拦 `Math.sin(90)`。这类检查留给下游校验器。 |
| 条件包含 / 片段组合 | 没有 `{% include %}` 的模板搜索路径，也没有"部分模板"；跨文件复用请用 YAML 锚点或把公共片段写成变量值。**注意区分**：跨 **YAML 文件**复用变量与模板有 `extends`（§16），这里说的是**模板片段**的复用 |
| 单个变量跨 Case 的联动 | 每个 Case 的取值互相独立；"B 列跟着 A 列算"这种联动要在 Excel 里用公式实现（读的是公式结果）。 |
| 超大工作簿的性能 | 读取走 openpyxl 全量加载；万行级别以内没问题，更大规模建议按规范集/分段拆工作簿。 |

> 判断标准很简单：**如果一件事需要"引用别的表/别的层"，本工具表达不了；
> 如果只是"同一套模板 × 多个工况取值"，它就能做。**

---

## 14. 公式模式（`engine: excel`）：改参数免重跑脚本

默认的 `engine: snapshot` 把渲染结果写成**文本快照** —— 改了参数必须重跑 `render --write-excel`。
给模板加一行 `engine: excel`，输出表里写的就是**Excel 公式**：改了 Global / Local 的参数，
Excel / WPS 打开即重算，**不需要任何脚本**。

```yaml
templates:
  - name: uart_init
    output_sheet: "Output"
    start_cell: "B2"
    direction: "horizontal"
    engine: "excel"                 # ← 只要这一行
    code: |
      UART_Init({{ baud }}, {{ port }}, {{ mode }});
      // 纯值: {{ port.value }}, 前缀: {{ port.prefix }}, 后缀: {{ port.suffix }}
```

`render --write-excel` 之后，`Output!B2` 里是（简化显示）：

```excel
="UART_Init("&INDEX('Global Parameter'!$D:$D,MATCH("baud",'Global Parameter'!$A:$A,0))
 &TEXT(IF(ISBLANK(INDEX('Global Parameter'!$B:$B,MATCH("baud",'Global Parameter'!$A:$A,0))),
   115200,INDEX(...)),"0")
 &", "&'Local Parameter'!$C$3&…&");"
```

### 14.1 支持范围

| 支持 | 说明 |
| --- | --- |
| `{{ x }}` | `Prefix & Value & Suffix`（与快照模式的组合值一致） |
| `{{ x.value }}` / `{{ x.text }}` | 纯值 |
| `{{ x.prefix }}` / `{{ x.suffix }}` | 前缀 / 后缀 |
| `{{ case_name }}` | 当前 Case 名（取 Local 表表头；横向布局向右拖会跟着换 Case） |
| `{{ template_name }}` | 模板名（编译成常量） |
| **行内 `{% if %}`** | `{% if 条件 %}A{% else %}B{% endif %}` → Excel `IF(...)`，见 14.1.1 |
| `{# 注释 #}` | 整段丢掉（不进入公式） |

**不支持**（会直接报错并给出行内容，让你决定是否改回快照模式）：

```text
{% for %} / {% set %} / {% include %}  → 循环与跨行分支没法用单元格引用表达
{{ port | upper }}                     → 过滤器
{{ baud + 1 }}                         → 表达式/运算
{{ port.name }}                        → 只支持 .value/.text/.prefix/.suffix
{{ 未在 YAML 中定义的变量 }}            → 公式需要知道它是 global 还是 local
```

### 14.1.1 行内 `{% if %}`

分支会让"一个 Case 占几行"变得不固定，而公式模式是**一行模板 → 一个单元格**，
所以有一条硬约束：

> **`{% if %}` 必须整段写在同一行内。**

```jinja
y = {% if flag.value == 1 %}{{ L }}{% else %}- {{ L }}{% endif %};
```

编译成 `IF(<条件>,<then>,<else>)`，分支里的 `{{ }}` 照常展开。可以嵌套。

**条件里必须写属性**（`.value` / `.text` / `.prefix` / `.suffix`），不能裸写变量名：

```jinja
{% if draft.value > 20 %}…{% endif %}          ✅
{% if kind.value == "EXT" %}…{% endif %}       ✅
{% if not flag.value %}…{% endif %}            ✅
{% if draft > 20 %}…{% endif %}                ❌ 报错：请写 draft.value
```

这不是洁癖。快照模式里裸变量是一个 `VarValue` 对象，`VarValue > 20` 会直接抛
`TypeError`；而 `draft.value` 在两种引擎里都是"纯值"—— **只有这种写法能让两边逐字一致**。
规则与 `case_filter` 相同（那里也是"要数值比较请用 `.value`"）。

条件支持：

| 类别 | 写法 |
| --- | --- |
| 比较 | `==` `!=` `<` `>` `<=` `>=` |
| 逻辑 | `and` / `or` / `not` |
| 真假判断 | `{% if flag.value %}` —— 数值比 `0`，文本比空串（与 `bool(value)` 的直觉一致） |
| 特殊名字 | `case_name` / `template_name` 可以裸写（它们没有 `.value`） |

`check` 会把编译出来的 `IF` 公式**算一遍**再与快照渲染比对（§14.3），
所以"两种引擎是否给出同一段代码"是有证据的，不用你自己逐个工况核对。

### 14.2 一致性保证（与 Python 侧对齐）

| 主题 | 做法 |
| --- | --- |
| 变量定位 | `INDEX/MATCH` 按**变量名**查，所以在 Global / Local 表里插行/删行都不会指错 |
| 空单元格 | 用 `ISBLANK()` 判断后回落 YAML `default`（**不能用 `=""`**：`INDEX` 对空单元格返回数值 `0`，会把空前缀变成 `0`） |
| 数值形态 | **不用 `TEXT()`**！`TEXT(20.559,"0.###############")` 会打出 `20.559000000000001`（二进制尾巴）。数值→文本走 Excel 的**隐式转换（General = 最多 15 位有效数字）**，与 Python 侧 `utils.to_text`（`%.15g`）逐字一致；整数浮点（`340.0`）两边都输出 `340` |
| 横向布局 | 局部变量的值列写成**相对列**（`E:E`），把输出格向右拖就跟着换 Case |
| 纵向布局 | 值列与 Case 表头写成**绝对列**（`$E:$E`、`$E$1`），因为纵向向右拖是"同一 Case 的下一行" |
| 重算 | 写公式时同时设置 `fullCalcOnLoad`，Excel / WPS 打开就重算，不显示旧缓存 |

### 14.3 三条必须知道的代价

1. **值只活在 Excel 里**。`openpyxl` 只写公式、不算公式，所以：
   - 导出代码文件（`--outdir`）仍由 Python 渲染（内容与公式等价，因为模板是纯替换）；
   - `check` 默认会把 Output 表里的公式**在 Python 里算一遍**（`formula_eval.py` 的求值器），
     与 Python 渲染逐行比对 —— 所以"公式算出来的值对不对"是有证据的，不用另写脚本。
     参数单元格本身是公式（用户手写 `=L/10`）时离线读不到值，会降级成"只比公式"并给出提示
     （`--no-values` 可显式关掉）。
2. **改模板要重跑一次**（`render --write-excel` 刷新公式），改参数不用 —— YAML 仍是唯一真源。
   这也是选择"YAML 当真源"的代价：公式写进 xlsx 之后没法像 Jinja 那样 diff / 评审。
   **改结构**要不要重跑，见 §14.5。
3. **两个已知差异**：`type: bool` 在 Excel 里是 `TRUE`/`FALSE`（Python 侧是 `true`/`false`）；
   `TEXT()` 的格式串受区域设置影响（中文/英文区域小数点是 `.`，欧洲区域是 `,`）。
   因此 **bool 变量建议用 `type: string`**，或该模板继续用快照模式。
4. **公式有长度成本**：一个 `{{ x }}` 大约展开 **300–400 字符**（`INDEX`+`MATCH`+`ISBLANK`+`TEXT`）。
   经验值：**一行 8 个以上占位符就该考虑拆行**（约 3000 字符，`validate` / `render` 会给出警告）；
   硬上限 8192 字符/格，超过直接报错。

### 14.5 改结构：什么要重跑、什么不用

| 动作 | 要不要重跑 `--write-excel` | 为什么 |
| --- | --- | --- |
| 改 **Global / Local 的参数值** | ❌ 不用 | 公式引用的是单元格，Excel 打开即重算 |
| 插 / 删 **变量行** | ❌ 不用 | 公式用 `INDEX/MATCH` **按变量名**定位 |
| 改 **变量行的前缀 / 后缀**（C/D 列） | ❌ 不用 | 同上，前缀后缀也是按名字查列 |
| **增 / 删 / 移动 Case 列** | ✅ **要** | 横向布局里公式用的是**相对列标**（`E:E`），列一移位就指着别的工况；纵向布局同理 |
| 改 **YAML**（模板 / 变量定义 / 引擎） | ✅ **要** | 公式是写进格子的文本，得刷一遍 |
| 改 `case_filter` / 新增 Case 表头 | ✅ **要** | 归属与列顺序变了 |

增删 Case 列这件事**工具能发现**：`check` 会报"公式算出来的文本与 Python 渲染不一致 →
（参数表结构改动过？插/删过 Case 列？请重跑 --write-excel）"，放进 CI 就能拦住。

> ⚠ 顺带一个坑：**手加 Case 列时先把归属变量填上**（`case_filter` 用的那个，例如 `kind`）。
> 留空会回落 YAML `default`，这一列就**悄悄进了某个规则集**；`render` / `check` 会对
> "整列都空的 Case"给出提醒。

### 14.6 什么时候不要用公式模式

* 模板里有**跨行**的 Jinja 控制流（`{% for %}`、跨行的 `{% if %}` / `{% set %}`）或过滤器。
  注意区分：**生成目标语言自己的 `if/else` 只是字面文本**（GeniE / C / Python 的语句），
  **不影响**公式模式。判断标准是"有没有 `{%`"，不是"代码里有没有 if"。
  本仓库 `abs_fpi` 的内外压模板就是这样：C_1 区间、k_lo、Girth 插值那些 `if/else` 是字面文本，
  真正的 Jinja 只有开头几行坐标映射，把那几行改写成纯替换之后，整个模板就能进公式模式了。
  **行内的 `{% if %}` 是支持的**（§14.1.1）—— 只有跨行的才不行。
* 需要用 `--outdir` 产出文件后又希望"改参数自动同步文件"—— 文件必须由 Python 再生成一次
  （注意：`render --write-excel` 之后 `check` 已经能证明"两边一致"，所以"文件与 Excel 不同步"只发生在
  你没重跑导出的时候）。
* 需要模板可评审 / 可回滚（公式的可读性远不如 Jinja）。
* 大多数行的占位符超过 8 个（公式会长到没人看得懂，见 §14.3 第 4 条）。

一条经验：**参数多、模板是纯替换** → 公式模式；**模板带 Jinja 控制流** → 快照模式。
同一个工作簿里可以两种混用（`engine` 是 per-template 的）。

---

## 15. 派生参数（`derived:`）：参数引用参数

有些量不是"填"出来的，而是从别的参数算出来的：中间量（`h_de = k_c * h_di`）、
需要**在表里看得到数值**的换算结果（`rho_g = rho * g`）、先算好再被模板引用的量。
给变量写一行 `derived:`，它就成了**派生参数**：不用填值，工具按表达式算，模板里照常 `{{ 名字 }}`。

```yaml
variables:
  global:
    - name: rho
      default: 1025
      type: float
    - name: g
      default: 9.81
      type: float
    - name: rho_g
      description: "rho * g"
      derived: "rho * g"          # ← 引用另外两个参数
      type: float
      suffix: " N/m^3"
  local:
    - name: draft
      default: 20.5
      type: float
    - name: z
      default: 5
      type: float
    - name: h_s
      description: "静水压头 = max(draft - z, 0)"
      derived: "max(draft - z, 0)"     # ← 引用同 Case 的 local
      type: float
      suffix: " m"
    - name: p_s
      description: "静水压力（链式：引用了派生参数 rho_g）"
      derived: "rho_g / 1000 * h_s"    # ← 派生参数可以引用派生参数
      type: float

templates:
  - name: pressure
    engine: "excel"
    code: |
      var h_s = {{ h_s }};
      var p_s = {{ p_s }} Pa;
```

### 15.1 语义

| 主题 | 说明 |
| --- | --- |
| 能引用什么 | ``global`` 与**同一个 Case 的** ``local``。**不跨 Case**、不跨工作簿 |
| ~~global 引用 local~~ | ❌ 报错：global 的派生参数没有"当前 Case"这个概念 |
| 引用拿到的是 | **纯值**（`x.value` 语义），不是带前后缀的组合值 —— 所以 `rho * g` 能直接算 |
| 前后缀 | 派生参数自己的 `prefix` / `suffix` 照常生效（`{{ rho_g }}` 带单位） |
| 链式 | 可以引用别的派生参数，工具按依赖顺序求值；**成环会报错**并给出环上的名字 |
| 名字范围 | 只能引用 YAML 里定义过的名字（表里临时加的变量不能作为派生引用） |
| 类型 | 按 `type` 转换结果（`int` / `float` / `string` / `bool` / `auto`） |

### 15.2 参数表里长什么样

派生格是**淡绿色**、描述里带「（自动计算：表达式）」，而且——**能写成 Excel 公式的就写成公式**：

```excel
=MAX((IF(ISBLANK(INDEX('Local Parameter'!$E:$E,MATCH("draft",…))),20.5,…)
      -IF(ISBLANK(INDEX('Local Parameter'!$E:$E,MATCH("z",…))),5,…)),0)
```

所以**改输入就自动重算**（和 `engine: excel` 一个道理），不用重跑脚本。
派生格里手工填的东西**会被忽略**（`render` / `check` 会给告警，`check` 还会因为
"Excel 里算出来的代码与导出一致性被破坏"而报错）。

### 15.3 能翻译成 Excel 公式的子集

| 可以 | 说明 |
| --- | --- |
| `+ - * / **` | `**` → `^` |
| `%` | → `MOD(a,b)` |
| `~` | 字符串拼接 → `&`（**不要用 `+` 拼字符串**，Excel 的 `+` 不做文本拼接） |
| 括号、数字、字符串、`TRUE/FALSE` | |
| `min / max / abs / int / float` | `int` → `TRUNC`（Excel 的 `INT` 是向下取整，负数会差 1） |
| `\|abs` / `\|int` / `\|float` / `\|string` | |
| 比较 `== != > < >= <=` | → `= <> > < >= <=` |
| `and / or / not` | **只能出现在条件里**（Python 的返回值语义与 Excel 不同，放别处不翻译） |
| `a if cond else b` | → `IF(cond,a,b)` |

| 不翻译（会降级成"写入算好的值"并告警） | 原因 |
| --- | --- |
| `round()` / `\|round` | Python 是**银行家舍入**（`round(2.5)==2`），Excel `ROUND` 是四舍五入（2.5→3）。翻过去就会出现"Excel 里看到的"与"导出文件"不一致 |
| `ceil / floor` | 边界行为两边不同 |
| `//` | Excel 没有等价的整除 |
| 其它函数 / 过滤器 / 属性访问 | 没有对应物 |

**降级**不是报错：格子写 Python 算好的值，`validate` / `render` 会告警
"这些派生参数写不成 Excel 公式…改输入后要重跑 --write-excel"。

### 15.4 三条实践建议

1. **`validate` 会提前编译一遍**：语法错、引用越界、循环、翻译不了，都在配置阶段报出来；
2. **`check` 会验证值**：它把 Output 表的公式（会递归算到派生格）算一遍，与 Python 渲染比对 ——
   所以"Excel 里看到的"和"导出的"是不是同一段代码，是有证据的（§14.3）；
3. **别把派生参数当"计算器"**：它适合一步到位的换算（单位、系数、中间量）。
   需要分支/迭代/复杂逻辑的，写在模板的 Jinja 里更清楚（或按工况填进去）。

> 什么时候**不要**用：需要 `round` 的精确语义、要用 `{% if %}` 表达的分段逻辑、
> 或者这个量根本不该给 Excel 读者看到 —— 那就留在模板里。

---

## 16. 跨文件复用：`extends`

一个 YAML = 一本工作簿。但一个项目常常要在一本工作簿里放几套规范、共用一张 Global 表
（船的主尺度只填一次）。`extends` 就是为此：**根配置把几个规则集文件合并成一份**，
不用手抄变量。

```yaml
# project.yaml —— 工作簿的真源
version: 1
extends:
  - rules/external.yaml      # 路径相对**本文件**解析
  - rules/internal.yaml

excel:
  output: "ABS_FPI_load_cases.xlsx"
  sheets:
    global: "Global Parameter"
    local: "Local Parameter"
    outputs: ["Code EXT", "Code INT"]

variables:
  global:
    - name: project_code     # 根文件自己的变量，追加在后面
      default: "P-001"
  local: []
templates:
  - name: index              # 根文件自己的模板
    output_sheet: "Code EXT"
    code: |
      // {{ project_code }}
```

被 extends 的文件就是普通的单文件配置（单独拿出来 `validate` 也必须是对的）。

### 16.1 合并规则

顺序：**按 `extends` 列表的顺序先合并各文件，最后合并根文件自己的内容**（同名时先出现的赢）。

| 对象 | 判据 |
| --- | --- |
| 同名变量 | **必须逐字一致**的字段：`prefix`、`suffix`、`type`、`derived`、`min`/`max`/`choices`/`pattern`（它们决定生成出来的文本）—— 不一致直接报错 |
| 同名变量 | **只告警**的字段：`default`、`description`（不同规则集的示例工况本来就不同），保留先出现的那个 |
| 同名模板 | 内容必须**完全一致**（一致则去重），否则报错 |
| `excel` | 只有根文件的生效；被 extends 的文件里写了 `excel` 会告警并忽略 |
| `version` | 取根文件的 |

`extends` 是**递归**的（A extends B，B extends C 没问题），循环引用会报错并打出引用链。
合并过程中的告警会由 CLI 打出来（`! 变量 'draft'（global）的 default 在两处不一致…`）。

### 16.2 `template_file` 相对谁解析

**相对声明它的那个文件**。所以 `rules/external.yaml` 里写 `template_file: tpl/body.j2`，
指的是 `rules/tpl/body.j2` —— 模板文件可以和它所属的规则集放在一起，根配置放在别处也不会错。

这条是 `extends` 相对"手工合并"最容易出错的地方，工具把它做成了自动的。

### 16.3 与 `abs_fpi/compose.py` 的关系

`abs_fpi/compose.py` 是这套合并逻辑的**前身**（当时工具还没这个能力）：它把规则集 YAML 合成
一份项目 YAML 写进磁盘。现在有了 `extends`，大多数场景可以直接用：

| | `compose.py`（项目侧脚本） | `extends`（工具内建） |
| --- | --- | --- |
| 产物 | 生成一份新的 YAML 文件（需要跟着一起提交/维护） | 内存里合并，根 YAML 就是唯一真源 |
| `output_sheet` | 脚本按规则集后缀**改写**表名 | 各规则集自己写，或用 `case_filter` 分流 |
| `template_file` | 各文件必须在同一目录 | 相对各自文件解析 |

`compose.py` 仍有它的用处：它会在合并时**改写输出表名**（`Code` → `Code EXT`），
适合"一个规则集不关心自己叫哪张表"的批量场景。两者判据一致，可以并存。


