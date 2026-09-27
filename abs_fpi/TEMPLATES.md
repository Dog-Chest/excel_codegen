# 模板数据库

一个**规则集** = 一个 YAML（变量定义 + 模板声明）+ 它引用的 `.j2` 模板。
本目录是 ABS FPI 的两个规则集，同时演示"多规则集合进一本项目工作簿"的做法。

生成物是**一本 Excel**：工程师在表里填参数，`Code` 表里的代码由 **Excel 公式**算出来 ——
改参数后 Excel / WPS 打开即重算，不需要跑任何脚本。

---

## 1. 库存

| 规则集 | 规范 | 章节 | 变量（全局/局部） | 模板 | 产物 |
|---|---|---|---|---|---|
| `abs_fpi_external.yaml` | ABS FPI 2025 | 5A-3-2/5.5 外压 | 14 / 10 | `genie_ext`、`ext_summary` | `abs_fpi_external.xlsx`（3 个样例工况） |
| `abs_fpi_internal.yaml` | ABS FPI 2025 | 5A-3-2/5.7 内压 | 15 / 30 | `genie_int`、`int_summary` | `abs_fpi_internal.xlsx`（4 个样例工况） |
| `abs_fpi.yaml` ⟵ 生成物 | 上面两个合并 | — | 17 / 35 | 4 个（带 `case_filter`） | `ABS_FPI_load_cases.xlsx`（7 个样例工况，共用一张 Global） |

```
        abs_fpi_external.yaml ─┐
                               ├─ compose.py ─► abs_fpi.yaml ─► ABS_FPI_load_cases.xlsx
        abs_fpi_internal.yaml ─┘   (+compose.yaml)      ▲
                                                        └── 生成物，不要手改
```

三个层次，各有唯一职责：

| 层 | 文件 | 是什么 | 谁改 |
|---|---|---|---|
| **真源** | `abs_fpi_*.yaml`、`templates/*.j2` | 变量定义、模板正文 | 你（进 Git、可评审） |
| **配方** | `compose.yaml` | 哪几个规则集、各自 `case_filter` | 你（加规范时加一行） |
| **产物** | `abs_fpi.yaml`、`*.xlsx`、`generated/` | 项目 YAML / 工作簿 / 导出的代码 | `python build.py` |

> **为什么要有 compose.py**：工具只认"一个 YAML = 一本工作簿"，没有 `extends` / `!include`。
> 项目要在一本工作簿里放几套规范、共用一张 Global（主尺度只填一次），
> 手抄变量就是漂移的温床。`compose.py` 把"手抄"换成"合成"，并在合成时检查：
> 同名变量的 `prefix` / `suffix` / `type` **必须逐字一致**（不一致直接报错），
> `default` / `description` 不一致只告警。见 `FINDINGS.md` 的「0.3.0 复测」。

---

## 2. 一条命令

```bash
cd excel_codegen/abs_fpi
python build.py            # compose → 建骨架（缺哪个建哪个）→ 渲染写回 → 导出代码 → 两层验证
python build.py --check    # 只验证，不动文件（CI 可用）
python build.py --init     # 骨架不存在时也重建（会清掉你填的参数）
```

`build.py` 的实测输出：

```
== compose  规则集 → 项目工作簿 YAML
wrote abs_fpi.yaml: 17 全局 / 35 局部变量, 4 模板 → Code EXT, Summary EXT, Code INT, Summary INT

== 写入     渲染 → 写回 Excel → 导出代码
[ok]    abs_fpi_external.xlsx: 3 工况 × 2 模板 → 6 个导出文件
[ok]    abs_fpi_internal.xlsx: 4 工况 × 2 模板 → 8 个导出文件
[ok]    ABS_FPI_load_cases.xlsx: 7 工况 × 4 模板 → 14 个导出文件，case_filter 跳过 14

== 验证     公式求值 ↔ Python 渲染
  [公式] abs_fpi_external.xlsx        6 项, 0 项失败
  [公式] abs_fpi_internal.xlsx        8 项, 0 项失败
  [公式] ABS_FPI_load_cases.xlsx     14 项, 0 项失败

== 验证     excel-codegen check
  [check] abs_fpi_external.xlsx      OK
  [check] abs_fpi_internal.xlsx      OK
  [check] ABS_FPI_load_cases.xlsx    OK

公式求值 28 项 / 0 项失败；check 全过
```

---

## 3. 加下一个规范（DNV / BV / CSR…）的清单

1. **抄一份规则集**：`cp abs_fpi_external.yaml dnv_rp_c201.yaml`，改 `excel.output`
   （例如 `dnv_rp_c201.xlsx`）与 `templates` 列表。
2. **写变量**：`variables.global` 放全船数据，`variables.local` 放逐工况数据。
   - 名字必须是 `[A-Za-z_][A-Za-z0-9_]*`，不能是 `case_name` / `template_name`；
   - **单位放 `suffix`**（`suffix: " m"` → `340 m`）；无量纲留空；
   - `type: float` / `int` 让公式侧用 `TEXT()` 规范化数值；
   - 纯标注用的名字（如 `tank_ref`）用 `type: string`。
3. **写模板**：放 `templates/<名字>.js.j2`，一行一条语句。
4. **决定引擎**：
   - 模板**只有字面文本 + `{{ 变量 }}`** → 可以 `engine: excel`（改参数免重跑）；
   - 需要 Jinja 的 `{% if %}` / `{% for %}` / 过滤器 → 只能 `engine: snapshot`。
   - ⚠ **GeniE 代码里自己的 `if / else / { }` 只是字面文本，不影响公式模式**；
     真正拦住公式模式的是 *Jinja* 的控制流。本目录的坐标映射就是为此改写成
     6 个变量的纯替换（见 §4）。
5. **要不要进项目工作簿**：进的话在 `compose.yaml` 的 `rulesets` 加一行，
   给 `set:`（输出表后缀用）与 `filter:`（`kind == 'XXX'`）；
   同时确认这个规则集 YAML 里也声明了 `kind` 这个局部变量（`compose` 会校验一致性）。
6. **跑**：`python build.py`（首次 `--init`）。
7. **看验证**：`build.py` 会打两层结果；不一致会指出**第几行**。
8. **交叉校验**（如果这个规范还有另一套独立实现）：照 `compare_with_rules.js` 的样子
   写一个逐行比对脚本，把两套实现的产物钉在一起。

---

## 4. 写模板时踩过的坑（都实测过）

| 坑 | 说明 |
|---|---|
| **单位不要写在模板里** | 0.1.0 时 `suffix: " m"` 被 `.strip()` 成 `"m"`，输出 `340m`；0.2.0 已修，所以现在**单位放 suffix**（`var L = {{ L }};`），一处可改。 |
| **数值用 `{{ x }}` 或 `{{ x.value }}` 都行** | 0.2.0 起两条路径都做整数浮点规范化（`340.0`→`340`），不会再有假差异。 |
| **公式模式不能有条件** | 所以坐标映射写成 6 个变量：`x = {{ neg_x }}t{{ ax_x }};` —— `neg_x` 填 `-` 或留空，`ax_x` 填 `1/2/3`。**恒等映射也会输出这 3 行**（`x = t1;` 是空操作），因为没法条件跳过。 |
| **一行里占位符越多，公式越长** | 上限 8000 字符/格。本例最长 2702 字符（汇总表里一行 6 个占位符）。一行放 8 个以上要留意。 |
| **`case_filter` 的 Case 必须成块连续** | 公式模式的横向相对列靠"输出列 = 工况列"恒定偏移；同一规则集的工况必须相邻（本例 EXT 在前、INT 在后）。 |
| **手加 Case 列要记得填 `kind`** | 留空会回落 `default`（本例是 `EXT`），于是这一列**悄悄进了外压规则集**。`check` 能发现（见 `probes/probe_formula_structure.py`）。 |
| **加 Case 列要重跑，插变量行不用** | 公式是写死在格子里的文本：`INDEX/MATCH` 按**变量名**定位 → 插/删变量行不影响；列标是**绝对字母** → 工况列一动就过期。实测见 `probes/probe_formula_structure.py`。 |
| **`if` 必须配 `else`** | GeniE 的硬规矩（`Unable to generate optimized function`，不指行号）。模板是纯文本，工具不会替你查；本目录靠 `GeniE/Rules` 侧的自检守着。 |

---

## 5. 守卫（四层，各管一件事）

| 层 | 命令 | 抓什么 | 抓不到什么 |
|---|---|---|---|
| 工具自带校验 | `excel-codegen validate -c <yaml>` | YAML 结构、模板语法、变量未定义 / 定义了没人用、`output_sheet` 未声明、公式模式的越界写法（带行内容） | 生成出来的代码对不对 |
| 工具自带过期检查 | `excel-codegen check` | ① 公式 / 快照是否与当前 YAML + 参数一致；② **把公式在 Python 里算一遍**与 Python 渲染逐行比对（0.4.0 起默认开，`--no-values` 可关） | 公式在**真 Excel** 里的行为（区域设置、`TEXT()` 格式串、浮点显示） |
| **本目录的公式求值（第二重证据）** | `python verify_excel_engine.py <yaml> <xlsx>` | 同样"把公式算一遍再比对"，但**实现是独立的**（另一套解析器）—— 两条独立实现都过，才说明不是同一个 bug 在两处复现 | 同上（真 Excel 行为） |
| 与独立实现比对 | `node compare_with_rules.js` | 与 `GeniE/Rules` 那套带量纲检查的产物逐行一致 | 规范本身理解错了 |

---

## 6. 已知边界（工具的，不是模板的）

- **舱数据用第三层作用域**（0.8.0 起）：内压规则集在 `variables.group` 里声明
  `Tank Data` 成员表（一行一个舱），`local.tank_ref` 指向它。所以逐工况的变量里**没有**
  `l_tank` / `rho_tank` 这些行 —— 找舱数据请去 `Tank Data` 表，不要在 Local 表里找。
  证据与"改一次舱数据影响哪些工况"的实测见 `probes/probe_group_table.py`。
- **`case_filter` 下的共享变量语义**：`kind`、`draft` 这类被两套规则共用的变量，
  在合并 YAML 里只有一份定义（`compose.py` 保证一致）。
- **公式模式的 `TEXT()` 受区域设置影响**：中文/英文区域小数点是 `.`，欧洲区域是 `,`。
  本模板用到 `TEXT(x,"0.############")`，在非 `.` 区域会出问题 —— 换区域前先验一遍。
- **公式值只活在 Excel 里**：`--outdir` 导出的文件由 Python 渲染（内容等价，因为是纯替换），
  但"改参数自动同步文件"做不到，要再跑一次 `build.py`。

---

## 7. 文件清单

```
abs_fpi/
├── TEMPLATES.md                  ← 本文件：数据库索引 + 加规范的清单
├── README.md                     怎么用、产物长什么样、校验链
├── FINDINGS.md                   对 excel_codegen 的实测报告（含修复复测）
│
├── abs_fpi_external.yaml         规则集 1（真源）
├── abs_fpi_internal.yaml         规则集 2（真源）
├── compose.yaml                  配方：哪几个规则集 + case_filter
├── abs_fpi.yaml                  生成物：项目工作簿的 YAML（勿手改）
├── templates/
│   ├── ext_body.js.j2            外压 GeniE 函数体
│   ├── ext_summary.md.j2         外压工况汇总
│   ├── int_body.js.j2            内压 GeniE 函数体
│   └── int_summary.md.j2         内压工况汇总
│
├── build.py                      一条命令：compose + 生成 + 导出 + 两层验证
├── compose.py                    规则集 → 项目 YAML（合并 + 冲突检查）
├── fill_cases.py                 样例工况取值（只有首次 --init 时才灌）
├── verify_excel_engine.py        公式求值器：验证"Excel 算出来的代码"对不对
├── compare_with_rules.js         与 GeniE/Rules 逐行比对
├── probes/                       对这一版工具的最小复现与探针
│
├── abs_fpi_external.xlsx         规则集工作簿（公式模式）
├── abs_fpi_internal.xlsx         规则集工作簿（公式模式）
├── ABS_FPI_load_cases.xlsx       项目工作簿（两套规则共用 Global）
└── generated/                    导出的代码（.js）与工况汇总（.md）
```
