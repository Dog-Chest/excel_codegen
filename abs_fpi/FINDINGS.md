# excel_codegen 实测报告（ABS FPI 内外压移植）

日期：2026-09-24 ｜ 被测版本：`excel_codegen 0.1.0`，Python 3.14.5，Windows / 中文区域
被测对象：`excel_codegen`（`pyproject.toml`、`excel_codegen/*.py`、`docs/template_guide.md`）
测试方式：把 `GeniE/Rules` 里已经过量纲校验的 ABS FPI 内外压 GeniE 加载代码，
按本项目的 YAML + Jinja2 格式重写一遍（见 [README.md](README.md)），跑通全流程并逐个试探边界。

**本报告只记录问题，没有改动 `excel_codegen` 的任何文件。**
每条都给了最小复现，`probes/run_probes.py` 可以一次跑完全部探针。

> **修复会话（2026-09-24 ｜ `excel_codegen` 0.2.0）**：本报告的 12 条待修项中 **11 条已修**，
> 剩下 1 条（#3(a)「第三层作用域」）判定为**能力边界**，已写进 `docs/template_guide.md` 的
> 「能力边界与不适用场景」一节。逐条状态与复测证据见文末「修复复测记录」；
> 下面的报告正文**保持原样**（它是 0.1.0 的快照，不要照它去判断 0.2.0 的行为）。
>
> 复测结论一览：`pytest` 58 项全过；探针**原样**重跑后「残留 / 静默 / 崩溃」类现象全部消失；
> `python build.py --check` 一致；`node compare_with_rules.js` 仍然 **21 项 0 失败**。
> 本轮还顺带修掉了一个报告里没写的新 bug（元信息两条读取路径的键名不一致）。

---

# ★ 必须修改项（按优先级）—— 给修复会话

前 4 条会造成**错误交付物**或**把失败当成功**，建议先修完再做别的。
每条都给了：位置、复现、**已实测验证过的**修法（见每条的"修法"小节里带 ✅/❌ 的验证结果）。

| P | # | 一句话 | 位置 | 改动量 |
|---|---|---|---|---|
| **P0-1** | #2 | Prefix/Suffix 被 `.strip()`，`" m"` 变 `"m"`，生成物 `340m` | `excel_io.py:316-317`、`361-362` | 小（加一个取值函数） |
| **P0-2** | #8.1 | `render` 默认不写回 Excel，却照样打印「✓ 渲染完成」 | `cli.py:157`、`212-213` | **极小（2 行）** |
| **P0-3** | #5 | 重渲染的清理底边按本次行数算，旧代码残留在表里 | `excel_io.py:476-506` | 中（加 `_last_used_row`） |
| **P0-4** | #1 | 中文 Windows 下 `✓` 抛 `UnicodeEncodeError`，退出码 1 | `cli.py:217`、`332`、`46` | 小（换 ASCII + 兜底） |
| P1-5 | #8.2 | 生成的工作簿里没有一句"下一步跑什么" | `excel_io.create_template` | 中 |
| P1-6 | #8.3 | 没有指纹/时间戳/`check`，"过期"不可检测 | `excel_io.write_results` + `cli` | 中 |
| P1-7 | #3(b) | 模板 × Case 是交叉积，没有 per-template 的 Case 过滤 | `renderer.render_all` + `models.TemplateDef` | 中 |
| P2-8 | #4 | `{{ x.value }}` 不做整数浮点规范化，与文档承诺不符 | `utils.VarValue` / 文档 | 小 |
| P2-9 | #6 | YAML 同键重复被静默吞掉 | `models.load_config` | 小 |
| P2-10 | #3(a) | 没有"第三层作用域"（舱/设备子表） | `models` | 大（能力边界） |
| P2-11 | #7 | CLI `init --cases` 只接受整数 | `cli.py:93` | 小 |
| P2-12 | #8.4 | 隐藏 `Template` 表可编辑但改了没用，表内无提示 | `excel_io._build_template_sheet` | 极小 |

## 修法草案（未实施，可直接抄）

### P0-1 —— prefix/suffix 不再 strip

`excel_io.py`，加一个取值函数并替换 `read_global_values` / `read_cases` 里那四行：

```python
def _cell_or(cell_value, default):
    """只有"单元格真的为空"才回落 default；非空原样使用（首尾空格有意义）。"""
    if cell_value is None:
        return default
    if isinstance(cell_value, str) and cell_value.strip() == "":
        return default
    return to_text(cell_value)


# read_global_values（原 316-317 行）
prefix = _cell_or(worksheet.cell(row=row, column=_GLOBAL_COL["prefix"]).value, definition.prefix if definition else "")
suffix = _cell_or(worksheet.cell(row=row, column=_GLOBAL_COL["suffix"]).value, definition.suffix if definition else "")
# 后面那两行 `if not prefix and definition:` 的回落逻辑可以删掉
```

`read_cases` 里同名变量的两处（行级 361-362、逐 Case 375-378）同样处理。
验证：`probes/probe_suffix.yaml` 应当从 `340m` 变成 `340 m`，且空格子仍回落 YAML 默认。

### P0-2 —— 摘要表永远显示是否写回（2 行）

`cli.py`，把 `212-213` 行：

```python
    if write_excel:
        summary.add_row("写回 Excel", "是")
```

改成：

```python
    summary.add_row("写回 Excel", "是" if write_excel else "否（需要 --write-excel）")
```

并在 `--no-show` 之外补一句提示（可选但推荐）：

```python
    if not write_excel and outdir is None:
        console.print("[bold yellow]![/] Code 表未更新：加 --write-excel 才会写回 Excel")
```

验证：`render` 不带 `--write-excel` 时摘要必须出现"否（需要 --write-excel）"。

### P0-3 —— 清理底边取"真实底边"，不取"本次行数"

`excel_io.py`，加一个与 `_last_used_column` 对称的函数，并在两个写入函数里用它定 `bottom`：

```python
def _last_used_row(worksheet, top: int, left: int, right: int) -> int:
    """给定列区间内最后一个有内容的行号（用于清理上一次可能更长的渲染）。"""
    last = top - 1
    for row in worksheet.iter_rows(min_row=top, max_row=worksheet.max_row, min_col=left, max_col=right):
        for cell in row:
            if cell.value not in (None, ""):
                last = max(last, cell.row)
    return last


# _write_horizontal
bottom = max(row + max_lines + 1, _last_used_row(worksheet, top, column, column + len(results) - 1))
# _write_vertical
bottom = max(row + len(results) + 1, _last_used_row(worksheet, row, left, worksheet.max_column))
```

验证：`probes/run_probes.py` 的探针 D（8 行 → 3 行）与探针 I（8 Case → 1 Case）
都应从"有残留"变成"（无）"；同时探针 D2（只减列、行数不变）必须仍然干净。

### P0-4 —— `✓` / `✗` 在 GBK 控制台的处理（三种候选，已实测）

我用这个环境（`sys.stdout.encoding == 'gbk'`）把三种常见修法各跑了一遍：

| 候选 | 结果 |
|---|---|
| A. `Console(legacy_windows=False)` | ❌ **更糟**：`UnicodeEncodeError: 'gbk' codec can't encode character '\u2713'`，直接把原来能"降级"的兜底路径也去掉了 |
| B. `sys.stdout.reconfigure(errors="backslashreplace")` | ✅ 不崩，但显示成字面量 `\u2713 配置校验通过`（`✗` 现在其实已经走的是这条降级路径） |
| C. 把 `✓` / `✗` 换成 ASCII（`OK` / `ERROR`） | ✅ 干净可用 |

建议 **C + B 一起**：`cli.py` 顶部加

```python
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(errors="backslashreplace")  # 任何非 GBK 字符都不再抛异常
    except Exception:
        pass
```

再把 `✓`（`cli.py:217`、`332`）与 `_fail` 里的 `✗`（`cli.py:46`）换成 ASCII。
**明确不建议候选 A。** 另外建议补一条测试：把 `sys.stdout` 换成只支持 GBK 的流再跑一遍
`validate`，断言退出码为 0。

---

## 结论摘要

| # | 严重度 | 问题 | 影响面 | 复现 |
|---|---|---|---|---|
| 1 | **高** | 中文 Windows（GBK 控制台）下 `validate` / `render` 收尾打印 `✓` 抛 `UnicodeEncodeError`，退出码 1 | 所有中文 Windows 用户；CI 把成功当失败 | 命令行 |
| 2 | **高** | Prefix/Suffix 读回时被 `.strip()`，带前导空格的单位被吞掉（`" m"` → `"m"` → `340m`） | 一切"值 + 单位"要带空格的生成目标，且**静默** | `probes/probe_suffix.yaml` |
| 3 | 中 | 只有 global / local 两种作用域，且模板 × Case 是交叉积 —— 没有"第三层作用域"，也没有 per-template 的 Case 过滤 | 多规范集共用一本工作簿；舱/设备等子表 | `probes/probe_twosets.yaml` |
| 4 | 中 | `{{ x.value }}` 不做整数浮点规范化，与 README / 指南的承诺不一致（`340.0` vs `340`） | 按文档写 `{{ x.value }}` 的人 | `probes/probe_value.yaml` |
| 5 | 中 | 重渲染的清理范围底边由"本次结果数"算出，旧内容会残留 | 改短模板 / 减少 Case 后 Output 表留旧代码 | `probes/probe_long.yaml` + `probe_short.yaml`、`probe_vert.yaml` |
| 6 | 低-中 | YAML 同键重复被静默吞掉，错误在三步之后以另一种形式出现 | 手写大 YAML 的人 | `probes/probe_dupkey.yaml` |
| 7 | 低 | CLI `init --cases` 只接受整数，库函数却支持 Case 名字列表 | 批量建有意义的工况名 | 命令行 |
| 8 | **高** | `render` 不带 `--write-excel`（默认）时**什么都不写回，却照样打印「✓ 渲染完成」**；工作簿本身也没有任何"下一步跑什么"的说明 | 所有用户："改了参数，代码没变" | 命令行 + 工作簿

全流程本身**跑通了**：2 个规范集、7 个工况全部渲染成功，导出的 7 段 GeniE 代码与
`GeniE/Rules` 的产物**逐行一致**（`node compare_with_rules.js` → 21 项全过）。
好消息与坏消息都在下面。

---

## #1（高）GBK 控制台：收尾打印 `✓` 抛 UnicodeEncodeError，退出码 1

**位置**：`excel_codegen/cli.py:217`（`render_command` 末尾）、`cli.py:332`（`validate_command` 末尾）

```python
console.print("[bold green]✓[/] 渲染完成")
...
console.print("[bold green]✓[/] 配置校验通过")
```

**环境**：Windows + 中文区域，`sys.stdout.encoding == 'gbk'`，Python 非 UTF-8 模式。

**复现**

```powershell
python -m excel_codegen validate -c abs_fpi_external.yaml -x external.xlsx
```

**实际**

```
UnicodeEncodeError: 'gbk' codec can't encode character '\u2713' in position 0: illegal multibyte sequence
[exit code: 1]
```

调用栈：`rich/_windows_renderer.py:17 legacy_windows_render` → `write_styled` → `write_text`
→ `UnicodeEncodeError`。

**关键点：工作已经做完了才崩。**
`render --write-excel --outdir generated` 时，Excel 的 Output 表和 6 个导出文件都已写好，
最后一行 `✓` 才抛异常并以 1 退出 —— 在 CI 里这就是"构建失败"，而人也会以为没渲染成功。

**边界**：`init` 不受影响（它不打印 `✓`，表格线 `│─┌┐` 落在 GBK 里）。
错误路径的 `✗`（`_fail`，走 `error_console`）**没有崩** —— rich 在 stderr 上退化成
`\u2717` 字面量，所以只看到乱码。也就是说 `✓` 与 `✗` 在不同 stream 上行为不一致。

**临时绕过**（本仓库的脚本都这么跑）

```powershell
$env:PYTHONUTF8='1'      # 或 chcp 65001
```

**建议修法**（不实施）

- 最省事：把 `✓` / `✗` 换成 ASCII（`OK` / `ERROR` 或 `[ok]` / `[fail]`）。
- 或者构建 `Console` 时避开 legacy 渲染路径，并给 stdout/stderr 加 `errors="replace"`。
- 或者把这两处收尾打印包在 `try/except UnicodeEncodeError` 里 —— 至少别用退出码撒谎。
- 无论怎么修，都建议补一条测试：把 `sys.stdout` 换成不支持的编码再跑 CLI。

---

## #2（高）Prefix / Suffix 被 `.strip()`，带前导空格的单位被吞掉

**位置**：`excel_codegen/excel_io.py:316-317`（`read_global_values`）、`excel_io.py:361-362`（`read_cases`）

```python
prefix = to_text(worksheet.cell(row=row, column=_GLOBAL_COL["prefix"]).value).strip()
suffix = to_text(worksheet.cell(row=row, column=_GLOBAL_COL["suffix"]).value).strip()
...
if not prefix and definition is not None:
    prefix = definition.prefix  # ← 只有"strip 完是空"才回落
```

`create_template` 是把 YAML 的 `suffix` **原样**写进单元格的
（`excel_io.py:154`），所以 `suffix: " m"` 在单元格里是 `" m"`，
读回来 `.strip()` 成 `"m"`，于是 `str(VarValue)` 得到 `340m`。

**复现**：`probes/probe_suffix.yaml`（`suffix: " m"` / `" kg/m^3"`）

```
YAML 里的 suffix : [('L', ' m'), ('rho_sea', ' kg/m^3')]
Excel E 列 L         = ' m'
Excel E 列 rho_sea   = ' kg/m^3'
上下文     L         prefix='' suffix='m' -> str='340m'        ← 空格没了
上下文     rho_sea   prefix='' suffix='kg/m^3' -> str='1025kg/m^3'
```

渲染结果：

```
combined   L           = 340m
combined   rho_sea     = 1025kg/m^3
suffix     L.suffix    = [m]
filter     pvs("", " m") = [340 m]        ← 过滤器路径没被 strip，是对的
```

**为什么这条最值得先修**

1. 它是**静默**的：不报错、不警告，只是生成出来的文本少了空格。
2. 它和文档矛盾。`template_guide.md` §4.5 只说"Prefix/Suffix 单元格留空 → 用 YAML 的"，
   没有说"非空也会被去空格"；§4.4 甚至建议用 `pvs` 自己拼前缀。
3. 对本用例是致命的：GeniE 的 `var L = 340 m;` 变成了 `var L = 340m;`。
   我无法从帮助文档确认 GeniE 是否接受 `340m`（文档是 API 索引，不是语法），
   但生成物与手写脚本、与 `GeniE/Rules` 的输出**不再一致**，这本身就不可接受。
4. `init` 写出去、`render` 读回来 —— **自己写的自己读不对**，是最容易撞上的一类。

**本次的绕过**：单位写在模板里（`var L = {{ L }} m;`），不用 suffix。
这只是绕过，不是修复。

**建议修法**（不实施）

```python
raw = worksheet.cell(row=row, column=_GLOBAL_COL["suffix"]).value
suffix = definition.suffix if raw is None or str(raw).strip() == "" else to_text(raw)
```

即：**只在"单元格为空"时回落，非空时原样使用**。这与文档语义一致，
`pvs` / `wrap` 的行为也已经是这样了。

---

## #3（中）只有两种作用域，且模板 × Case 是交叉积

### (a) 没有"第三层作用域"

`VariablesConfig` 只有 `global` / `local`。ABS 内压需要三层：船（global）、工况（local）、
**舱**（一个舱被多个工况引用）。本工具表达不了，只能把舱的 13 个参数按工况摊平：
WBT6 被两个工况用到，它的数据就在表里写两遍；改一个舱的尺寸要改 N 列。

这不是"做错了"，是**能力边界**，但值得写进文档的"不适用场景"里，
因为它直接决定了移植方案（我这一版就是摊平写的，`tank_ref` 列只作标注）。

### (b) 模板 × Case 是交叉积，没有 per-template 的 Case 过滤

`renderer.render_all` 的循环是外层 templates、内层 cases，结果必定是
`len(templates) × len(cases)`。`--case` 是**全局**过滤，不是 per-template 的。

**复现**：`probes/probe_twosets.yaml` —— 一个 Local 表、两个模板（`ext_code` / `int_code`）、
两个 Case（一个 `kind=EXT`、一个 `kind=INT`）：

```
期望：ext_code 只出 case_ext，int_code 只出 case_int（各 1 个）
实际：4 个渲染结果
  ext_code  -> ['case_ext', 'case_int']
  int_code  -> ['case_ext', 'case_int']
Code EXT  表头行 = ['case_ext', 'case_int']
Code INT  表头行 = ['case_ext', 'case_int']
```

**后果**：我原本想在**一本工作簿**里放"共用 Global（船的主尺度只填一次）+ 外压模板 + 内压模板"
（这也是我上一版 `GeniE/Rules` 的做法），在本工具里表达不出来，
只能拆成 `abs_fpi_external.yaml` / `abs_fpi_internal.yaml` 两套，
**主尺度要填两遍** —— 正是我想避免的那种"两边不一致"。

**建议修法**（不引入破坏性改动的两个方向）

- 给 `TemplateDef` 加一个可选的 `case_filter`（Jinja 表达式，对 Case 的变量求值），
  例如 `case_filter: "kind == 'EXT'"`；`render_all` 里在渲染前筛掉不匹配的 Case，
  `write_results` 的表头也就自然只写该模板适用的列。
- 或者给变量加 `scope_key: tank`（或 `group`），让一个 local 变量声明"我的取值来自哪张子表"，
  这样舱层可以用"每个 Case 引用一个子表键"表达。

---

## #4（中）`{{ x.value }}` 不做整数浮点规范化

**文档承诺**（`README.md`「数值与 Prefix / Value / Suffix」小节、`template_guide.md` §3.2）：

> 数值变量不会出现 `115200.0` 这种尾巴：整数浮点在转字符串时被规范化为 `115200`。

**实现**：规范化只在 `VarValue.__str__`（→ `utils.to_text`）里做。
`{{ x }}` 走的是 `__str__`，对；`{{ x.value }}` 拿的是原始 Python 值，
`float 340.0` 直接 `str()` 成 `"340.0"`。而 `template_guide.md` §4.2 恰恰把
`{{ port.value }}` 推荐为"取纯值"的写法。

**复现**：`probes/probe_value.yaml`

```
L   type=float value=340.0  str(v)='340'  repr(v.value)=340.0
g   type=float value=9.81   str(v)='9.81' repr(v.value)=9.81
Li  type=int   value=340    str(v)='340'  repr(v.value)=340
La  type=auto  value=340    str(v)='340'  repr(v.value)=340
draft type=float value=20.0 str(v)='20'

渲染结果：
L      (float) combined = 340        .value = 340.0     ← 不一致
draft  (float) combined = 20         .value = 20.0      ← 不一致
```

**影响**：按文档写 `{{ x.value }}` 的人会得到 `340.0` / `20.0`。
在 GeniE、C、Python 里多数能编译，所以不会立刻炸，但：
单位/数值的文本形态与另一条路径不同，做 diff、做 golden file 就会一直看到假差异。

**建议修法**（不实施）

- 最小改动：`VariableDef` 加一个"规范化后的纯值"入口（`VarValue.text` 已经有了，
  它返回的就是 `to_text(self.value)`），文档改成推荐 `{{ x.text }}`。
- 或者让 `value` 在 `type` 为数值时存规范化后的值（会改变 `type: auto` 的语义，需谨慎）。
- 至少要把 README 那句话限定为"`{{ x }}` 组合值"，并在 §4.2 旁边注明
  `.value` 是原始类型、不做规范化。

---

## #5（中）重渲染的清理范围用"本次结果数"算底边，旧内容残留

**位置**：`excel_io.py:480`（horizontal）、`excel_io.py:501-504`（vertical）

```python
def _write_horizontal(worksheet, template, results, column, row):
    max_lines = max(result.line_count for result in results)
    ...
    bottom = row + max_lines + 1
    last_column = _last_used_column(worksheet, top, bottom, column)  # 只看这几行
    _clear_region(worksheet, top, bottom, column, last_column)
```

`_last_used_column` 也只在 `top..bottom` 这个行带里找"最右列"，所以**底边之外的东西
既不会被发现、也不会被清掉**。而底边是拿**本次**的行数算的。

`README.md` 与 `write_results` 的 docstring 都说"重新渲染时会自动清理上一次写下的区域，
不会残留旧 Case 的内容"—— 这个承诺只在"新内容不比旧内容短"时成立。

### 症状 A：horizontal，行数变短（`start_cell: B2`）

**复现**：`probes/probe_long.yaml`（8 行）渲染一次，再用 `probes/probe_short.yaml`（3 行）
写同一个文件、同一张 `Output`、同一个 `B2`。

```
长渲染之后 B 列： ['Case1', line1..line8 of LONG, None]
短渲染之后 B 列： ['Case1', line1..line3 of SHORT, None, None, 'line 6 of the LONG render',
                  'line 7 of the LONG render', 'line 8 of the LONG render', None]
残留的 LONG 行：['line 6 of the LONG render', 'line 7 of the LONG render', 'line 8 of the LONG render']
```

对照：`bottom = 2 + 3 + 1 = 6`，所以只清了第 1..6 行；旧的第 7..9 行（LONG 的第 6..8 行）原封不动。

> 反例（说明根因就在 `bottom` 这一行）：**只减少 Case 列数、行数不变**时第 3 列被清得很干净 ——
> 因为 `bottom` 仍在旧内容之下。`probes/run_probes.py` 的探针 D2 就是这个反例。

### 症状 B：vertical，Case 数大幅变少

**复现**：`probes/probe_vert.yaml`，先 8 个 Case 再只渲染 1 个。

```
8 个 Case 之后 A 列行 1..10： [None, Case1..Case8, None]
只渲染 1 个 Case 之后 A 列：  [None, 'Case1', None, None, Case4, Case5, Case6, Case7, Case8, None]
Case2..Case8 残留：['Case4', 'Case5', 'Case6', 'Case7', 'Case8']
```

`bottom = row + len(results) + 1 = 4`，清的还是第 2..4 行。
（3 个 Case 减到 1 个时不会露，因为 `len(results)+1` 恰好覆盖 —— 这也是为什么它不容易被发现。）

**为什么这条对"生成代码"这个用途是危险的**：用户改短模板 → 重渲染 → Output 表里
下半段还是旧代码。复制整列粘进 GeniE 的人会把上一版的计算一起带走，
而 GeniE 只会算，不会报警。

**建议修法**（不实施）
清理范围的下边界不要用"本次行数"，而要用**该区域原有内容的实际下边界**：
先扫 `start_cell` 起、`worksheet.max_row` 止的整块找出真正的底边，再 `_clear_region`，最后写。
或者干脆把"上一次写过的区域"记在工作表里（例如藏在 `Template` 表的元信息中），
渲染时按记录清理。

---

## #6（低-中）YAML 同键重复被静默吞掉，错误在三步之后才出现

**位置**：`models.load_config` 用 `yaml.safe_load`；PyYAML 对重复键**不报错**，后者覆盖前者。

**复现**：`probes/probe_dupkey.yaml`（`variables:` 出现两次，第一份只有 `global`，第二份只有 `local`）

```
load_config 没有报错；global 变量数 = 0，local 变量数 = 1
global 名字 = []
生成的 Global 表数据行 = （空）
直到渲染才炸：RenderError: 模板 'probe_dupkey' 变量缺失: 'L' is undefined
（请检查 YAML 的 variables 与 Excel 中的变量名是否一致）
```

**问题不在"报错"，在于报错了但指向错的地方**：
`L` 明明定义在 YAML 里，工具却让人去"检查 YAML 的 variables 与 Excel 中的变量名是否一致"；
而中间那个**空 Global 表**被写进工作簿、看起来像"工具没写进去"。
（我自己第一次写这份 YAML 就犯了这个错，两个 `variables:` 块。）

**建议修法**（不实施）
用自定义 Loader 检测重复键并报错（约 10 行）：

```python
class _NoDupLoader(yaml.SafeLoader): ...


def _no_dup(loader, node, deep=False):
    mapping = {}
    for key_node, value_node in node.value:
        key = loader.construct_object(key_node, deep=deep)
        if key in mapping:
            raise ConfigError(f"YAML 键重复: {key!r}（第 {key_node.start_mark.line + 1} 行）")
        mapping[key] = loader.construct_object(value_node, deep=deep)
    return mapping


_NoDupLoader.add_constructor(yaml.resolver.BaseResolver.DEFAULT_MAPPING_TAG, _no_dup)
```

另外 `validate` 里可以顺手加一条："Global 表有 0 行数据，但模板引用了变量" —— 现在这种情况
只在 render 时才暴露。

---

## #7（低）CLI `init --cases` 只接受整数

**位置**：`cli.py:93` `cases: int = typer.Option(2, "--cases", min=1, ...)`；
而库函数是 `create_template(config, path, *, cases: int | Sequence[str] = 2, ...)`。

**复现**

```
$ python -m excel_codegen init -c probe_suffix.yaml --cases Case1
Invalid value for '--cases': 'Case1' is not a valid int range.
exit=2
```

**影响**：要给 7 个工况起有意义的名字（`EXT-T20.559-mu90-kf+1`），
只能 init 之后再改表头 —— 本次是用 `fill_cases.py`（openpyxl）顺手改的。
对"一列一工况、工况名就是块名"的工作流，这是个每轮都要付的成本。

**建议修法**（不实施）：`--cases` 接受逗号分隔的名字列表（同时保留 `--cases N` 的老写法），
或另加一个 `--case-names a,b,c`。

---

## #8（高）"改了参数，代码没变" —— 默认的 `render` 静默不写回，工作簿也不说该跑什么

这条是实际使用时最先撞上的：**打开 `external.xlsx` 改了参数，`Code` 表里的代码没有任何变化。**

先说清楚**不是**哪一类问题：生成本身是对的。作者实测过（见下面的证据），
只要跑 `render --write-excel`，`Code!B6` 会立刻从 `var L = 340 m;` 变成 `var L = 400 m;`。
`excel_codegen` 是**快照式**工具：Excel 是"表单"，`Code` 表是 `render` 写下的一份快照，
不是活公式 —— 这一点在 `README.md` 与 `template_guide.md` §1 的心智模型里是讲了的。

但这里有三个真实的缺陷叠在一起，使得"改了参数、代码没变"变成**默认结局**。

### 8.1 `render` 不带 `--write-excel` 时照样报成功

**位置**：`cli.py:157` `write_excel: bool = typer.Option(False, "--write-excel", "-w", ...)` ——
**默认不写回**；`cli.py:212-213`

```python
if write_excel:
    summary.add_row("写回 Excel", "是")  # ← 不写回时这一行只是"不出现"
```

**实测**（`external.xlsx` 里把 L 改成 500，然后两种跑法）：

```
$ excel-codegen render -c abs_fpi_external.yaml -x external.xlsx --no-show
│ 渲染结果 │ 6 个（2 模板 × 3 Case） │
✓ 渲染完成                            ← 什么都没写回，但说了"完成"
  -> Code!B6 = 'var L = 340 m;'       ← 还是旧值

$ excel-codegen render -c abs_fpi_external.yaml -x external.xlsx --write-excel --no-show
│ 写回 Excel │ 是 │
✓ 渲染完成
  -> Code!B6 = 'var L = 500 m;'       ← 这次才变
```

**问题**：不带 `--write-excel` 时的摘要表**没有"未写回"这一行**，
只有"写回 Excel │ 是"在带参数时才出现。用户看到的是同一句 `✓ 渲染完成`。
对一个"Excel 是主要交付物"的工具，"成功"与"成功但没动文件"必须能一眼分开。

**建议修法**（不实施）：把那一行改成永远输出，例如
`写回 Excel │ 否（需要 --write-excel）`，并在没写回时补一句
`提示：Code 表未更新；加 --write-excel 才会写回`。
（是否把 `--write-excel` 改成默认打开，是产品决定；但至少不能沉默。）

### 8.2 生成的工作簿里没有任何"下一步做什么"的说明

`init` 生成的工作簿只有 `Global Parameter` / `Local Parameter` / Output 表 / 隐藏的 `Template`
—— **没有说明表**。实测：

```
external.xlsx :: Global Parameter   state=visible dims=A1:E15
external.xlsx :: Local Parameter    state=visible dims=A1:G10
external.xlsx :: Code               state=visible dims=B1:D83
external.xlsx :: Summary            state=visible dims=A2:P4
external.xlsx :: Template           state=hidden  dims=A1:B133
Code!A2 = None                       ← Code 表本身也没有一句说明
```

用户拿到的是一个"填了数就以为会自动出代码"的表。真正要跑的命令只在
`init` 的**终端输出**里出现过一次，关掉终端就没了；
`docs/template_guide.md` §1 与 README「快速开始」里有，但那是**文档**，不是**工作簿**。

**建议修法**（不实施）：`create_template` 顺手生成一张 `HOWTO`（或 `README`）表放在最前面，
写清三步与那条命令；`write_results` 之后把"生成时间 / 参数指纹 / 输出指纹"也写进去
（**本目录的 `build.py` 已经这么做了，可以直接照抄那张表的字段**）。
这样"这份代码是什么时候、由哪一版参数生成的"就跟着文件走。

### 8.3 没有任何"过期"标记

`Code` 表与参数在同一个文件里，但两者之间**没有任何一致性信息**：
没有时间戳、没有指纹、没有 `--check`。所以：
改了参数没重跑 → 表里是新参数、旧代码，看起来完全正常；
配合 #5（清理范围算错、旧行残留），旧代码还会以"多出来几行"的形式混进去。

**建议修法**（不实施）：渲染时把参数的指纹写进 `Template` 表（它已经是隐藏的元信息表），
再给 CLI 加一个 `excel-codegen check`（或 `render --check`）：重算指纹/重新渲染并与表内比对，
不一致就非 0 退出。**这一条同时是 CI 里的护栏**：可以断言"提交的工作簿与其代码是同步的"。

### 8.4 顺带的兄弟陷阱：隐藏的 `Template` 表可编辑、但改了没用

`template_guide.md` §6.4 说了"它是**只读参考**，修改它不会影响渲染结果"，
但表本身没有任何视觉提示（它只是 `hidden`，解隐藏之后和别的表长得一样）。
用户如果在那里改模板、再渲染发现没变化，会得到与 8.1 一模一样的困惑。
建议在表的 A1 或元信息里加一行 `<-- 只读参考：模板真源是 YAML / template_file`。

### 8.5 本目录的临时对策（不改项目）

`abs_fpi/build.py` 把整条链封成一条命令，并把说明写进工作簿：

```bash
python build.py            # 渲染 → 写回 Excel → 导出代码 → 刷新 HOWTO 表
python build.py --check    # 只检查 Code 表是不是已经过期（非 0 退出）
python build.py --init     # 首次生成骨架
```

`--check` 的实测输出：

```
external.xlsx
  记录时的参数指纹  e0ff46fe8dc0
  当前的参数指纹    c0d598d8db3d   ← 参数改过了
  ✗ Code 表已过期：
      Code 第 2 列（EXT-T20.559-mu90-kf+1）：第 5 行不同：表里 'var L = 380 m;'，应为 'var L = 395 m;'
      Code 第 3 列（EXT-T15-mu0-kf-1）：第 5 行不同：表里 'var L = 380 m;'，应为 'var L = 395 m;'
    → 跑一次 `python build.py` 刷新
```

这是**绕过**，不是修复：真正的修复应该在 `cli.py` / `excel_io.py` 里（8.1–8.4）。
但 `build.py` 里那张 HOWTO 表的字段设计（时间 / 参数指纹 / 输出指纹 / 工况一览）
可以直接拿去用在 `create_template` 里。

---

## 顺带确认没问题的几件事

写完报告也要把"好"的说清楚，免得下次重复验证：

| 项 | 结果 |
|---|---|
| 全流程 | 2 个规范集、7 个工况，`init` → 填写 → `validate` → `render --write-excel --outdir` 全部跑通 |
| 生成正确性 | 7 段 GeniE 代码与 `GeniE/Rules` 的产物**去注释后逐行一致**（`compare_with_rules.js`，21 项） |
| Jinja2 空白控制 | `{%- ... -%}` 用对之后**零空行**，代码不会莫名下移一行 |
| 隐藏 `Template` 表 | `render --write-excel` 之后仍然是 `hidden`，内容还在（探针 H） |
| 右拉加 Case | `Local Parameter` 的 Case 列扫描逻辑按预期工作（本项目的 `_case_columns` 遇到空表头才停，这也是文档写明的） |
| 长行 | `_write_vertical` 的列宽 `min(160, len+2)`；示例里最长的一行（内压汇总的 beta 行）约 120 字符，无异常 |
| 错误提示质量 | `start_cell: "2B"`、`output_sheet` 未声明、Case 列缺失等都在配置阶段就报出来，指向具体字段（探针 F 顺带验证） |
| 自带测试 | `pytest` 38 项全过，与本次实测没有冲突 |

## 几点非缺陷的观察（可以考虑写进文档）

1. `validate` **不会**提示"YAML 里定义了、但没有任何模板引用"的变量。
   对内压的 `tank_ref` 这种纯标注变量是好事；但变量名拼错（模板写 `{{ draf }}` 而定义是 `draft`）时，
   `StrictUndefined` 会报错 ✓，可反过来"定义了却没人用"就完全没人管 —— 那是"参数改了没生效"的典型来源。
2. `--cases` 生成的 Case 名固定是 `Case1..CaseN`，而 `filename: "{{ case_name }}.js"`
   让文件名直接吃工况名；工况名里出现 `+`、`.`, 是安全的（`safe_filename` 只换 `<>:"/\|?*` 与控制字符），
   这一点文档没说，值得补一句。
3. `template_guide.md` §7 的图里"第 n 行的内容来自渲染结果的第 n 行"是对的，
   但没说**渲染结果里的空行也会各占一行**。写 `{% if %}` 时不加空白控制就会多出空行，
   对"逐行写进 Excel 再整列复制"的用法是实打实的噪音 —— 建议在 §5.3 之前加一句警告。

---

# 修复复测记录（2026-09-24）

被测版本从 `0.1.0` 升到 `0.2.0`（`pyproject.toml` / `excel_codegen/__init__.py` 同步）。
本节记录"改了什么、拿什么证据确认改对了"。命令都在 `excel_codegen/` 下执行，
`<venv>` 指仓库内的 `.venv\Scripts\python.exe`。

## 复测命令与结果

| 命令 | 结果 |
| --- | --- |
| `<venv> -m pytest` | **58 passed**（0.1.0 时 38 项；新增 20 项回归，覆盖下面每一条） |
| `<venv> abs_fpi/probes/run_probes.py` | exit 0，`probes/_probe_run.txt`（UTF-8，138 行） |
| `<venv> abs_fpi/build.py` | 两个工作簿刷新成功（`[ok] … 参数指纹 abc6f4021c38 / 7a723e97e95a`） |
| `<venv> abs_fpi/build.py --check` | `[ok] Code 表与当前参数一致`（两个工作簿，exit 0） |
| `node abs_fpi/compare_with_rules.js` | **21 checks, 0 failure(s)**（与 `GeniE/Rules` 的产物仍然逐行一致） |
| `PYTHONIOENCODING=gbk <venv> -m excel_codegen validate -c …` | exit 0（0.1.0 时是 `UnicodeEncodeError` + exit 1） |

## 逐条状态

| # | 严重度 | 结论 | 落点（0.2.0） | 复测证据（探针 / 用例） |
| --- | --- | --- | --- | --- |
| 1 | 高 | **已修** | `cli.py`：标记改 ASCII（`OK`/`ERROR`/`!`），导入时给 stdout/stderr 加 `errors="backslashreplace"` | 探针 F：GBK 下 `validate` → `OK 配置校验通过`，exit 0；`test_cli_survives_gbk_stdout`、`test_cli_success_marker_is_ascii` |
| 2 | 高 | **已修** | `excel_io._cell_or/_text_or`：只有单元格真为空才回落，非空原样（不再 `.strip()`） | 探针 A：`suffix=' m'` → `340 m`；`test_prefix_suffix_keep_leading_space` |
| 3(a) | 中 | **不改（能力边界）** | 文档新增「能力边界与不适用场景」 | `docs/template_guide.md` §12 |
| 3(b) | 中 | **已修** | `TemplateDef.case_filter` + `renderer.compile_case_filter/case_matches/FilterValue` | 探针 J：`ext_code -> ['case_ext']`、`int_code -> ['case_int']`；`test_case_filter_limits_template_to_matching_cases` |
| 4 | 中 | **已修** | `renderer.build_environment` 的 `finalize`：`{{ x.value }}` 与 `{{ x }}` 形态一致 | 探针 B：`L (float) .value = 340`（0.1.0 时 `340.0`）；`test_value_attribute_is_normalised_like_combined` |
| 5 | 中 | **已修** | `excel_io._write_horizontal/_write_vertical`：清理底边改用 `_last_used_row` 扫真实底边 | 探针 D「残留的 LONG 行：（无）」、探针 I「Case2..Case8 残留：（无）」、探针 D2 仍干净；`test_rerender_clears_older_longer_output`、`test_vertical_rerender_clears_extra_cases` |
| 6 | 低-中 | **已修** | `models._StrictLoader`：重复键在**配置阶段**报错并给行号 | 探针 E：`ConfigError: YAML 键重复: 'variables'（第 20 行）`；`test_duplicate_yaml_key_is_rejected` |
| 7 | 低 | **已修** | `cli._parse_cases`：`--cases` 接受 `3` 或 `EXT-T20,INT-T15`；`create_template(include_howto_sheet=…)` 同步暴露 | 探针 G：`--cases case_alpha,case_beta` → exit 0；`test_cli_init_accepts_case_names` |
| 8.1 | 高 | **已修** | `cli.render_command`：摘要**永远**打印"写回 Excel"行，未落盘时再补一条 `!` 提示 | 探针 K：`写回 Excel │ 否（需要 --write-excel）` + `! 本次只预览…`；`test_cli_render_never_silently_skips_writeback` |
| 8.2 | 高 | **已修** | `excel_io.write_howto_sheet`：`create_template` 生成 `HOWTO` 表（第一张），写清三步 + 命令 + 快照提醒 + 每个模板的输出位置 | 探针 H / `test_create_template_layout`；`--howto/--no-howto` 可关闭（`test_cli_init_no_howto`） |
| 8.3 | 高 | **已修** | `write_results` 写入时间 / 参数指纹 / 输出指纹 / 工况一览（HOWTO 表 + 隐藏 Template 表的 `## excel-codegen-meta` 块）；新增 `excel-codegen check` 命令，过期即 exit 1 并指出第几行不同 | 探针 K：写回后 `check` → `记录 80ecfe5dc6e7 / 当前 80ecfe5dc6e7`、exit 0；`test_cli_check_reports_stale_and_fresh`、`test_cli_check_detects_missing_render`、`test_write_results_records_fingerprints` |
| 8.4 | 极小 | **已修** | `Template` 表首行加"!! 只读参考：模板真源是 YAML / template_file"红字提示 | `test_create_template_layout` 断言 |
| 观察 1 | — | **已做** | `validate` 新增"YAML 定义了但没有被任何模板引用"的告警 | `test_cli_validate_warns_about_unused_variable` |
| 观察 2 | — | **已做（文档）** | `template_guide.md` §8 说明 `safe_filename` 只替换 `<>:"/\|?*` 与控制字符（`+`、`.` 安全） | — |
| 观察 3 | — | **已做（文档）** | `template_guide.md` §5.3 / §7 增加"空行也各占一行"的警告 | — |

## 本轮新发现（报告里没有的）

**元信息两条读取路径的键名不一致**：`read_metadata` 从隐藏 `Template` 表的
`## excel-codegen-meta` 块读出来是中文键（`时间` / `参数指纹` / `输出指纹`），
而从 `HOWTO` 表（`template_sheet: null` 时的退化路径）读出来却是英文键
（`generated_at` / `input_fingerprint` / `output_fingerprint`）。
于是 `template_sheet: null` 的工作簿里，`check` 永远显示"记录 （无）"，
明明有记录却判定"没有渲染记录" —— 属于本轮引入功能时踩到的坑，已统一为中文键，
并加了回归用例 `test_read_metadata_howto_fallback_uses_same_keys`。

> 这条正好是原报告 #8.3 想解决的问题的一个变体：**"有没有记录"和"记录能不能被读到"
> 是两件事**，只写不读（键名不匹配）会静默退化成"无记录"。所以 `check` 在
> 完全没有记录时会明确打出 `!` 提示，而不是假装通过。

## 明确不做 / 改法不同的地方

1. **#3(a) 第三层作用域**：`global / local` 之外再加"舱 / 设备"一层会改变变量解析、
   Excel 表结构与 `render` 循环的语义，属于产品级改造。本轮只做到
   "两根模板共用一本工作簿 + 各写一列"（case_filter），并在文档里把
   "舱数据按工况摊平"标为已知代价。
2. **#5 的另一种修法**（把"上次写过的区域"记在元信息里按记录清理）没做：
   现在用"扫真实底边"实现，逻辑更短、不依赖元信息是否被删。
3. **`--write-excel` 没有改成默认打开**：改默认值会让"只想看看渲染结果"的人
   意外改动交付物。改用"摘要永远显示 + 未落盘时告警"把状态说清楚（#8.1 的建议）。

## 追加：0.3.0 的公式模式（针对"改参数不想跑脚本"）

0.2.0 之后又补了一档能力：`engine: excel` —— 输出表里写 **Excel 公式**而不是文本快照，
所以改 Global / Local 的参数之后，**Excel / WPS 打开就重算，不用跑任何脚本**。

* 只支持"纯替换"子集：`{{ x }}` / `.value` / `.text` / `.prefix` / `.suffix` / `{{ case_name }}` / `{{ template_name }}`。
* 出现 `{% if %}` / `{% for %}` / 过滤器会**直接报错并给出行内容**。
* 公式用 `INDEX/MATCH` 按变量名定位（插行不指错）、`ISBLANK` 处理空单元格回落、
  `TEXT()` 规范化数值；横向布局用相对列（右拉换 Case）、纵向布局用绝对列。
* 代价：值只活在 Excel 里（`--outdir` 导出与 `check` 的比较仍走命令行），
  改模板仍要重跑一次 `--write-excel` 刷新公式。

**本项目的内外压模板继续用 `engine: snapshot`**：它们有大段 `{% if %}` 分段链与 `{% for %}`，
超出"纯替换"子集 —— 这正是 §14.6 说的"模板复杂 → 快照模式"。

> ⚠ **上面这两句是错的**（把"生成目标语言的 `if/else`"当成了 Jinja 控制流）。
> 下一节「0.3.0 复测」给出了更正与实测：本项目模板**可以**进公式模式。

公式模式的最小演示见探针 L（`probes/probe_formula.yaml`）：
`fullCalcOnLoad = True`，`Output!B4` 是公式，把 Local 表改成 12 之后 Python 侧渲染给出
`var draft = 12 m;` —— 在 Excel 里这一步只需要"打开文件"。

---

# 0.3.0 复测：内外压模板改用公式模式（本节更正上面那个结论）

> 复测日期 2026-09-24 ｜ 被测：`excel_codegen 0.3.0`（`engine: excel`）
> 复测产物：`abs_fpi_external.xlsx` / `abs_fpi_internal.xlsx` / `ABS_FPI_load_cases.xlsx`
> 一句话：**ABS FPI 的内外压模板可以进公式模式**，上面那段"必须留在 snapshot"的判断是错的。

## A. 更正：那"大段 `{% if %}`"其实是 GeniE 的字面文本

上一节写的是"它们有大段 `{% if %}` 分段链与 `{% for %}`"。实际上 0.1.0 版模板里的
**Jinja** 控制流只有开头那 5 行坐标映射：

```jinja
{%- set ax = {'x': GRT_ax_x.value | int, ...} -%}
{%- set identity = ax.x == 1 and ax.y == 2 and ax.z == 3 and ... -%}
{% if not identity -%}
{% for a in ['x', 'y', 'z'] -%}
{{ a }} = {% if sg[a] %}-{% endif %}t{{ ax[a] }};
{% endfor -%}
{% endif -%}
```

而 `C_1` 区间、`k_lo`、Girth 插值、`k_s`、`GM`、`theta` 那些 `if/else`，
**是模板里的字面文本**（将来要粘进 GeniE 的语句），不是 Jinja 控制流 ——
公式编译器只拒 `{%` / `{#` / `|`，对 `{`、`}`、`;`、`&&` 一概不管。

把这个前导改写成**纯替换**（6 个变量：取哪个轴 + 是否取反）之后，整个模板就落进
"纯替换"子集了：

```jinja
x = {{ neg_x }}t{{ ax_x }};
```

于是三个工作簿全部改成 `engine: excel`。实测（探针 M：`probes/probe_genie_shape.yaml`
就是"字面 if/else 链 + 纯替换映射 + `engine: excel`"的最小组合）：

```
│ genie_shape │ excel·公式 │ Output @ B2 │
OK 配置校验通过
Output!B4 = ="var t2 = y;"
B6 = ="x = "&IF(ISBLANK(INDEX('Global Parameter'!$D:$D,MATCH("neg_x",...))),"",...)&"t"&...
```

**给文档的建议**：`docs/template_guide.md` §14.6 与本节上一段都把"模板里有 `if/else`"
当成了不能用公式模式的理由。建议改成**按"是不是 Jinja 控制流"判断**，
并加一句："生成目标语言自己的 `if/else`（GeniE / C / Python 的语句）只是字面文本，
不影响公式模式"。这一条差异直接决定了能不能用上 0.3.0 的主要能力。

## B. 缺口：公式模式下"算出来的值"此前没有任何东西验证

`docs/template_guide.md` §14.3 如实写了："`check` 在公式模式下比的是**公式是否与当前
YAML 一致**，不是公式算出来的值"。于是有这么个空档：

* 工具侧：`check` 只比公式文本；
* 用户侧：打开 Excel 看到的代码，谁也没验过它等于 `--outdir` 导出的那份。

本目录补了一个 `verify_excel_engine.py`：实现公式子集的小求值器
（`&` 拼接 / `IF` / `ISBLANK` / `TEXT` / `INDEX`+`MATCH` / 相对列引用），
把 `Output` 表每格公式算成文本，再与 `render_all()` 的结果逐行比对。实测：

```
[公式] abs_fpi_external.xlsx        6 项, 0 项失败
[公式] abs_fpi_internal.xlsx        8 项, 0 项失败
[公式] ABS_FPI_load_cases.xlsx     14 项, 0 项失败
```

带 `--change L=400`（改一格参数、**不重跑 render**、重新求值）也是 0 失败 ——
"改参数自动重算"这句话本身被验证了。

**建议**：把这件事做进工具（例如 `check --values`，或 `render` 时顺带自检一遍），
至少在 §14.3 加一句"想知道公式算出来的值对不对，需要外部求值器；本仓库没有"。
现在的措辞容易让人以为 `check` 通过就等于代码对。

## C. 新观察：改*结构*的两面性，文档没写

§14.3 只说了"改**模板**要重跑"。改**结构**是另一回事，实测两面
（`probes/probe_formula_structure.py`）：

| 动作 | 实测 | 为什么 |
|---|---|---|
| 插 / 删**变量行** | ✅ `verify` exit 0，`check` exit 0 | 公式用 `INDEX/MATCH` **按变量名**定位 |
| 插一个 **Case 列** | ❌ `verify` exit 1（10/16 失败），`check` exit 1 | 公式里的列标是绝对字母，工况移位后就指错 |

`check` 能发现结构变化（第 4 层证据里它的报错是 `Code INT 第 B 列 ... 第 1 行不同`）。
**建议**：§14.3 补一句"增删 Case 列也要重跑 `--write-excel`；
插删变量行不用（按名字定位）"。

## D. 新观察：`case_filter` + 空 Case 列会悄悄改变归属

在 Local 表里新加一列、`kind` 留空时，`kind` 回落 YAML `default`（本例是 `EXT`），
于是这一列**自动进了外压规则集**，而表上看起来只是一个空列。
实测里 `check` 报了 `Code EXT 第 E 列（EXT-T15-mu90-kf+1）：行数不同：表里 0 行，应为 91`
—— 这一类"新列悄悄进了某个规则集"的行为，建议在 §9.2 的 `case_filter` 说明里点明，
并建议用户新增列时先把 `kind` 填上。

## E. 新观察：公式模式下 `check` 的差异信息贴的是公式片段

公式模式不一致时，`check` 打印的是**公式文本**开头：

```
ERROR 输出表已过期，共 16 处不一致：
    Code INT 第 B 列（WBT6-d8-mu90）：第 1 行不同：表里 '="// "&\'Local
```

对用户没什么可读性（他看不到"哪一行模板错了"）。建议公式模式下换成
"公式与当前 YAML 不一致（模板 <名> 第 N 行）"，或者至少截断到人能看懂的长度。
快照模式下贴文本是好的，公式模式下贴公式不是。

## F. 新观察：公式长度与"一行几个占位符"的经验值

§14.3 提了 8192 字符/格上限。实测本模板：

```
abs_fpi_external.xlsx      公式格  342 个，最长 1750 字符
abs_fpi_internal.xlsx      公式格  616 个，最长 2702 字符
ABS_FPI_load_cases.xlsx    公式格  958 个，最长 2702 字符（汇总表一行 6 个占位符）
```

一个 `{{ x }}` 展开约 300–400 字符（3 列 × `IF(ISBLANK(...))`）。
所以经验值是：**一行 8 个以上占位符就要留意**（约 3000 字符）。
建议文档给这个量级，比只给 8192 更有用。

## G. 功能请求：模板库需要 `extends` / `!include`

工具是"一个 YAML = 一本工作簿"，没有跨文件复用。项目要在一本工作簿里放几套规范、
共用一张 Global 表，就得把变量手抄一遍 —— 那是漂移的温床。
本目录用 `compose.py` 绕过（把规则集 YAML 合并成项目 YAML，并检查同名变量的
`prefix`/`suffix`/`type` 是否一致），但这本该是工具的事。

建议：给根节点加一个可选的 `extends: [a.yaml, b.yaml]`（或 `!include`），
合并语义与 `compose.py` 相同（严格字段冲突报错，宽松字段告警）。
`compose.py` 可以直接当参考实现。

## H. 本轮新增/更新的产物

| 文件 | 说明 |
|---|---|
| `abs_fpi_external.yaml` / `abs_fpi_internal.yaml` | 重写为 0.3.0 版式 + `engine: excel`（原文件是 0.1.0 版式） |
| `templates/*.j2` | 前导改为纯替换；单位改用 `suffix` |
| `compose.yaml` / `compose.py` | 规则集 → 项目工作簿 YAML（合并 + 冲突检查） |
| `abs_fpi.yaml` / `ABS_FPI_load_cases.xlsx` | 合并后的项目工作簿（两套规则共用 Global） |
| `verify_excel_engine.py` | 公式求值器（补 §14.3 的缺口） |
| `build.py` | 一条命令：compose + 生成 + 导出 + 两层验证 |
| `TEMPLATES.md` | 模板数据库索引 + 加规范的清单 |
| `probes/probe_genie_shape.yaml` | 探针 M：字面 if/else 链能不能进公式模式 |
| `probes/probe_formula_structure.py` | 探针 N：改结构（插变量行 / 插 Case 列）会怎样 |

---

# 0.4.0 回应：本轮 A–H 的处置（修复会话）

> 修复日期 2026-09-24 ｜ 被测：`excel_codegen 0.4.0`
> 复测：`pytest` 86 项全过；`python build.py` 的公式求值 28 项 0 失败；
> `node compare_with_rules.js` 21 项 0 失败；三个工作簿 `check`（含值校验）exit 0。

| 条目 | 处置 | 落点 / 证据 |
| --- | --- | --- |
| **A. 错误判断** | ✅ **已更正** | `docs/template_guide.md` §14.6 + 根 `README.md`：判断标准改成"有没有 `{%`"，并写明"生成目标语言自己的 `if/else` 只是字面文本"。上节末尾也加了指向本节的更正提示。 |
| **B. 公式值没人验** | ✅ **已做进工具** | 新增 `excel_codegen/formula_eval.py`（`&` / `IF` / `ISBLANK` / `TEXT` / `INDEX`+`MATCH` / 相对列）；`check` **默认**在公式模式下把 Output 表里的公式算一遍，与 `render_all` 逐行比对（`--no-values` 可关，参数单元格本身是公式时自动降级并提示）。`abs_fpi/verify_excel_engine.py` 仍然保留 —— 它是项目侧的独立回归工具，现在是**第二重**证据。 |
| **C. 改结构的两面性** | ✅ **已写进文档** | §14.5 新增表格：改参数 / 插删变量行 / 改前后缀 → 不用重跑；增删移动 Case 列 / 改 YAML / 改 `case_filter` → 要重跑。差异信息也改成人读的形态。 |
| **D. 空 Case 列悄悄改归属** | ✅ **已加告警** | `read_cases` 记录 `explicit_values`，`render` / `check` 对"整列都空的 Case"提示它可能悄悄落进某个 `case_filter` 规则集，并点名归属变量。 |
| **E. 公式模式下贴公式片段** | ✅ **已改** | 公式模式不一致时贴**模板第几行 + 该行原文**，并对长文本做截断（`_clip`）；`check` 的输出现在能直接指向"哪一行模板错了"。 |
| **F. 公式长度经验值** | ✅ **已加告警** | `formula.LONG_FORMULA_WARN = 3000`；`validate` 与 `render --write-excel` 都会警告"一行 8 个以上占位符就该拆行"，并给出实测到的长度。 |
| **G. `extends` / `!include`** | ⏸ **本轮不做（明确记档）** | 合并语义（严格字段报错 / 宽松字段告警）在 `compose.py` 里已经跑通，但把它做进 `load_config` 会改到"一个 YAML = 一本工作簿"的核心约定、`source_dir` 解析与 `template_file` 相对路径，属于产品级改动。已在根 `README.md` 的"未做 / 路线图"里列出，`compose.py` 继续作为参考实现。 |
| **H. 本轮产物** | — | 未改动（工作簿、模板、compose、TEMPLATES.md 原样）。 |

修复过程中另外发现并修掉的一处：公式模式下 `check` 的差异信息以前会把
"公式文本不一致"和"公式算出来的值不一致"混在一起报（同一条问题报两遍），
现在分成两条独立证据：**公式文本不一致** → 重跑 `--write-excel`；
**文本一致但求值结果不同** → 说明参数表结构变了或公式生成有 bug，单独报出来。

---

# 0.5.0 回应：派生参数（`derived:`）与数值形态统一

> 修复日期 2026-09-24 ｜ 被测：`excel_codegen 0.5.0`
> 需求（来自本项目的移植经验）：中间参数（"先算好、表里看得到、模板直接引用"）此前
> 只能写死在 Jinja 模板里；希望参数之间能有引用关系 —— 仅限**同 Case 的其他参数**与**全局参数**。

| 变化 | 说明 |
| --- | --- |
| **新能力 `derived:`** | 变量可以不填值，写一段表达式引用同 Case 的 local 与 global：`derived: "max(draft - z, 0)"`。支持链式、按 `type` 转换、循环/越界/未定义引用都在配置阶段报错 |
| **参数表里是活公式** | 能翻译成 Excel 公式的就写公式（`INDEX/MATCH` 按变量名定位 + `ISBLANK` 回落默认值），**改输入自动重算**；不可翻译的写 Python 算好的值并告警 |
| **故意不翻译的写法** | `round` / `ceil` / `floor` / `//`：Python 与 Excel 语义不同，翻过去会让"Excel 里看到的"与"导出文件"不一致 —— 宁可降级成"写值 + 告警" |
| **数值形态统一（会影响本目录）** | 公式里**不再用 `TEXT()`**：`TEXT(20.559,"0.###############")` 会打出 `20.559000000000001`。现在两边都按 **15 位有效数字**（Excel General / Python `utils.to_text` 的 `%.15g`） |
| **同步改了本目录的 `verify_excel_engine.py`** | 它的 `as_text` 原来用 `str()`（会出现同样的二进制尾巴），现在直接复用工具侧 `to_text`，与本目录工作簿的公式语义保持一致 |
| **它的局限（用到 derived 时要注意）** | `verify_excel_engine.py` 的解析器只覆盖拼接 / `IF` / `ISBLANK` / `TEXT` / `INDEX`+`MATCH`，**不含算术**。本目录现在没有派生参数，所以照旧全过；一旦用上 `derived:`，它会对含算术的公式报"不能识别的公式片段" —— 那时要么给它补算术文法，要么用工具自带的 `excel-codegen check --values`（0.4.0 起已含算术、且会递归求值派生格） |
| **本目录的产物未改动** | 三本工作簿 / 模板 / `compose.py` / `compare_with_rules.js` 都没动；复测：`build.py` 公式求值 28 项 0 失败、`check` 全过、`node compare_with_rules.js` 21 项 0 失败 |

**为什么 `round` 要故意不翻译**（值得记一笔）：`round(2.5)` 在 Python 里是 2（银行家舍入），
在 Excel 的 `ROUND()` 里是 3。如果把它翻译过去，"Excel 里看到的代码"与"`--outdir` 导出的代码"
就会差一个数 —— 正是本报告 #2 / #4 那类"静默不一致"。所以派生参数遇到 `round` 时降级为
"写入 Python 算好的值 + 告警（改输入后要重跑 --write-excel）"。要精确控制舍入，就在模板里用 Jinja。


