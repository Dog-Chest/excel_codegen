# abs_fpi —— 现场脚本与实测报告（开发侧）

这个目录**不随包发布**（`MANIFEST.in` 里 `prune abs_fpi`）：它是"拿这个工具干真活"时
留下的**脚本、探针与逐轮实测报告**。

**示例资产与用户文档已经搬到包里**：

| 想要什么 | 去哪 |
|---|---|
| 三本工作簿 / 规则集 YAML / 模板（用户向） | [`../excel_codegen/examples/abs_fpi/`](../excel_codegen/examples/abs_fpi/) |
| 怎么用、公式模式下什么免重跑、已知边界 | [那份示例的 README](../excel_codegen/examples/abs_fpi/README.md) |
| **对 `excel_codegen` 的实测报告**（逐轮发现的问题与修复复测） | [`FINDINGS.md`](FINDINGS.md) |
| 模板库索引、怎么加下一个规范（DNV / BV …） | [`TEMPLATES.md`](TEMPLATES.md) |

> ⚠️ **规范版权与免责**：本目录及其产出的示例依据 **ABS《Rules for Building and
> Classing》** 实现。**规范文本版权归 American Bureau of Shipping 所有**；本仓库只含演示
> 所必需的公式与系数，**不含**规范原文、表格或图表。使用者须**自行取得正式规范**并以其为
> 准；本实现**未经 ABS 审核或认可**，本项目与 ABS **无关联**，`ABS` 是其商标。
> 输出**不能替代**船级社审查；用于实际设计 / 建造 / 送审前必须由具备资格的人员独立复核。
> 完整声明见仓库根的 [`NOTICE.md`](../NOTICE.md)。

---

## 这里的脚本做什么

这些脚本按**绝对路径**指向 `../excel_codegen/examples/abs_fpi/` 里的示例资产
（`build.py` / `compose.py` / 两个探针里都有 `EXAMPLES` 常量），所以仓库结构变了不会指错。

```bash
python build.py            # compose → 建骨架（缺哪个建哪个）→ 渲染写回 → 导出 → 两层验证
python build.py --check    # 只验证，不动文件
python build.py --init     # 骨架不存在时也重建（会清掉你填的参数）

python probes/run_probes.py            # 0.1.0 那批最小复现（A–H）
python probes/probe_group_table.py     # 探针 G：成员表（第三层作用域）
python probes/probe_formula_structure.py  # 探针 N：改结构会怎样

node compare_with_rules.js             # 第 4 层：与外部 GeniE/Rules 逐行比对
```

`build.py --check` **只写 Output 表、绝不碰**你在 Global / Local 里填的值；
首次跑用 `--init`。

## 校验链（四层）

| 层 | 抓什么 |
|---|---|
| `validate` | YAML 结构、模板语法、变量未定义 / 定义了没人用、`output_sheet` 未声明、公式模式的越界写法 |
| `check` | ① 表里的公式 / 快照是否与当前 YAML + 参数一致；② **把公式在 Python 里算一遍**与 Python 渲染逐行比对 |
| **`verify_excel_engine.py`** | **同一个断言，但实现独立**（项目自己写的另一套公式解析器）：两条独立实现都过才算证据 |
| `compare_with_rules.js` | 与 `GeniE/Rules` 那套带量纲检查的产物逐行一致 |

第 3 层与工具自带的 `check` 都保留，因为它们**实现独立**（`formula_eval.py` vs 本项目自己的
解析器）——同一个 bug 不太可能在两条独立实现里同时出现，所以两条都过才算
"Excel 里看到的"与"`--outdir` 导出的"确实是同一段代码。

> ⚠ 第 4 层需要**本仓库之外**的同级目录 `../../GeniE/Rules`（独立仓库）。
> 该目录不存在时脚本会打印 `SKIP` 并以退出码 **2** 结束（加 `--allow-missing` 则按跳过处理、
> 退出码 0）—— 明确跳过，**不会**把"没验证"伪装成"通过"。
> 其余三层不依赖外部仓库，可照常跑。

## 文件

| 文件 | 作用 |
|---|---|
| `build.py` | 一条命令：compose + 生成 + 导出 + 两层验证 |
| `compose.py` | 合并规则集 → 项目工作簿 YAML（含冲突检查、`case_filter` 注入） |
| `verify_excel_engine.py` | 独立实现的公式求值器（第 3 层） |
| `fill_cases.py` | 样例工况取值（只有首次 `--init` 用） |
| `compare_with_rules.js` | 与 `GeniE/Rules` 逐行比对（第 4 层，需外部仓库） |
| `probes/` | 对这一版工具的最小复现与探针（`_probe_run.txt` 是运行记录） |
| `FINDINGS.md` | 对 `excel_codegen` 的实测报告（逐轮） |
| `TEMPLATES.md` | 模板数据库索引 + 加新规范的清单 |
