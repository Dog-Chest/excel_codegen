# 变更历史（CHANGELOG）

本项目遵循"每个版本对应一次真实测试驱动"的节奏：0.1.0 落地 → 0.2.0 修实测报告 →
0.3.0 加公式模式 → 0.4.0 补齐公式模式的验证链 → 0.5.0 派生参数。逐条实测证据见
[`abs_fpi/FINDINGS.md`](abs_fpi/FINDINGS.md)。

---

## 0.5.2 — 多平台零配置：uv（2026-09-27）

**需求**：Windows 上开发、Ubuntu 上干活，两边各装一次依赖太麻烦 —— 希望"换机器 / 换系统
不用配环境"。做法是把这个仓库变成标准 uv 项目：**装了 uv 之后不需要 venv、pip、apt，
也不用管本机 Python 是哪个版本**。

| 变化 | 说明 |
| --- | --- |
| 新增 `uv.lock` | **跨平台锁文件**（27 个包 / 733 行）。同一份锁里同时有 `win_amd64`、`win32`、`macosx_11_0_arm64`、`macosx_*_x86_64`、`manylinux*`、`musllinux*` 的 wheel —— 三个系统装出来的版本**完全一致** |
| 新增 `.python-version` | `3.12`。uv 据此挑解释器；本机没有就**自动下载**（不需要管理员权限） |
| `pyproject.toml` 新增 `[dependency-groups] dev` | PEP 735 依赖组，是 uv 的**默认**组：`uv run pytest` 直接可用，不需要 `--extra dev` 这种参数。原 `[project.optional-dependencies] dev` 保留给 pip 用户（两边内容一致） |
| `setup.sh` 支持 uv | 有 uv → `uv sync`（跳过 ensurepip / pip / apt 检查）；没有 uv → 原 venv + pip 路径。**两条路都实测过** |
| **支持下限 3.10 → 3.11** | `requires-python` 改成 `>=3.11`：3.10 于 2026-10 结束支持。同时把 CI 矩阵改成「**底线 3.11** + 3.13」，并写清"改底线要同时改 `requires-python` / CI 矩阵 / README"的规矩。重新锁定时去掉了只为 3.10 存在的 `exceptiongroup`（26 个包） |
| 文档 | README 新增「想在三个系统上零配置跑起来：用 uv」与「Python 版本策略」两节（含 Windows PowerShell 的装法与 `UV_PROJECT_ENVIRONMENT` 用法） |

**更正 0.5.1 的一条过强结论**：上次把"外置 NTFS 盘上 `rm -rf .venv` 卡在 `D` 状态"归因于
"NTFS 写碎文件"。本轮实测：干净写入并干净删除一个 **36MB / 1080 个文件**的 venv 在 ntfs3 上
**正常且很快**。真正的触发条件是**安装写到一半被强杀** —— 一次被打断的
`pip install --target` 留下了一个读不动的 `__pycache__`，此后 `rm -rf` / `rmdir` 全部停在
`D`（不可中断睡眠），只能重启清除。结论改为："环境放原生盘更省心，并且**不要在装依赖时中断**"。

**验证**：`uv run pytest` 107 项全过；**Python 3.11 / 3.12 / 3.13 / 3.14 各跑一遍都是 107 项全过**；
`uv lock --check` 通过；`setup.sh` 的 uv 路与 pip 路各跑通一次；`abs_fpi/build.py --check`
公式求值 28 项 0 失败。**工具行为零改动**（`excel_codegen/` 只动了版本号字符串）。

---

## 0.5.1 — Ubuntu 迁移：环境与文档（2026-09-27）

**背景**：仓库从 Windows 整目录拷到 Ubuntu（`.venv/` 里还留着 Windows 的绝对路径）。
代码本身是跨平台的（无盘符、无 Windows API、全部相对路径），坏的是"环境类"的东西。
本轮**不改工具行为**，只修环境与文档；复测：`pytest` 107 项全过，
`abs_fpi/build.py --check` 公式求值 28 项 0 失败、三本工作簿 `check` 全过。

| 变化 | 说明 |
| --- | --- |
| 新增 `setup.sh` | 一键建 venv + 装依赖 + 跑测试。**先探测仓库所在文件系统**：NTFS / exFAT 上自动把 venv 放到 `~/.venvs/excel_codegen`（`EXCEL_CODEGEN_VENV` 可覆盖），原生文件系统上照常用仓库内 `.venv/` |
| 新增 `.gitattributes` | `* text=auto eol=lf`：本仓库有 Windows 时期生成的 CRLF 文件，不声明的话与 Linux 侧同一文件会整文件 diff；`*.xlsx` 等标为 `binary`，生成物标 `linguist-generated` |
| README「安装」补三条 | ① **不要复用**从 Windows 拷来的 `.venv/`（`pyvenv.cfg` 里是 Windows 绝对路径）；② Ubuntu 要 `sudo apt install python3-venv`，否则 `python3 -m venv` 报 `ensurepip is not available`；③ **venv 别放 NTFS / exFAT 外置盘**——venv 与 pip 写上万个小文件，实测在 ntfs3 上 `rm -rf .venv` 会让进程停在 `D`（不可中断睡眠）状态 |
| `compare_with_rules.js` 明确跳过 | 第 4 层交叉校验依赖**仓库之外**的 `../../GeniE/Rules`。缺失时打印 `SKIP` + 原因并以退出码 **2** 结束（`--allow-missing` → 退出码 0），既不 `Cannot find module` 崩掉，也不把"没验证"伪装成"通过" |
| 文档数字对齐 | README 里"75 项 / 86 项 / 107 项"统一为 **107 项**；`tests/` 目录树补上漏写的 `test_derived.py`；目录树里 `CHANGELOG.md` 的说明改成 0.1.0 → 0.5.0 |
| 过时表述更正 | `abs_fpi/build.py` 与 `verify_excel_engine.py` 里"工具自带的 `check` 在公式模式下比不了公式值"是 0.4.0 之前的说法 —— 改为"0.4.0 起 `check` 默认值校验，本脚本因此从补缺口变成**独立复核**"；`FINDINGS.md`、`probes/probe_genie_shape.yaml` 中指向已不存在的 §14.4 改为 §14.6（并注明旧编号） |

**没有改的**：`excel_codegen/` 下任何一行工具代码、`tests/`、三本工作簿与其产物。

---

## 0.5.0 — 派生参数：参数引用参数（2026-09-24）

**需求**：中间参数（"先算好、表里看得到、模板直接引用"）此前只能写死在 Jinja 模板里；
希望参数之间能有引用关系 —— 仅限**同 Case 的其他参数**与**全局参数**。

| 变化 | 说明 |
| --- | --- |
| 新增 `derived.py` + `VariableDef.derived` | 一行表达式定义一个派生参数；支持链式与类型转换；循环引用 / 越界引用（global 引用 local、引用未定义名字）明确报错 |
| 参数表里是**活公式** | 可翻译的表达式写成 Excel 公式（`INDEX/MATCH` 按变量名定位、`ISBLANK` 回落默认值），改输入自动重算；不可翻译的写入 Python 算好的值并告警 |
| 可翻译子集 | `+ - * / ** % ~`、括号、常量、`min/max/abs/int/float`、`\|abs\|int\|float\|string`、比较、条件表达式（`IF`）；`and/or/not` 仅限条件 |
| **故意不翻译** | `round` / `ceil` / `floor` / `//`：两边语义不同（银行家舍入 vs 四舍五入等），翻了就会让"Excel 里看到的"与导出的不一致 |
| 数值形态统一 | 移除公式里的 `TEXT()`（`TEXT(20.559,"0.###############")` = `20.559000000000001`），两边统一按 **15 位有效数字**（Excel General / Python `%.15g`）；求值器的 `as_text` 复用 `utils.to_text` |
| 求值器补算术 | `formula_eval` 增加 `+ - * / ^`、括号、`MIN/MAX/ABS/TRUNC/MOD`，于是"引用派生格的 Output 公式"也能离线算出来 |
| 结构 | Jinja 环境抽到 `jinja_env.py`（`renderer` 与 `derived` 共用一套语义，避免循环依赖） |
| CLI | `validate` 的变量清单新增「来源」列（填写 / 计算）并提前编译派生表达式；派生引用计入"变量被使用"（不再误报未使用） |
| 测试 | 107 项（新增 `test_derived.py` 21 项） |

---

## 0.4.0 — 公式模式的验证链（2026-09-24）

**背景**：`abs_fpi` 的第二轮复测把内外压模板改成 `engine: excel`，同时发现两件事：
① 0.3.0 文档里"有 `if/else` 就不能用公式模式"的判断**是错的**（那是生成语言的字面文本，不是 Jinja 控制流）；
② `check` 在公式模式下只比公式文本，**"公式算出来的值对不对"没人验过**。

| 变化 | 说明 |
| --- | --- |
| 新增 `formula_eval.py` | 公式求值器：解析工作簿里**真实的公式文本**（`&` / `IF` / `ISBLANK` / `TEXT` / `INDEX`+`MATCH` / 整列与单元格引用 / 相对列），算成文本 |
| `check` 默认值校验 | 公式模式：① 比公式文本；② 求值后与 Python 渲染逐行比对。`--no-values` 可关；参数单元格本身是公式时自动降级为"只比公式"并提示 |
| 差异信息给人看 | 公式模式不一致时贴**模板第几行 + 该行原文**（截断），不再贴几百字符的公式 |
| 空 Case 列告警 | `read_cases` 记录 `explicit_values`；整列都空的 Case（新插的空列）会提示"可能悄悄落进某个 `case_filter` 规则集" |
| 超长公式告警 | `formula.LONG_FORMULA_WARN = 3000`，`validate` / `render --write-excel` 都会提醒拆行（一个 `{{ x }}` ≈ 300–400 字符） |
| 文档更正 | 指南 §14.6 改成按"有没有 `{%`"判断；新增 §14.5「改结构：什么要重跑」（插删变量行 ✅ 不用 / 增删 Case 列 ❌ 要重跑） |
| 测试 | 86 项（新增 `test_formula_eval.py` 与 CLI 值校验/告警用例） |

## 0.3.0 — 公式模式：改参数免跑脚本（2026-09-24）

| 变化 | 说明 |
| --- | --- |
| 新增 `engine: excel` | 输出表写 **Excel 公式**而不是文本快照；改 Global / Local 参数后 Excel / WPS 打开即重算 |
| 支持子集 | `{{ x }}` / `.value` / `.text` / `.prefix` / `.suffix` / `{{ case_name }}` / `{{ template_name }}`；`{% %}` / 过滤器 / 算式报错并给出行内容 |
| 与 Python 侧对齐 | `INDEX/MATCH` 按变量名定位（插行不指错）、`ISBLANK` 处理空单元格回落（`INDEX` 对空格返回 `0`，不能用 `=""`）、`TEXT()` 规范化数值、横向相对列 / 纵向绝对列、`fullCalcOnLoad` |
| 新增 `formula.py` | 公式编译器；`examples/example_formula.yaml` + `examples/template_formula.xlsx` |
| `validate` / `HOWTO` | 模板清单新增"引擎"列并提前编译；HOWTO 表区分"公式·自动重算"与"快照·需重跑" |

## 0.2.0 — 实测报告修复会话（2026-09-24）

来源：`abs_fpi` 第一轮实测报告列出的 12 条问题，修掉 11 条，1 条判为能力边界。

| 变化 | 说明 |
| --- | --- |
| Prefix/Suffix 不再被 strip | `suffix: " m"` 得到 `340 m`（原来静默变 `340m`） |
| `render` 不再静默 | 摘要永远打印"写回 Excel │ 是 / 否（需要 --write-excel）"，只预览时另给 `!` 提示 |
| GBK 控制台 | 标记改 ASCII（`OK`/`ERROR`/`!`）+ stdout/stderr `errors="backslashreplace"`，成功路径不再以退出码撒谎 |
| 清理旧结果按真实底边 | 改短模板 / 减少 Case 后重渲染不再残留上一版代码 |
| 新增 `HOWTO` 表 | 三步说明 + 命令 + 快照提醒 + 输出位置，放第一张（`--no-howto` 可关） |
| 指纹 + `check` 命令 | 每次写回记录时间 / 参数指纹 / 输出指纹；`check` 判定是否过期（CI 可用） |
| 新增 `case_filter` | per-template 的 Case 过滤，一本工作簿放多套规则、共用一张 Global 表 |
| `{{ x.value }}` 形态一致 | 整数浮点规范化对 `{{ x }}` 与 `{{ x.value }}` 同时生效（Jinja `finalize`） |
| YAML 重复键报错 | 配置阶段给出"第 N 行键重复" |
| `--cases` 支持名字 | `--cases EXT-T20.559,INT-T15` |
| `Template` 表首行提示 | 明说"只读参考：模板真源是 YAML / template_file" |
| 文档 | 能力边界（没有第三层作用域）、`safe_filename` 规则、"空行各占一行"的噪音警告 |

## 0.1.0 — 首个可用版本

YAML 定义变量与模板 → 生成 Excel 表单 → 填参数 → Jinja2 渲染 → 写回 Output 表 / 导出代码文件。
Prefix+Value+Suffix（`VarValue` + `pvs` / `wrap` 过滤器）、`horizontal` / `vertical` 两种布局、
`init` / `render` / `validate` 三个命令、`tests/` 38 项用例、`docs/template_guide.md`。
