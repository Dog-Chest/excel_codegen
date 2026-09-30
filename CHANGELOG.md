# 变更历史（CHANGELOG）

本项目遵循"每个版本对应一次真实测试驱动"的节奏：0.1.0 落地 → 0.2.0 修实测报告 →
0.3.0 加公式模式 → 0.4.0 补齐公式模式的验证链 → 0.5.0 派生参数 → 0.5.x 跨平台与文档 →
0.6.0 取值约束 / `extends` / 行内 `{% if %}` / 质量护栏 → 0.7.0 校验与复用补齐 →
0.8.0 行列风格 / 成员表落地 / NASTRAN 工况控制 → 0.8.1 环境体检（`setup.sh`）→
0.9.0 内置示例随包发布（`excel-codegen examples`）。逐条实测证据见
[`abs_fpi/FINDINGS.md`](abs_fpi/FINDINGS.md)。

---

## 0.9.0 — 内置示例随包发布 + `excel-codegen examples`（2026-09-30）

这一版把"干过真活的用例"变成**软件自带的东西**：装了 pip 包、没克隆仓库的人也能直接跑。

### 1. 新增 `excel-codegen examples`

```bash
excel-codegen examples                      # 列出示例、各自演示什么
excel-codegen examples --copy ./examples    # 拷出来用（--only abs_fpi 只拷一个）
```

* 加上 `--force` 才会覆盖已有目录 —— 拷出来的文件是**用户的**，默认不碰；
* 拷出来的每个示例都带一本**已经填好样例参数**的工作簿，打开就能改；
  公式模式下改完自动重算，不需要跑命令。

### 2. 三个内置示例（`excel_codegen/examples/`）

示例从"仓库里的文件"变成**包数据**：wheel 与 sdist 里都有（`MANIFEST.in` +
`[tool.setuptools.package-data]`），`tests/test_examples_command.py` 守住"清单与目录
一一对应"以及"拷出来仍然解得出 `template_file`"。

| 示例 | 内容 |
| --- | --- |
| `basic/` | `example_formula.yaml`（默认公式模式）+ `example.yaml`（显式 snapshot：过滤器 / 循环 / 导出文件） |
| `nastran/` | 纵向布局的 NASTRAN 工况控制语句（§19 / §20） |
| `abs_fpi/` | **现场用例**：ABS FPI 内外压 → GeniE，两个规则集可各自单用、也可合成项目工作簿共用一张 Global 表；内压用成员表 `Tank Data` |

### 3. 仓库结构整理

* 示例**资产**（YAML / `.j2` 模板 / 工作簿 / 导出产物）从 `abs_fpi/` 搬到
  `excel_codegen/examples/abs_fpi/`，随包发布；仓库根的 `abs_fpi/` 只留**现场脚本与实测报告**
  （`FINDINGS.md` / `TEMPLATES.md` / `probes/` / `build.py` / `compose.py` / `verify_excel_engine.py` /
  `compare_with_rules.js`），仍然 `prune` 掉不进包；
* 现场脚本按 `EXAMPLES` 常量指向新位置，`build.py --check`（28 项）、三个探针、
  成员表探针全部复跑通过 —— 搬完没有丢证据；
* 仓库顶层不再有 `examples/`，避免"两份示例漂移"。

### 4. 文档

* 新增 [`excel_codegen/examples/README.md`](excel_codegen/examples/README.md)（示例索引 + 每个怎么跑）
  与 `abs_fpi/README.md`（现场用例怎么用、什么免重跑、已知边界）；
* `abs_fpi/README.md` 改写为**开发侧**索引（脚本 / 探针 / 四层校验链）；
* README 快速开始改成"`examples --copy` → 打开工作簿 → 要导文件才回命令行"；
* `docs/cli.md` 补第六个命令；`docs/setup.md` 新增「Windows 上怎么用」并说明 `setup.sh` 仅 POSIX。

### 真 bug：`setup.sh` 在 macOS 上第一屏就崩（bash 3.2 把中文首字节吞进变量名）

macOS 自带 **bash 3.2**。它的解析器会把紧跟在 `$VAR` 后面的多字节字符**首字节**当作
变量名的一部分，于是

```bash
warn "仓库所在文件系统是 $REPO_FS（非原生 Linux 文件系统）。"
```

在 macOS 上变成 `REPO_FS<0xEF>: unbound variable` —— `set -u` 让脚本立刻退出，
**任何 macOS 用户跑 `./setup.sh` 都在第一屏就挂**（而 `docs/setup.md` 里正是让人跑它）。
Linux 的 bash ≥ 4 解析正常，所以这个坑从 0.5.1 加入 `setup.sh` 起就只在 macOS 上炸，
本地与 ubuntu / windows 的 CI 一直全绿。

它是先以"CI 上一个 `UnicodeDecodeError`"的面目出现的：bash 把坏字节连同变量名一起写进
stderr，测试的严格 UTF-8 解码先崩了 —— 顺着那条线才挖到真正的病因。

* `setup.sh` 里 5 处 `$VAR` 紧邻中文一律改成 `${VAR}`（`${arg}` / `${PYTHON}` /
  `${REPO_FS}` ×2 / `${reason}`）：花括号只是显式界定名字，任何 bash 版本都无歧义；
* `tests/test_setup_script.py` 增加**静态守卫**：扫 `setup.sh`、随包发布的 `*_render.sh`、
  以及**新生成**的渲染脚本，凡 `$VAR` 紧邻非 ASCII 字节就失败 —— 这类坑不该等 macOS CI
  才发现（已确认去掉花括号后守卫会红）；
* 顺带把测试读子进程输出改成显式 `encoding="utf-8", errors="replace"`，断言锚在 **ASCII
  证据**上（版本号 / `--recreate` / `site-packages` / `pip` / 退出码 / `ERROR`），
  免得下次脚本一出声就被解码异常盖住真正的原因。

### 回归

349 项测试（+8）、覆盖率 87.91%、ruff / format / mypy 全过；
`uv build` 出的 wheel 与 sdist 里示例文件**逐一致**（各 64 个）、`twine check` 通过；
sdist 解包后 `pytest` 全过；`abs_fpi/build.py --check` 公式求值 28 项 0 失败、
三个工作簿 `check` 全过；探针 A–H / G / N 全部通过；
净 venv 只装 wheel 后 `examples --copy` / `check` / `render` / `init` / `*_render.sh` 全走通。

---

## 0.8.1 — 坏掉的旧 venv 必须在装依赖之前被拦住（2026-09-30）

**工具行为无变化**：改的是仓库自带的开发环境脚本 `setup.sh`，外加一组回归。
不涉及 `excel_codegen` 包本身。

### 问题（真踩过：开发机换了 Python 之后环境静默失效）

`~/.venvs/excel_codegen` 是按 Python 3.12 建的，后来系统的 `/usr/bin/python3.12`
被删、`python3` 指向 3.14。venv 里的 `bin/python3` 是**通用**链接
（`-> /usr/bin/python3`）而不是钉死版本，于是解释器悄悄漂到 3.14，而依赖还躺在
`lib/python3.12/site-packages` 里 —— 结果很反直觉：

* `bin/python -c ""` **照样返回 0**：老版体检只测这一句，于是打印"复用已有环境"，
  一路走到升级 pip；
* 可 `import jinja2` 与 `-m pip` 全都失败，脚本停在一句 `No module named pip`
  上 —— 这句话与真正的病因（解释器被换掉了）毫无关系，用户拿不到任何线索。

讽刺的是老代码那行注释写的正是"或它依赖的 Python 被换掉了"：它声明要挡这个场景，
实际挡不住。

### 改动

* 新增 `venv_problem()`，复用旧环境之前逐项体检：
  解释器能执行 → `pyvenv.cfg` 记的版本与现在真跑的对得上 →
  `site-packages` 在当前解释器的搜索路径上 → `pip` 可用。
  任一项不过就**在碰 pip 之前**停下，报出具体原因
  （"解释器被换掉了：环境是按 Python 3.12.3 建的，现在跑的是 3.14.4"）
  与 `--recreate` 的出路；文件头的说明同步补上这条体检。
* `tests/test_setup_script.py` 4 项：三种坏法（解释器漂移 / `site-packages` 不匹配 /
  没有 pip）各钉一条，外加"正常环境不许被误判"。用例在临时目录里造**真的** venv
  （`--without-pip` 建一次 + `copytree` 分发，整个文件 0.4 秒），再用桩 `pip` 模块
  让体检通过 —— 全程不联网。已确认前三条在改动前**全部失败**，失败信息里正是那句
  `No module named pip`。

### 回归

330 项测试（+4）、覆盖率 87.75%、ruff / format / mypy 全过；
`./setup.sh --recreate` 在 Python 3.14.4 上实跑通过（330 项测试 + 覆盖率门槛）。

---

## 0.8.0 — 默认公式模式 + 行列风格 + NASTRAN 工况控制（2026-09-27）

### 0. **默认引擎改成公式模式**（`engine: excel`）

本工具的用法是"**写一个 YAML → 生成一本 Excel → 之后就在 Excel 里干活，YAML 攒成模板库**"。
按这个用法，工作簿必须**独立可用**：改一格参数输出自己就变，不需要再跑命令。于是：

* `TemplateDef.engine` 的默认值 **`snapshot` → `excel`**；不写 `engine` 就是公式模式；
* `snapshot` 仍然保留，但必须**显式**声明 —— 只在两种情况需要：模板里有
  `{% for %}` / 过滤器 / 多行 `{% if %}` / `{% include %}`（公式模式表达不了），
  或者你要的是"导出的代码文件跟着参数走"；
* 报错信息补齐：公式模式表达不了的写法，一律在消息里给出"加 `engine: snapshot`"
  （`{% for %}` / `{% include %}` / 过滤器原本就有，这次补上**跨行 `{% if %}`**）；
* **建表时就拦住**：`create_template` 现在会预检每个模板（语法、`{% include %}` 片段、
  公式能否编译）。以前是等你把参数填完、跑 `render` 才炸 —— 而这个用法下工作簿只建一次，
  越早报错越好。报错时**不会**生成那个 .xlsx。

仓库内随之显式标注：`examples/example.yaml`（演示过滤器 + 循环）、
9 个 `abs_fpi/probes/probe_*.yaml`（记录的是快照模式下的行为）、若干测试夹具。
`examples/example_formula.yaml` 那两行 `engine: "excel"` 现在只是"写出来更清楚"。

**`init` 生成完就是可用的工作簿**：默认 `--prerender`，建完骨架顺手把输出写一遍 ——
公式模式下输出表里是活公式，于是"一条命令生成 → 打开 Excel 干活"，不用再跑
`render --write-excel`。参数还是默认值、暂时过不了取值约束 / `asserts` 时**跳过预填并说明原因**
（骨架照常生成，`init` 不失败）。HOWTO 表第 3 步也按引擎改写：全公式时说"改完直接看输出表，
不用跑命令"。新增 `--no-prerender` 只要骨架。

新增 `tests/test_engine_default.py` 13 项（默认值 / 默认写公式 / 显式 snapshot /
建表期报错 / 报错里给出路 / `init` 预填与 `--no-prerender` / 预填尽力而为 /
示例与默认值一致）。

### 1. 行列风格：`excel.local_direction`

原先 Local Parameter 表**只能**一个工况一列（E 列起）。有些软件里工况控制语句是按**行**
给的（NASTRAN 的 `SUBCASE` / `SUBCOM` 就是典型），从别处拿到这种数据只能手工转置。

新增 `excel.local_direction: horizontal | vertical`（默认 `horizontal`，老行为不变）：

* `vertical`：**一行一个工况** —— A 列写 Case 名（A2 起）、第 1 行 B 列起写变量名；
  下拉复制行即可加工况，遇到第一个空 A 格就停。
* 纵向布局**没有** Prefix / Suffix 列（第 1 行整行都是变量名），前后缀仍来自 YAML，
  点表头格的批注能看到。
* 与模板自己的 `direction`（**输出**排布）**完全独立**：输入竖着填、输出横着写是最常用的组合。

配套改动：

* `CaseData` 的 `column` / `row` 二选一，新增 `where`（"第 N 列" / "第 N 行"）——
  所有面向人的报错不再假设"工况是列"。
* 公式模式纵向引用：`INDEX('Local Parameter'!$A:$ZZ,$2,MATCH("port",'Local Parameter'!$1:$1,0))`；
  公式求值器跟着支持**二维 `INDEX`**、**行区间 `MATCH`**（`'sheet'!$1:$1`）与
  `A2` 这种"列标+行号挤在一个词法 token 里"的引用形态（原来的横向布局只生成 `E$1`，没暴露这个洞）。
* `check` / `doctor` 的提示语跟着变"Case 列"或"Case 行"；`check_required_sheets` 按布局校验表头。
* 指南新增 §19。

新增 `tests/test_local_direction.py` 16 项（建表 / 读值 / 两种布局对拍 / 公式值对拍 /
插行加工况 / 约束报错定位）。

### 2. abs_fpi 重做工作表：舱数据改用成员表

0.7.0 加的能力（`variables.group`，指南 §18）在真实项目里落地：

* `abs_fpi_internal.yaml` 的 13 个舱参数从 `local` 挪进 `variables.group`
  （工作表 `Tank Data`，一行一个舱：WBT6 / WBT7 / COT1）；`local.tank_ref` 变成
  **指针**（带 `choices` 数据有效性）。
* `compose.py` 学会合并成员表声明：`sheet` / `key` 必须一致，成员行与变量按名字合并
  （一本工作簿只有一张成员表）。
* `fill_cases.py` 改成"逐工况只写 `tank_ref` + 逐工况参数，舱数据写进 Tank Data"。
* `abs_fpi_external.xlsx` 不受影响；两本内压相关的工作簿**重建**（多了 `Tank Data` 表）。

迁移的正确性证据：

| 断言 | 结果 |
| --- | --- |
| 生成的 `.js` 与迁移前**逐字节相同** | ✅ 只有汇总 `.md` 的"出处"一列改了话术 |
| 改一次 WBT6 的 `l_tank` → 两个 WBT6 工况一起变，WBT7 / COT1 不动 | ✅ `probes/probe_group_table.py` |
| 改一个 Case 的 `tank_ref` → 整组舱数据换掉 | ✅ 同上 |
| `tank_ref` 指向不存在的成员 → 明确报错（`choices` 先拦；去掉 `choices` 是成员表查找拦） | ✅ 两种报错都点名成员 + 列可选值 |
| 公式模式下 Case → 舱 → 变量的两级 `INDEX/MATCH` 链条 | ✅ 工具自带 `check` 会整条算一遍；28 项公式求值 0 失败 |

新增 `abs_fpi/probes/probe_group_table.py`（5 步实测）**与 `tests/test_abs_fpi_group.py` 9 项回归**
（把这次迁移钉进 CI：配置层面舱参数不在 local 里、数据层面两个工况共用一份舱数据、
产物层面 `abs_fpi/generated/` 与当前 YAML + 工作簿一致 —— 免得改了 YAML 忘了重新生成）。
`abs_fpi/README.md`、`TEMPLATES.md`、`FINDINGS.md` #3(a) 从"尚未迁移的能力边界"改成"已迁移 + 证据"。

### 3. 生成测试：NASTRAN 工况控制语句

`examples/nastran_case_control.yaml` —— 用 §19 的纵向布局填工况控制语句（一个子工况一行，
整块可粘贴），生成 `SUBCASE` / `SUBCOM` 块。语句留空就不输出（`SUBSEQ` 的内容原样输出，
工具不解释"哪几个子工况、各乘多少"）。

* **两个模板都是 `engine: excel`（公式模式）**：在 Excel 里改一格参数，Code 表里的语句
  立刻跟着变，不用跑命令 —— 这才是这个用例的意义（快照模式得重跑 `render`）；
* 每个子工况导出一个 `.inc`，文件名按只用于排序的 `seq` 编号（`cc_01_LC1.inc`），
  `sed '/^$/d' deck/cc_*.inc > case_control.deck` 就是完整的 case control 段；
* 用 `asserts` 拦住"填错地方"：`SUBCOM` 必须给 `SUBSEQ`、`SUBSEQ` 不许写在 `SUBCASE` 行上；
* 指南新增 §20：公式模式与快照模式"可选行"的两种写法与取舍
  （公式模式一行一条语句、空语句=空格子；快照模式把换行写进 `{% if %}` 里面，逐字节干净）。

顺带修掉三个在这次实测里暴露的问题：

| 问题 | 修法 |
| --- | --- |
| **`filename` 里只能写 `{{ case_name }}` / `{{ template_name }}`** —— 用别的变量（比如只用来排序的 `seq`）会在导出时炸 | `RenderResult` 带上渲染上下文，`export_files` 用它渲染文件名；现在文件名里可以用**这个 Case 的任意参数** |
| **只写在 `filename` 里的变量被判成"定义了没人用"** | `collect_variables` 也解析 `filename`，于是这个假警告消失，文件名引用到不存在的变量也会在 `validate` 阶段就报出来 |
| **求值器读不了绝对行号** `INDEX(...!$A:$ZZ,$2,...)` —— 只有"输入纵向 + 输出纵向"才会生成它，`check --values` 直接抛 `看不懂的记号 punct='$'` | `formula_eval.term()` 把 `$<数字>` 当数字 |

### 4. 顺手修掉三个"悄悄算错 / 装不上"的问题

| 问题 | 影响 | 修法 |
| --- | --- | --- |
| **`evaluate_template_values` 按"第几个 Case"读输出列** | 传 Case 子集时**不报错、直接取到别的工况的值**（对拍测试因此可能假通过 —— 本仓库真有一个测试是这么蒙对的） | 保留"按输出顺序"的位置语义（`case_filter` 会跳工况，位置才是对的），但当表头写着名字时**逐列核对**，对不上就报"输出表与参数对不上（传了子集？表是旧的？请重跑 `--write-excel`）" |
| **`setup.sh` 的 Python 下限写 3.10**，而 `pyproject.toml` 要求 3.11 | 3.10 的机器上脚本放行，一路装到 `pip` 才报错 | 下限改成 3.11 并由脚本自己拦 |
| **`setup.sh` 复用已有环境前不验证** | 环境半坏（上次装到一半 / 依赖的 Python 被换掉）时报的是 pip 的 `Errno 13 权限不够`，看不出该干什么 | 复用前先 `python -c ""` 试一下，坏了就明确让 `--recreate`；`pip install` 失败时列出两种常见原因与下一步。另修掉文件系统探测的 `UNKNOWN*` 大小写（`case` 分支永远匹配不上小写的 `unknown`） |

还有一处**测试自身的错误**：`test_shipped_example_matches_the_test_fixture` 把示例工作簿
的取值写死了，而那个工作簿是**给人改的** —— 一改就红。现在示例相关的测试只钉"形态"
（纵向 + 公式模式）与"自洽"（形状、公式算出来 == Python 渲染、`.deck` 与工作簿一致），
不再钉具体数值。

新增 `tests/test_nastran_case_control.py` 17 项 + `tests/test_renderer.py` 2 项 +
`tests/test_formula_eval.py` 1 项。
`examples/generated_nastran/` 是示例工作簿的实测产物（含拼好的 `case_control.deck`）。
**测试 268 → 326 项。**

---

## 0.7.0 — 发布前打磨：校验 / 复用 / 体检 / 第三层作用域（2026-09-27）

发布 PyPI 之前的一轮打磨，8 个提交（0 必修 → ③ 批注 → ① 脚本 → ④ check 报告 →
② asserts → ⑤ include → ⑦ doctor → ⑧ 成员表）。**测试从 195 项涨到 268 项**，
其中 sdist 打包那个问题是**真会出丑**的（发布出去的源码包里测试跑不起来）。

### 0. 发布前必修

| 问题 | 证据与修法 |
| --- | --- |
| **sdist 里的测试是坏的** | setuptools 默认只收 `test*.py`，把 `tests/conftest.py` 漏了 —— 解包后 `pytest` 全部 ERROR（`fixture 'config_text' not found`）；`examples/` 与 `docs/` 也不在 sdist 里，而 README 的快速开始指着 `examples/example.yaml`。新增 `MANIFEST.in`（测试带全 + examples/docs 带上 + `prune abs_fpi`）。验证：解包后 195 项全过、示例配置 validate 通过 |
| **HOWTO 表印着过时的话** | 「公式只做占位符替换」—— 0.6.0 之后公式模式支持行内 `{% if %}` 了，这句会直接印到用户的工作簿上 |
| **`Documentation` 是占位符** | 填上真地址，并补 `Repository` / `Changelog` / `Issues` |
| **两个公式求值器让人困惑** | 给 `abs_fpi/verify_excel_engine.py` 的 docstring 加开头警告：它是项目侧的**独立复核**工具；使用者要找的是 `check --values`（文法更全） |

**CI 加了一条守卫**：每次推送都构建 sdist、解包、跑测试、校验示例配置 —— 上面第 1 条坏过一次，不能坏第二次。

### 3. Excel 批注：把"这一格填什么"写在变量名上

`VariableDef` 新增 `unit` 字段（纯文档）；`init` 给**变量名那一格（A 列）**加批注，
内容是：描述 / 填写位置 / 单位 / 类型 / 约束 / 前缀后缀 / 默认值 / **模板里怎么引用**
（`{{ draft }}`）；派生参数的批注改说「自动计算：<表达式>，不用手填」。

只加在名字格而不是每个取值格：取值格已经有数据有效性的输入提示，而 A 列是冻结的、
永远可见。`--comments/--no-comments` 可关。

新增 `tests/test_comments.py` 7 项。

### 1. 一键刷新脚本：`<工作簿名>_render.bat` / `.sh`

`init` 在**工作簿旁边**生成两个脚本（`--no-scripts` 可关）：

* `cd` 到脚本自己所在目录，所以工作簿/配置放在哪儿都行；
* 优先 `uv run excel-codegen`，**没装 uv 就退回** PATH 里的 `excel-codegen`；
* 出错时打印常见原因（依赖没装 / Excel 正开着文件）并**以非 0 退出**，批处理与 CI 里也能用。

两个平台都生成（一本工作簿常在 Windows 与 Linux 之间传），`.bat` 用 CRLF + 反斜杠，
`.sh` 带可执行位 + 正斜杠。HOWTO 表会指过去（"懒得开终端就双击…"）。

`ProjectConfig` 顺带记住 `config_path`（脚本要算相对路径）；`write_results` 刷新 HOWTO
时会**现查**脚本是否还在旁边，所以重新加载配置后那句话不会凭空消失。

新增 `tests/test_run_scripts.py` 8 项 —— 其中两项**真的执行**生成的 `.sh`（用桩程序
冒充 `excel-codegen`），确认 cd、参数转发与退出码都对。

### 4. `check` 报全部差异 + `--json`

* 差异原来每个 Case 只报**第一处**；现在**按行列出**（每个 Case 最多 5 行，其余折成
  一句"还有 N 行不同"）。CI 里定位问题不用再"改一处、跑一次"。
* 新增 `--json`：stdout 上**只有** JSON（表格与提示都跳过），结构是
  `ok` / `config` / `excel` / `recorded` / `current` / `drift` / `warnings` / `problems`，
  退出码与表格模式一致。适合 `jq` 与看板消费。

新增 `tests/test_check_report.py` 9 项（差异汇总的 4 种形态 + CLI 表格模式 + `--json` 的
新鲜 / 过期 / 参数漂移 / 无渲染记录）。

### 2. 跨变量校验：根级 `asserts`

单变量约束只能看一列；工程上真正容易出事的是**组合**（吃水超过型深、内压工况带真空压力）。

```yaml
asserts:
  - "draft.value <= d_tank.value"
  - "not (kind.value == 'INT' and p_vp.value > 0)"
```

* 对**每个 Case** 求值，为假就报错并指名道姓（哪张表、哪一列、哪个 Case、哪条规则）；
* 复用 `case_filter` 的表达式机制（`renderer.compile_asserts` / `check_asserts`），
  所以 `.value` 语义与报错风格完全一致：裸变量是组合值，**数值比较必须写 `.value`**；
* `render` / `validate -x` / `check` 都会查；不带 `--excel` 的 `validate` 只编译（查语法）；
* `extends` 进来的文件里的 `asserts` 全部保留（先被 extends 的在前）。

`validate` 会把 assert 清单一并打出来，方便一眼看到"这本工作簿受哪些规则约束"。

新增 `tests/test_asserts.py` 15 项（配置期 5、求值 6、与 extends 的关系 1、端到端 3）。

### 5. 模板片段复用：`{% include %}`（快照模式）

`extends` 复用"变量与模板定义"；`{% include %}` 复用"模板里的一段代码"。

```jinja
// ===== {{ case_name }} =====
{% include "frag/constants.j2" %}
```

* 路径**相对声明模板的那个文件**（与 `template_file` 同一套规则）—— `extends` 进来的
  模板也能找到自己旁边的片段；
* 可嵌套；循环 include 会被跳过而不是无限递归；
* 片段里的变量**同样算"被引用"**：`validate` 不再误报"定义了没人用"，也会查它们有没有定义；
  改了片段 → `check` 会判定过期（片段参与渲染与指纹）；
* **公式模式明确报错**：片段会让"一个 Case 占几行"不可预测，而公式模式是一行一个单元格。

实现上给 `build_environment` 加了 `search_path`，并新增 `renderer.template_environment()` ——
每个模板用"声明它的那个目录"的环境（`render_all` 里按目录缓存）。

顺带把"`Local Parameter` 没有定义任何局部变量"的报错说清楚：本工具用「局部变量 × Case 列」
定位工况，至少要有一个局部变量。

新增 `tests/test_include.py` 10 项；指南新增 §17，§13 的能力边界条目同步更正。

### 7. `excel-codegen doctor`：一条命令体检

体检**环境 / 配置 / 工作簿**三层，把散在各处的常见坑一次说清；**有 ERROR 时退出码 1**。

| 层 | 检查项 |
| --- | --- |
| 环境 | Python 是否 ≥ `requires-python`、六个运行时依赖、uv / git、是不是虚拟环境 |
| 配置 | YAML 加载（含 `extends`）、模板与 `{% include %}` 片段、派生参数、`asserts` 语法、**一条约束都没有**、定义了却没人用的变量 |
| 工作簿 | 工作表与 Case 列、取值约束与 `asserts` 是否满足、**整列都空的 Case**、指纹是否一致、一键脚本 |

典型用途："我这儿跑不起来"的第一条命令、发布前自检、定期确认这本工作簿受哪些规则管着。

写它时踩到并修掉的两件事：① 用 `defined_names` 算"未使用的变量"会把保留名
`case_name` / `template_name` 也算进去（那是"模板里能引用的名字"，不是"定义了没人用"）；
② 参数违反 `asserts` 时 `render_all` 会抛异常，得接住变成一条 ERROR 而不是甩 traceback。

新增 `tests/test_doctor.py` 9 项。

### 8. 第三层作用域：成员表（船 → 工况 → 舱/设备）

**这是 `abs_fpi` 那条线上最大的痛点**：一个舱（WBT6）被多个工况引用，只有 global / local
两层时它的 13 个参数只能**按工况摊平**（同名舱在每个 Case 列里各写一遍，改一个舱要改 N 列）。
`FINDINGS.md` 把它记为"能力边界"，现在补上了。

```yaml
variables:
  local:
    - name: tank_ref
      choices: ["WBT6", "WBT7"]
  group:
    sheet: "Tank Data"
    key: tank_ref              # 哪个 local 变量指向成员（必须是 local）
    members: ["WBT6", "WBT7"]  # init 时先建这几行；之后插行即可加成员
    variables:                 # 舱自己的参数，只写一遍
      - name: l_tank
        type: float
        min: 0
```

| 设计点 | 选择与理由 |
| --- | --- |
| **布局** | 与 Global / Local 都不同：**一行一个成员，B 列起一个变量一列**。工程师写舱容表就是这个样子，而且这样公式模式能复用**同一形态**的一维 `INDEX/MATCH` |
| **优先级** | 成员 < 全局 < 局部（成员描述"这是什么"，global/local 描述"怎么算"）；三类变量**不许重名**（配置期报错） |
| **成员从表里发现** | `members` 只是 init 时先建哪几行；之后在表里插一行就多一个成员，不用改 YAML |
| **公式模式** | 支持：`INDEX('Tank Data'!$B:$B, MATCH(<本 Case 的 tank_ref>, 'Tank Data'!$A:$A, 0))`，嵌套一次一维查找；`check` 会把整条链算一遍再比对 |
| **约束** | 成员表变量支持 `min` / `max` / `choices` / `pattern`（写成数据有效性），也支持批注 |
| **不支持** | 成员变量不能用 `derived`（派生依赖图只覆盖 global / local）；再深一层（舱里分部件）表达不了 —— 都写进了 §18.4 与 §13 |
| **报错** | Case 的 key 指向不存在的成员时，列出可选成员；key 为空时说清它该指向谁 |

顺带修的一处：`asserts` 的上下文原来没带成员值 —— 写 `asserts: ["h_tank.value <= 32"]`
会报 undefined。现在 `asserts` 与 `case_filter` 都看得到成员值。

新增 `tests/test_group.py` 15 项；指南新增 §18，§13 的能力边界条目同步更正。

---

## 0.6.0 — 取值约束 / `extends` / 行内 `{% if %}` / 质量护栏（2026-09-27）

四个功能 + 一道质量闸，共 4 个提交（① 取值约束 → ③ `extends` → ④ 行内 `{% if %}` →
⑤ 质量护栏），另有一个发布流程。**测试从 107 项涨到 195 项**，其中两个是 lint / 类型检查
当场揪出来的真 bug（见 §5）。

### 1. 取值约束：`min` / `max` / `choices` / `pattern`

**需求**：工具此前只查"变量有没有定义、类型对不对"，**完全不看值** —— 把吃水 `20.559`
手滑打成 `205.59`，模板照样渲染，`check` 也只会说"与参数一致"。这是唯一一类
"工具不报错、但交付物是错的"的缺口，正好撞在项目"不许静默"的原则上。

| 变化 | 说明 |
| --- | --- |
| `VariableDef` 新增 `min` / `max` / `choices` / `pattern` | 配置期就拦住写错的约束：`min > max`、给 `type: string` 加范围、`default` 越界、给派生参数加约束、空 / 重复 / 非列表的 `choices`、非法正则 —— 全部在 `load_config` 阶段报错 |
| 新增 `check_value_constraints()` | `render` / `validate -x` / `check` 每次读表后都校验一遍，**一次列出全部问题**并指名道姓（哪张表、哪一列、哪个 Case、哪个变量）。越界 → 退出码 1 |
| Excel 数据有效性 | `create_template` 把约束写成工作簿的 DataValidation：`choices` → 下拉列表，`min` / `max` → 数值范围（`type: int` 用整数校验），并带上悬停提示。**判据仍是工具侧那道闸** —— Excel 的校验挡不住粘贴、脚本写入与别人发来的老文件 |
| 空值算不合格 | 声明了约束就意味着"必须给一个合法取值"。这顺手治了"新插一个空 Case 列悄悄落进某个规则集"（§9.2 的老问题）：`kind` 有 `choices` 又没给 `default` 时，新列会在渲染时报"取值（空）不在允许列表里" |
| 表达不了就报错 | 下拉选项里含逗号、或选项拼起来超过 Excel 的 255 字符上限 → 直接报错；`pattern` 因 Excel 没有正则而**不写**数据有效性（写一个乱报错的校验不如不写），只由工具侧检查 |
| 示例 | `examples/example.yaml` 的 `baud` / `mcu` / `port` / `mode` 都加上了约束 |

新增 `tests/test_constraints.py` 29 项：配置期 10 项、判定细节 4 项、数据有效性 6 项、
运行期 3 项、端到端（`render` / `validate` / `check` 三个命令的退出码）3 项。

### 2. 文档重构：README 变回落地页

README 从 686 行压到 ~150 行 —— 它现在只回答"这是什么 / 怎么跑起来 / 去哪看细节"。
被移出的内容没有删，只是搬到了更合适的位置：

| 新位置 | 收了什么 |
| --- | --- |
| `docs/cli.md` | 四个命令的全部选项表、作为 Python 库使用、错误类型表 |
| `docs/setup.md` | 安装（uv / venv+pip / `setup.sh`）、Ubuntu 与外置 NTFS 盘的坑、Python 版本策略、验证链的外部依赖 |
| `docs/publishing.md` | 发布到 GitHub / PyPI、密钥扫描、push protection、noreply 邮箱 |

同时：
* README 顶部注明**本项目由 DeepSeek 自动生成**，并加上 CI / Python / License 徽章。
* 删掉了 README 里与 `CHANGELOG.md`、`docs/template_guide.md` 重复的版本变更史与语法细节。
* 新增「现场用例：ABS FPI 内外压 → GeniE」一节，并注明里面的舱容与工况是**示意数值**。

### 3. 跨文件复用：`extends`

**需求**：一个项目要在一本工作簿里放几套规范、共用一张 Global 表，此前只能把变量**手抄**
进一个文件（`abs_fpi` 是靠项目侧脚本 `compose.py` 绕过的）。现在工具内建：

```yaml
extends: [rules/external.yaml, rules/internal.yaml]
```

| 变化 | 说明 |
| --- | --- |
| 根节点新增 `extends` | 递归合并（A extends B extends C 可以），循环引用会报错并打出引用链 |
| 合并判据与 `compose.py` 一致 | 同名变量：`prefix` / `suffix` / `type` / `derived` / 取值约束**必须逐字一致**，`default` 与 `description` 不一致只**告警**（保留先出现的）；同名模板内容必须一致。取值约束也纳入了"必须一致"的字段 |
| **`template_file` 相对声明它的文件解析** | 这是 `compose.py` 做不到的那件事（它要求各文件在同一目录）。新增 `TemplateDef.source_dir`，`renderer` / `excel_io` 两处路径解析都优先用它 |
| 告警可见 | `ProjectConfig.load_warnings` + CLI 的 `_load_project()`：合并告警以 `!` 开头打印出来 |
| 文档 | 指南新增 §16（合并规则、`template_file` 解析、与 `compose.py` 的分工）；`compose.py` 的 docstring 改写成"它现在还多做什么" |

新增 `tests/test_extends.py` 16 项：合并语义 4 项、冲突判定 4 项、错误写法 4 项、
`template_file` 路径解析 2 项、无 `extends` 时行为不变 2 项。

### 4. 公式模式支持行内 `{% if %}`

**需求**：公式模式此前只做纯替换，`{% if %}` 一律报错。但"有分支就退回快照模式"代价很大 ——
退回快照就意味着改参数得重跑脚本，正是 §14 想解决的问题。

| 变化 | 说明 |
| --- | --- |
| 行内 `{% if 条件 %}A{% else %}B{% endif %}` | 编译成 Excel `IF(条件,A,B)`，可嵌套；`{# 注释 #}` 也一并支持（整段丢掉） |
| 硬约束：**必须整段写在同一行** | 一行模板 = 一个单元格，跨行分支会改变行数，映射不到固定单元格；报错信息直接讲明这一点 |
| 条件里必须写 `.value` | 裸变量在快照模式里是 `VarValue` 对象，`VarValue > 20` 直接抛 `TypeError`；`.value` 在两种引擎里都是纯值。规则与 `case_filter` 一致，报错时给出改法 |
| 条件子集 | 比较（`==` `!=` `<` `>` `<=` `>=`）、逻辑（`and` `or` `not`）、真假判断（数值比 0 / 文本比空串）、`case_name` / `template_name` |
| 复用派生参数的翻译器 | `derived.to_excel(..., condition=True)` —— 同一套"表达式 → Excel 公式"逻辑，不另写一份 |

**顺带修掉一个真 bug**：jinja2 里 `nodes.And` / `nodes.Or` 是 `nodes.BinExpr` 的**子类**，
而 `derived._translate` 先命中 `BinExpr` 分支 —— 于是 `and` / `or` 从来没能翻译成功
（那两行是死代码，测试里还有一个 `if False else` 的规避写法）。现在把 And/Or 的判断排到
BinExpr 前面，`{% if flag.value > 0 and kind.value == "EXT" %}` 才真正可用。

**求值器同步扩展**（`formula_eval.py`）：补上 `<` `>` `<=` `>=` 与 `AND` / `OR` / `NOT`，
否则 `check --values` 遇到这些公式会报"不能识别的公式片段"—— 那就等于新功能没有被验证。

新增 `tests/test_formula_if.py` 25 项（编译形态 11、报错 8、端到端 2、其余边界）；
`tests/test_formula_eval.py` 补 12 项运算符用例；`tests/test_derived.py` 补 4 项（含把
那个规避写法换成真断言）。

### 5. 质量护栏：ruff + mypy + 覆盖率门槛

CI 此前只跑测试 —— "能不能跑"有了，"写得对不对"没人管。现在四条命令进 CI
（`ruff check` / `ruff format --check` / `mypy` / `pytest --cov`），三个系统 × 两个
Python 版本各跑一遍。

| 变化 | 说明 |
| --- | --- |
| `[tool.ruff]` | 显式列出规则集（不依赖 ruff 默认值，免得升级时行为漂移）：`E4/E7/E9`、`F`、`I`、`UP`、`B`、`SIM`、`RUF`、`PLW1510`。中文项目要关掉 `RUF001-003`（全角标点被当成"歧义字符"）；Typer 的 `B008`、行长 `E501` 也关掉 |
| **`target-version = "py311"`** | 与 `requires-python` 必须一致。这条护栏**当场就抓到一个真问题**：`formula.py` 里 f-string 的表达式里写了引号（Python 3.12 才允许），本地 3.12 跑得好好的，一 push 到 3.11 就是语法错误 |
| `[tool.mypy]` | `python_version = "3.11"`、只查 `excel_codegen/`、`check_untyped_defs`。新增代码一律带注解 |
| `[tool.coverage.report]` | `fail_under = 85`（当前 86.4%，分支覆盖）。覆盖率掉了直接红 |
| `ruff format` | 全仓库归一（45 个文件） |

**顺手修掉两个被 lint/类型检查抓出来的真 bug**：

1. **`excel_io.py` 用了没导入的 `DerivedError`**（F821）—— 派生参数的 resolver 在
   "引用了不合法的名字"时本该抛出带提示的 `DerivedError`，实际会先炸 `NameError`。
2. **`cli.py` 里 `_first_difference` 定义了两次** —— 0.4.0 重写差异信息时留下的旧版本被
   新版遮蔽，成了几百行里没人注意的死代码（mypy 的 `no-redef` 把它揪出来了）。

顺带清掉一批：未使用的导入/变量、`zip()` 缺 `strict=`、`try/except/pass` 改
`contextlib.suppress`、嵌套 `if` 合并、`FilterValue` 的 `__slots__` 属性补类型标注、
`Evaluator` 的 `tokens/index` 在 `__init__` 里显式初始化（原来是 `getattr(..., None)`）。

### 6. 发布流程：`.github/workflows/release.yml`

推 `v*` tag → 校验 tag 与 `pyproject.toml` 的版本一致 → 跑一遍质量闸 → `uv build` →
`twine check` → 用 **Trusted Publishing（OIDC）** 发到 PyPI，**仓库里不存任何 token**。
手动触发（`workflow_dispatch`）默认发到 TestPyPI，用来演练。

发布前的登记步骤见 [`docs/publishing.md`](docs/publishing.md)。

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
| 发布准备（公开仓库） | 补 `LICENSE`（MIT / `JG. Luo`）并升级到 PEP 639 元数据；`.gitignore` 增加密钥与凭据兜底规则；README 新增「发布与隐私」一节（密钥扫描、GitHub push protection、noreply 邮箱）。提交作者改为 `129145708+Dog-Chest@users.noreply.github.com` |

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
