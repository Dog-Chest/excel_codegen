# ABS FPI 内外压 —— excel_codegen 模板库

用 `excel_codegen` 生成 **GeniE 面载荷函数体**（ABS FPI 2025 Part 5A Ch.3 Sec.2 的内压与外压）。

**成品是一本 Excel**：在 `Global Parameter` / `Local Parameter` 表里填参数，
`Code` 表里就是可以直接粘进 GeniE 的整段函数体 —— 而且是 **Excel 公式**，
改参数后 Excel / WPS 打开即重算，**不需要跑任何脚本**。

```
function(x, y, z) {
    // 把 Code 表某一列的内容粘在这里
}
```

或整段函数交给 `Pressure2dJavascript(userFunction, refinement, positionSystem)`
（GeniE Help 8.12，8. Loads → 8.1.3.3 Intensity/Pressure，Javascript 项）。
GeniE 会对被加载面的**每一个离散点**调用它，所以 x/y/z 是当前加载点。

> 模板库的索引、以及"怎么加下一个规范（DNV / BV …）"的清单在 **[TEMPLATES.md](TEMPLATES.md)**。
> 对 `excel_codegen` 本身的实测报告（含修复复测）在 **[FINDINGS.md](FINDINGS.md)**。

---

## 1. 三本工作簿

| 工作簿 | 内容 | 工况 |
|---|---|---|
| `abs_fpi_external.xlsx` | 外压 5A-3-2/5.5 单用 | 3 |
| `abs_fpi_internal.xlsx` | 内压 5A-3-2/5.7 单用 | 4 |
| `ABS_FPI_load_cases.xlsx` | **项目工作簿**：两套规则、**共用一张 Global 表**（主尺度只填一次） | 7 |

第三本由 `compose.py` 把前两个规则集合并而成，用 `case_filter` 分流
（`genie_ext` 只对 `kind == 'EXT'` 的工况出代码，`genie_int` 只对 `INT`），
输出表分成 `Code EXT` / `Summary EXT` / `Code INT` / `Summary INT`。

## 2. 一条命令

```bash
cd excel_codegen/abs_fpi
python build.py            # compose → 建骨架（缺哪个建哪个）→ 渲染写回 → 导出 → 两层验证
python build.py --check    # 只验证，不动文件
python build.py --init     # 骨架不存在时也重建（会清掉你填的参数）
```

首次跑 `--init`，之后 `build.py` **只写 Output 表，绝不碰**你在 Global / Local 里填的值。

## 3. ⚠ 公式模式：什么免重跑、什么不免

**改参数** → 免重跑。打开 Excel 就重算（写公式时同时设了 `fullCalcOnLoad`）。

**改模板 / YAML** → 要重跑 `build.py`（公式是写进格子的文本，得刷新一遍）。

**改结构** → 分两种，都实测过（`probes/probe_formula_structure.py`）：

| 动作 | 结果 |
|---|---|
| 插 / 删 **变量行** | ✅ 不用重跑。公式用 `INDEX/MATCH` **按变量名**定位，插行不指错 |
| 增 / 删 / 移动 **Case 列** | ❌ 必须重跑。公式里的列标是绝对字母，工况一移位就指着别的工况了 |

最后一行的实测证据（`probe_formula_structure.py` N2：在 Local 表插一个 Case 列）：

```
verify_excel_engine.py → exit 1   公式已过期 ✓（求值出来的代码与参数不一致）
           第 3 行：公式算出 '| case_name | WBT6-d15.059-mu0 | ...'
                     Python  '| case_name | WBT7-d20.559-mu90 | ...'
excel-codegen check    → exit 1   发现过期 ✓（可放进 CI）
```

顺带一个坑：**手加 Case 列时一定要填 `kind`**。留空会回落 `default`（本例是 `EXT`），
于是这一列会**悄悄进外压规则集**。

## 4. 校验链（四层）

```bash
python build.py                                    # 第 1、2 层，28 项
node compare_with_rules.js                         # 第 4 层，21 项（需外部 GeniE/Rules 仓库）
python -m excel_codegen validate -c abs_fpi.yaml   # 第 0 层：配置与模板语法
```

> ⚠ 第 4 层需要**本仓库之外**的同级目录 `../../GeniE/Rules`（独立仓库，本工作区没有）。
> 该目录不存在时脚本会打印 `SKIP` 并以退出码 **2** 结束（加 `--allow-missing` 则按跳过处理、
> 退出码 0）—— 明确跳过，**不会**把"没验证"伪装成"通过"。
> 其余三层不依赖外部仓库，可照常跑。

| 层 | 抓什么 |
|---|---|
| `validate` | YAML 结构、模板语法、变量未定义 / 定义了没人用、`output_sheet` 未声明、公式模式的越界写法（带行内容） |
| `check` | ① 表里的公式 / 快照是否与当前 YAML + 参数一致；② **把公式在 Python 里算一遍**与 Python 渲染逐行比对（0.4.0 起默认开） |
| **`verify_excel_engine.py`** | **同一个断言，但实现独立**（另一套公式解析器）：两条独立实现都过，才算证据 |
| `compare_with_rules.js` | 与 `GeniE/Rules` 那套带量纲检查的产物逐行一致 |

第 3 层是 0.3.0 那轮加的，**0.4.0 起工具自带的 `check` 也会做同一件事**（`--no-values` 可关：
参数单元格本身是公式时读不到值，会自动降级并提示）。两者都保留，因为它们**实现独立** ——
`verify_excel_engine.py` 是这个项目自己的解析器，`check` 用的是工具里的 `formula_eval.py`；
同一个 bug 不太可能在两条独立实现里同时出现，所以两条都过才算"Excel 里看到的"与
"`--outdir` 导出的"确实是同一段代码（都抓不到的是"真 Excel 的行为"：区域设置、`TEXT()` 格式串）。

它还带一个 `--change VAR=VALUE`：改一格参数、**不重跑 render**、重新求值，
用来验证"改参数自动重算"这句话本身：

```
== 改参数后重算（副本 abs_fpi_external.xlsx）
  [改参数] Global Parameter!B2  L = 400（没有重跑 render）
  [公式] genie_ext  3 个 Case  → Code @ B2
  6 项，0 项失败
```

## 5. 生成物长什么样

`generated/EXT-T20.559-mu90-kf+1.js`：

```javascript
// EXT-T20.559-mu90-kf+1
// The rules live in the RULE frame: x from the A.P., y from the centre line
// positive starboard, z above the baseline. ...
var t1 = x;
var t2 = y;
var t3 = z;
x = t1;
y = t2;
z = t3;
// --- global parameters (same for every case)
var L = 340 m;
...
var p_s = Math.max(rho_sea * g * (draft - z), 0 Pa);   // still-water pressure, ...
var p_d = rho_sea * g * ESF_side * k_u * h_de;   // hydrodynamic pressure, ...
return Math.max(p_s + p_d, 0 Pa);   // the pressure the function returns
```

坐标映射由 6 个变量纯替换而成（`x = {{ neg_x }}t{{ ax_x }};`），
所以**公式模式也能表达**：把 `ax_y` 填 `3`、`neg_z` 填 `-` 再打开 Excel，
变换语句立刻变成 `y = t3; z = -t2;`。

> 恒等映射下这 3 行是空操作（`x = t1;`），仍然会输出 —— 公式模式没有条件，
> 没法"恒等就不输出"。GeniE 侧无副作用。

## 6. 与 `GeniE/Rules` 的关系

`node compare_with_rules.js` 把两边的产物做**去注释、去坐标变换后的逐行比对**，
并借 `GeniE/Rules` 的检查器跑一遍：

```
== EXT  (3 cases)
== INT  (4 cases)
21 checks, 0 failure(s)
```

即这 7 段代码与 `GeniE/Rules`（带量纲检查、`if/else` 检查、48 种映射检查）的输出**完全一致**。

两者是**互补**关系，不是替代：

- `excel_codegen` 管"参数怎么填、模板怎么复用、结果怎么写回"；它不认识 GeniE 的
  量纲，不知道 `if` 必须配 `else`，也不会拦 `Math.sin(90)`（裸数会被当弧度）。
- `GeniE/Rules` 管"生成的东西对不对"：量纲闭合、声明顺序、保留字、每一层 `if/else`、
  48 种坐标映射。

所以本目录的 `build.py` 把两边**串起来跑**：`excel_codegen` 负责生成，
`compare_with_rules.js` 负责确认生成结果与那套带检查的实现一致。

## 7. 已知边界

- **本项目当前仍把舱数据按工况摊平**：一个舱被 N 个工况用到就写 N 遍，`tank_ref` 只作标注。
  > 工具 **0.7.0 起支持第三层作用域**（`variables.group` 成员表，见
  > `docs/template_guide.md` §18）—— 一个舱的参数可以只写一遍、由 Case 指向它。
  > 本项目**尚未迁移**（迁移意味着重做工作簿布局与 `compose.py` 的输出表改写），
  > 所以下面是"当前状态"，不是"工具做不到"。见 `FINDINGS.md` #3(a)。
- **`case_filter` 的 Case 必须成块连续**（公式模式横向相对列靠恒定偏移），
  所以项目工作簿里 EXT 的 7 个工况在前、INT 的 4 个在后。
- **公式的 `TEXT()` 格式串受区域设置影响**：中文/英文区域小数点是 `.`；
  换到欧洲区域（`,`）之前要重验一遍。
- **公式值只活在 Excel / WPS 里**：`--outdir` 导出的文件由 Python 渲染（内容等价，
  因为模板是纯替换），但"改参数自动同步文件"做不到，要再跑一次 `build.py`。

## 8. 文件

见 [TEMPLATES.md §7](TEMPLATES.md#7-文件清单)。要点：

| 文件 | 作用 |
|---|---|
| `abs_fpi_external.yaml` / `abs_fpi_internal.yaml` | 规则集真源（变量 + 模板声明） |
| `compose.yaml` + `compose.py` | 合并规则集 → 项目工作簿 YAML（含冲突检查） |
| `build.py` | 一条命令：compose + 生成 + 导出 + 两层验证 |
| `verify_excel_engine.py` | 公式求值器（补 `check` 在公式模式下比不了值的缺口） |
| `fill_cases.py` | 样例工况取值（只有首次 `--init` 用） |
| `compare_with_rules.js` | 与 `GeniE/Rules` 逐行比对 |
| `probes/` | 对这一版工具的最小复现与探针 |
| `FINDINGS.md` | 对 `excel_codegen` 的实测报告 |
