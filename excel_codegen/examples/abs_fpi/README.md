# 现场用例：ABS FPI 内外压 → GeniE

用 `excel_codegen` 生成 **GeniE 面载荷函数体**（ABS FPI 2025 Part 5A Ch.3 Sec.2 的内压与外压）。

**成品是一本 Excel**：在 `Global Parameter` / `Local Parameter` 表里填参数，`Code` 表里就是
可以直接粘进 GeniE 的整段函数体 —— 而且是 **Excel 公式**，改参数后 Excel / WPS 打开即重算，
**不需要跑任何脚本**。

```
function(x, y, z) {
    // 把 Code 表某一列的内容粘在这里
}
```

或整段函数交给 `Pressure2dJavascript(userFunction, refinement, positionSystem)`
（GeniE Help 8.12，8. Loads → 8.1.3.3 Intensity/Pressure，Javascript 项）。
GeniE 会对被加载面的**每一个离散点**调用它，所以 x/y/z 是当前加载点。

> 这是随包发布的**示例**。开发侧的现场脚本、探针与逐轮实测报告在仓库根的
> [`abs_fpi/`](../../../abs_fpi/)（不进包）。

> ⚠️ **规范版权与免责**
>
> * 本示例依据 **ABS《Rules for Building and Classing》**（5A-3-2/5.5 外压、5.7 内压）
>   实现。**规范文本版权归 American Bureau of Shipping 所有**；本仓库只包含演示所必需的
>   公式与系数，**不含**规范原文、表格或图表。
> * 使用者须**自行取得正式规范**并以其为准。本实现**未经 ABS 审核或认可**，
>   本项目与 ABS **无关联**，`ABS` 是 American Bureau of Shipping 的商标。
> * 输出**不能替代**船级社审查与批准。用于实际设计 / 建造 / 送审前，必须由具备资格的人员
>   独立复核。
> * 示例中的舱容与工况只是**示意数值**，不代表任何真实船舶。
>
> 完整声明见仓库根的 `NOTICE.md`（也随发行包发布在 `.dist-info/licenses/` 下）。

---

## 1. 三本工作簿

| 工作簿 | 内容 | 工况 |
|---|---|---|
| `abs_fpi_external.xlsx` | 外压 5A-3-2/5.5 单用 | 3 |
| `abs_fpi_internal.xlsx` | 内压 5A-3-2/5.7 单用 | 4 |
| `ABS_FPI_load_cases.xlsx` | **项目工作簿**：两套规则、**共用一张 Global 表**（主尺度只填一次） | 7 |

打开任意一本就能改参数 —— 三本都已经填好样例数值。

第三本由 `compose.yaml` 的配方把前两个规则集合并而成，用 `case_filter` 分流
（`genie_ext` 只对 `kind == 'EXT'` 的工况出代码，`genie_int` 只对 `INT`），
输出表分成 `Code EXT` / `Summary EXT` / `Code INT` / `Summary INT`。

后两本还有一张 **`Tank Data` 成员表**：一行一个舱（WBT6 / WBT7 / COT1），
B 列起一个变量一列。舱的参数只写一遍，工况用 `tank_ref` 指向它（指南 §18）。

## 2. 怎么用

```bash
# 打开工作簿改参数即可；要把代码导出成文件时：
excel-codegen render -c abs_fpi_internal.yaml -x abs_fpi_internal.xlsx --outdir generated

# 核对「Excel 里算出来的」与「Python 渲染的」是否一致：
excel-codegen check -c abs_fpi_internal.yaml -x abs_fpi_internal.xlsx

# 改了模板 / YAML 之后，要把新公式写回 Excel：
excel-codegen render -c abs_fpi_internal.yaml -x abs_fpi_internal.xlsx --write-excel
```

同目录的 `*_render.bat`（Windows 双击）/ `*_render.sh` 就是最后一条命令。

## 3. 公式模式：什么免重跑、什么不免

**改参数** → 免重跑。打开 Excel 就重算（写公式时同时设了 `fullCalcOnLoad`）。

**改模板 / YAML** → 要重跑 `--write-excel`（公式是写进格子的文本，得刷新一遍）。

**改结构** → 分两种，都实测过：

| 动作 | 结果 |
|---|---|
| 插 / 删 **变量行** | ✅ 不用重跑。公式用 `INDEX/MATCH` **按变量名**定位，插行不指错 |
| 增 / 删 / 移动 **Case 列** | ❌ 必须重跑。公式里的列标是绝对字母，工况一移位就指着别的工况了 |
| 改 **Tank Data** 里的舱数据 | ✅ 不用重跑。公式按"Case → `tank_ref` → 成员行 → 变量列"两级 `INDEX/MATCH` 定位 |
| 增 / 删 **Tank Data** 的成员行 | ✅ 不用重跑（同上，按成员名 `MATCH` 找行） |

顺带一个坑：**手加 Case 列时一定要填 `kind`**。留空会回落 `default`（本例是 `EXT`），
于是这一列会**悄悄进外压规则集**。

## 4. 生成物长什么样

`generated/EXT-T20.559-mu90-kf+1.js`：

```javascript
// EXT-T20.559-mu90-kf+1
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

## 5. 已知边界

- **`case_filter` 的 Case 必须成块连续**（公式模式横向相对列靠恒定偏移），
  所以项目工作簿里 EXT 的工况在前、INT 的工况在后。
- **公式的 `TEXT()` 格式串受区域设置影响**：中文/英文区域小数点是 `.`；
  换到欧洲区域（`,`）之前要重验一遍。
- **公式值只活在 Excel / WPS 里**：`--outdir` 导出的文件由 Python 渲染（内容等价，
  因为模板是纯替换），但"改参数自动同步文件"做不到，要再跑一次 `render`。

## 6. 这个目录里有什么

| 文件 | 作用 |
|---|---|
| `abs_fpi_external.yaml` / `abs_fpi_internal.yaml` | 两个规则集真源（变量 + 模板声明） |
| `abs_fpi.yaml` | 项目工作簿 YAML（由 `compose.yaml` 的配方合并生成） |
| `compose.yaml` | 合并配方：把哪些规则集并进一本、各自用什么 `case_filter` |
| `templates/*.j2` | GeniE 函数体与摘要模板 |
| `generated/` | 导出产物（提交进仓库的实测结果，可用来对照） |
| `*.xlsx` | 三本已经填好样例参数的工作簿 |

> 用例里的舱容与工况组合只是**示意数值**（随手取的几个数），不代表任何真实船舶。
