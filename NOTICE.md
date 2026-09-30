# 商标、版权与免责声明（NOTICE）

本文件是 [excel_codegen](README.md) 的法律声明汇总。

> **本项目与下列任何厂商均无关联，也未获得其赞助、认可或背书。**

---

## 1. 商标归属

文档与示例中会出现下列名称。它们仅用于**说明兼容性、或说明所依据的规范**（描述性 /
指示性使用），相关商标归各自所有者所有：

| 名称 | 商标所有者 | 在本项目中的用途 |
| --- | --- | --- |
| Microsoft Excel / Excel | Microsoft Corporation | 说明本工具生成与读写的是 Excel 工作簿（`.xlsx`） |
| NASTRAN | MSC Software Corporation（最初由 NASA 开发） | 说明示例生成的是 NASTRAN 工况控制语句 |
| GeniE | DNV AS | 说明示例生成的是 GeniE 的 JavaScript 函数体 |
| ABS（含 ABS FPI / Rules for Building and Classing） | American Bureau of Shipping | 说明示例实现的是一种船级社规范的载荷计算方法 |
| WPS | 金山办公（Kingsoft Office） | 说明生成的工作簿也能用 WPS 打开 |
| STM32 | STMicroelectronics | 示例中的 MCU 名称，仅作举例 |
| Python、PyPI | Python Software Foundation | 运行环境与分发平台 |

**本项目不使用**上述任何厂商的 logo、图标、字体或其他品牌视觉元素。

---

## 2. 关于规范文本与工程计算

`abs_fpi` 示例实现的是 **ABS《Rules for Building and Classing》**（5A-3-2/5.5 外压、
5.7 内压）中的面载荷计算方法。请注意：

* **规范文本本身受版权保护，版权归 American Bureau of Shipping 所有。** 本仓库只包含
  为演示工具用法所必需的公式与系数，**不包含**规范原文、表格或图表；
* 使用者应**自行取得正式规范**，并以其为准；本项目的实现**未经 ABS 审核或认可**；
* 本项目**不是**任何船级社的送审工具，输出结果**不能替代**船级社的审查与批准。

---

## 3. 工程用途免责

本工具生成的是**工程计算代码**。生成结果可能因输入数据、模板版本、区域设置
（例如公式里 `TEXT()` 的数字格式串）等原因而与预期不符：

* **用于实际设计、建造或送审之前，必须由具备资格的人员独立复核**；
* 作者与贡献者**不对任何使用后果承担责任**（另见 [`LICENSE`](LICENSE) 的免责条款）。

---

## 4. 关于 AI 生成

本项目由 AI（DeepSeek）生成，人工负责提需求、定方向与验收（见 [README](README.md)）。

请注意：纯 AI 生成内容在部分法域可能**不构成受版权保护的客体** —— 也就是说，
[`LICENSE`](LICENSE) 里的 MIT 授权对这部分内容未必可执行，你也未必能阻止他人复制。
这不影响你使用本项目，但会限制你**主张权利**。
