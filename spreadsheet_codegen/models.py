"""Pydantic 数据模型：YAML 配置结构 + 运行期数据（CaseData / RenderResult）。

配置文件结构::

    version: 1
    excel:
      output: "template.xlsx"
      template_sheet: "Template"
      sheets:
        global: "Global Parameter"
        local: "Local Parameter"
        outputs: ["Output"]
    variables:
      global: [...]
      local: [...]
    templates: [...]
"""

from __future__ import annotations

import re
from collections.abc import Mapping
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Literal

import yaml
from pydantic import (
    BaseModel,
    ConfigDict,
    Field,
    ValidationError,
    field_validator,
    model_validator,
)

from .utils import (
    ConfigError,
    VarValue,
    column_index_to_letter,
    is_identifier,
    parse_cell,
    split_lines,
    to_text,
)

__all__ = [
    "FIRST_CASE_COLUMN",
    "CaseData",
    "Direction",
    "ExcelConfig",
    "ExcelSheets",
    "ProjectConfig",
    "RenderResult",
    "TemplateDef",
    "VarType",
    "VariableDef",
    "VariablesConfig",
    "format_validation_error",
    "load_config",
]

#: Local Parameter 工作表中第一个 Case 列（E 列），A-D 为变量名/描述/Prefix/Suffix。
FIRST_CASE_COLUMN: int = 5

Direction = Literal["horizontal", "vertical"]
VarType = Literal["auto", "string", "int", "float", "bool", "raw"]
#: ``excel``    = 把模板编译成**Excel 公式**写进输出表（**默认**）：改参数由 Excel 自己重算，
#:               工作簿脱离命令行也独立可用；
#: ``snapshot`` = 脚本渲染后把**文本**写进输出表，改参数必须重跑命令。
#:                模板要用 ``{% for %}`` / 过滤器 / 多行 ``{% if %}`` / ``{% include %}`` 时才需要。
Engine = Literal["snapshot", "excel"]

#: 渲染上下文里由程序注入、不允许作为变量名使用的保留名。
RESERVED_NAMES: frozenset[str] = frozenset({"case_name", "template_name"})


def _duplicates(items: list[str]) -> list[str]:
    seen: set[str] = set()
    dupes: list[str] = []
    for item in items:
        if item in seen and item not in dupes:
            dupes.append(item)
        seen.add(item)
    return dupes


def _number_text(value: float) -> str:
    """数字的提示文本：整数不显示小数点（``20.0`` → ``20``）。"""
    return to_text(int(value)) if float(value).is_integer() else to_text(value)


def _as_number(value: Any) -> float | None:
    """把取值当数字读；读不出来返回 ``None``（bool 不算数字）。"""
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = to_text(value).strip()
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def _valid_sheet_name(value: Any) -> str:
    """校验并规范化一个 Excel 工作表名。"""
    name = to_text(value).strip()
    if not name:
        raise ValueError("工作表名不能为空")
    if len(name) > 31 or any(char in name for char in "[]:*?/\\"):
        raise ValueError(f"工作表名 {name!r} 非法（Excel 限制：<=31 字符且不能含 []:*?/\\）")
    return name


def format_validation_error(exc: ValidationError) -> str:
    """把 pydantic 的校验错误整理成多行、便于阅读的中文提示。"""
    lines: list[str] = []
    for error in exc.errors():
        location = ".".join(str(part) for part in error.get("loc", ())) or "<root>"
        lines.append(f"  - {location}: {error.get('msg', '非法取值')}")
    return "\n".join(lines)


# --------------------------------------------------------------------------- #
# 变量定义
# --------------------------------------------------------------------------- #
class VariableDef(BaseModel):
    """一个变量的定义（名称、描述、默认值、前缀、后缀、类型）。"""

    model_config = ConfigDict(extra="forbid")

    name: str
    description: str = ""
    #: 单位（纯文档，只写进 Excel 批注与 `validate` 的清单；**不会**出现在生成的文本里）。
    #: 单位要进生成结果请用 ``suffix``（``suffix: " m"`` → ``340 m``）。
    unit: str = ""
    default: Any = ""
    prefix: str = ""
    suffix: str = ""
    type: VarType = "auto"
    #: 派生参数：一段 Jinja2 **表达式**，只能引用 ``global`` 与**同一个 Case** 的 ``local``。
    #: 例如 ``derived: "k_c * h_di"``。派生参数不用在 Excel 里填值（那一格由工具写成公式或算好的值）。
    derived: str | None = None
    #: ``init`` 建表时是否把 ``default`` 预填进新格子（默认 ``True``，老行为不变）。
    #: 设为 ``False``：新表格里这一列/行留空，``default`` 只作为"读取时的兜底值"存在
    #: —— 适合"每次都该重新填"的参数。
    prefill: bool | None = None
    #: 单元格为空时是否回落 ``default``（默认 ``True``，老行为不变）。
    #:
    #: 设成 ``False`` 是"某类型才有的字段"的正解：``default`` 仍然预填进新表（起提示作用），
    #: 但用户在表里**把这格清空**就表示"这条记录没有这个字段"，生成结果里它真的是空的 ——
    #: 不会被默认值悄悄填上。默认语义下"没填"与"填了默认值"在结果里不可区分，是**生成错数据**的常见来源。
    #: 与 :attr:`prefill` 一起写 ``false`` 时，这个变量就完全不受 ``default`` 影响。
    fallback: bool | None = None

    #: 取值约束（可选，只对"填写型"变量有效；派生参数的值是算出来的，不能加）。
    #: ``min`` / ``max``：数值上下限。``type`` 为 string / bool / raw 时不允许。
    min: float | None = None
    max: float | None = None
    #: 允许的取值集合，按**文本**比较（``1`` 与 ``1.0`` 视为同一个值）。
    #: 会同时写成 Excel 的下拉列表。
    choices: list[Any] | None = None
    #: 整串匹配的正则（``re.fullmatch``）。
    pattern: str | None = None
    #: 声明了约束时**是否允许空值**（默认 ``False``，老行为不变）。
    #:
    #: 默认语义是"有约束就意味着必须给一个合法取值"，空值算不合格。这对"某类型才有的字段"
    #: 很别扭：``card_type`` 只有银行卡那几行才有，于是只能往 ``choices`` 里塞一个空串
    #: （``choices: ["", "储蓄卡", ...]``）—— 下拉列表第一项是空的，读起来像"允许空"，
    #: 而文档又说空值不合格，两边说法冲突，而且这个绕法没有任何文档提示。
    #:
    #: 写 ``allow_blank: true`` 就直说"这个变量可以留空"：
    #:
    #: * 空值通过校验（不再是"取值（空）不在允许列表里"）；
    #: * Excel 的下拉列表里**不再出现那个空选项**（选项就是你写的那些）；
    #: * 与 :attr:`fallback` 搭配最自然：``fallback: false`` + ``allow_blank: true``
    #:   = "这条记录没有这个字段"。
    #:
    #: 注意"必填"仍然该由 :attr:`asserts`（或取值约束）表达 —— 这个开关只管"空值本身合不合法"。
    allow_blank: bool = False

    @field_validator("name")
    @classmethod
    def _check_name(cls, value: str) -> str:
        name = to_text(value).strip()
        if not name:
            raise ValueError("变量名不能为空")
        if not is_identifier(name):
            raise ValueError(
                f"变量名 {name!r} 不是合法标识符（需匹配 [A-Za-z_][A-Za-z0-9_]*），因为要在 Jinja2 模板中直接引用"
            )
        if name in RESERVED_NAMES:
            raise ValueError(f"变量名 {name!r} 是保留字（渲染时由程序注入），请改名")
        return name

    @field_validator("description", "unit", "prefix", "suffix", mode="before")
    @classmethod
    def _to_str(cls, value: Any) -> str:
        return "" if value is None else to_text(value)

    @field_validator("derived", mode="before")
    @classmethod
    def _derived_to_str(cls, value: Any) -> str | None:
        if value is None:
            return None
        text = to_text(value).strip()
        return text or None

    @field_validator("min", "max", mode="before")
    @classmethod
    def _number_or_none(cls, value: Any) -> float | None:
        if value is None:
            return None
        if isinstance(value, bool):
            raise ValueError("min / max 必须是数字，不能是 true/false")
        if isinstance(value, (int, float)):
            return float(value)
        text = to_text(value).strip()
        if not text:
            return None
        try:
            return float(text)
        except ValueError:
            raise ValueError(f"min / max 必须是数字，得到 {value!r}") from None

    @field_validator("choices", mode="before")
    @classmethod
    def _check_choices(cls, value: Any) -> list[Any] | None:
        if value is None:
            return None
        if isinstance(value, (str, bytes)) or not isinstance(value, (list, tuple)):
            raise ValueError("choices 必须是列表，例如 choices: [EXT, INT]")
        items = list(value)
        if not items:
            raise ValueError("choices 不能是空列表（不想要约束就删掉这个字段）")
        texts = [to_text(item) for item in items]
        dupes = sorted({text for text in texts if texts.count(text) > 1})
        if dupes:
            raise ValueError(f"choices 里有重复取值: {', '.join(dupes)}")
        return items

    @field_validator("pattern", mode="before")
    @classmethod
    def _check_pattern(cls, value: Any) -> str | None:
        if value is None:
            return None
        text = to_text(value).strip()
        if not text:
            return None
        try:
            re.compile(text)
        except re.error as exc:
            raise ValueError(f"pattern 不是合法的正则: {exc}") from None
        return text

    @model_validator(mode="after")
    def _check_derived(self) -> VariableDef:
        if self.derived and to_text(self.default) != "":
            raise ValueError(
                f"变量 {self.name!r} 同时写了 derived 与 default —— 派生参数的值由表达式算出来，"
                "不能同时给它一个填写值；请删掉其中一个"
            )
        return self

    @model_validator(mode="after")
    def _check_constraints(self) -> VariableDef:
        constrained = self.min is not None or self.max is not None or self.choices or self.pattern
        if not constrained:
            if self.allow_blank:
                raise ValueError(
                    f"变量 {self.name!r} 写了 allow_blank 但没有任何取值约束 —— "
                    "本来就没人拦空值，删掉 allow_blank 或加上 choices / min / max / pattern"
                )
            return self
        if self.is_derived:
            raise ValueError(
                f"变量 {self.name!r} 是派生参数（值由表达式算出来），不能加 min / max / choices / pattern —— "
                "请在表达式里约束（例如 max(x, 0)），或把它改成填写型变量"
            )
        if self.min is not None and self.max is not None and self.min > self.max:
            raise ValueError(f"变量 {self.name!r} 的 min({_number_text(self.min)}) 大于 max({_number_text(self.max)})")
        if (self.min is not None or self.max is not None) and self.type in ("string", "bool", "raw"):
            raise ValueError(
                f"变量 {self.name!r} 的 type 是 {self.type!r}，不能加 min / max —— "
                "数值范围只对 type: int / float / auto 有意义"
            )
        # 默认值是 YAML 自己写的，违反约束属于配置错误，立刻指出（空默认值允许：表示"必须去表里填"）
        if to_text(self.default) != "":
            problem = self.value_problem(self.default)
            if problem:
                raise ValueError(f"变量 {self.name!r} 的 default {problem}")
        return self

    @model_validator(mode="after")
    def _check_prefill_and_fallback(self) -> VariableDef:
        """把 ``prefill`` / ``fallback`` 的默认值定下来，并拦住自相矛盾的写法。

        默认（两个都不写）= **老行为**：既预填进新表，空单元格也回落 —— 已有配置一行都不用改。
        ``prefill: false`` 时兜底值没有再落回的必要（表里根本没写过它），所以 ``fallback``
        默认跟着变成 ``false``；显式写 ``fallback: true`` 是自相矛盾，直接报错。
        """
        if self.prefill is None:
            self.prefill = True
        if self.fallback is None:
            self.fallback = bool(self.prefill)
        if not self.prefill and self.fallback:
            raise ValueError(
                f"变量 {self.name!r} 同时写了 prefill: false 与 fallback: true —— 自相矛盾："
                "既不预填进新表，又要空单元格回落，那 default 到底从哪来？"
                "两个都想要就删掉 prefill，只要'空就是空'就两个都写 false"
            )
        if self.is_derived and (self.prefill is False or self.fallback is False):
            raise ValueError(f"变量 {self.name!r} 是派生参数（值由表达式算出来），prefill / fallback 都不适用，请删掉")
        return self

    @property
    def effective_default(self) -> Any:
        """**回落**时该用的值：``fallback: false`` 时永远是 ``""``（空就是空）。

        读表处一律用它代替裸的 ``default`` —— 否则"清空单元格表示没有这个字段"表达不出来。
        """
        return self.default if self.fallback else ""

    @property
    def is_derived(self) -> bool:
        return bool(self.derived)

    @property
    def has_constraints(self) -> bool:
        return bool(self.min is not None or self.max is not None or self.choices or self.pattern)

    @property
    def constraint_text(self) -> str:
        """给人和给 Excel 提示用的一句话约束描述。"""
        parts: list[str] = []
        if self.choices:
            allowed = [to_text(item) for item in self.choices if to_text(item) != ""]
            parts.append("可选: " + " / ".join(allowed))
        if self.pattern:
            parts.append(f"格式: {self.pattern}")
        if self.min is not None and self.max is not None:
            parts.append(f"范围: {_number_text(self.min)} ~ {_number_text(self.max)}")
        elif self.min is not None:
            parts.append(f"范围: >= {_number_text(self.min)}")
        elif self.max is not None:
            parts.append(f"范围: <= {_number_text(self.max)}")
        if self.allow_blank:
            parts.append("可以留空")
        return "；".join(parts)

    def value_problem(self, value: Any) -> str | None:
        """检查一个取值是否满足约束；返回问题描述，没问题返回 ``None``。

        空值（``""``）在声明了约束时**默认算不合格** —— ``choices`` 之类的约束意味着
        "必须有个合法取值"。想让"某类型才有的字段"能留空，写 ``allow_blank: true``
        （不必再往 ``choices`` 里塞空串）。
        """
        text = to_text(value)
        if self.allow_blank and text.strip() == "":
            return None
        if self.choices:
            allowed = [to_text(item) for item in self.choices]
            if text not in allowed:
                shown = text if text != "" else "（空）"
                return f"取值 {shown} 不在允许列表 {'/'.join(allowed)} 里"
        if self.pattern is not None and re.fullmatch(self.pattern, text) is None:
            shown = text if text != "" else "（空）"
            return f"取值 {shown} 不匹配格式 {self.pattern}"
        if self.min is not None or self.max is not None:
            number = _as_number(value)
            if number is None:
                shown = text if text != "" else "（空）"
                return f"取值 {shown} 不是数字，但该变量声明了 min/max"
            if self.min is not None and number < self.min:
                return f"取值 {_number_text(number)} 小于下限 {_number_text(self.min)}"
            if self.max is not None and number > self.max:
                return f"取值 {_number_text(number)} 大于上限 {_number_text(self.max)}"
        return None


class GroupConfig(BaseModel):
    """第三层作用域：**成员表**（船 → 工况 → 舱/设备）。

    没有它的时候，一个被多个工况引用的舱只能把参数**按工况摊平**（同名舱在每个 Case 列里
    各写一遍，改一个舱的尺寸要改 N 列）。有了它，舱的参数只写一遍，Case 用一个"指针变量"
    （``key``）指向自己用哪个成员。

    Excel 布局与 Global / Local 都不同：**一行一个成员，B 列起一个变量一列**
    （B 列表头是变量名）—— 这正是工程师写"舱容表"的习惯，也让公式模式能用与
    global / local 同一形态的 ``INDEX/MATCH`` 定位。
    """

    model_config = ConfigDict(extra="forbid")

    #: 成员表的工作表名。
    sheet: str = "Group Data"
    #: 哪个 **local** 变量存"这个 Case 用哪个成员"（成员表里 A 列的名字之一）。
    key: str
    #: init 时先建这几行（可留空：之后再自己在表里插行也行）。
    members: list[str] = Field(default_factory=list)
    #: 成员自己的参数。**不支持 derived**（派生的依赖图只覆盖 global / local）。
    variables: list[VariableDef] = Field(default_factory=list)

    @field_validator("sheet")
    @classmethod
    def _check_sheet(cls, value: str) -> str:
        return _valid_sheet_name(value)

    @field_validator("key")
    @classmethod
    def _check_key(cls, value: str) -> str:
        text = to_text(value).strip()
        if not text:
            raise ValueError("variables.group.key 不能为空（它要指向一个 local 变量名）")
        return text

    @field_validator("members", mode="before")
    @classmethod
    def _check_members(cls, value: Any) -> list[str]:
        if value is None:
            return []
        if isinstance(value, (str, bytes)) or not isinstance(value, (list, tuple)):
            raise ValueError("variables.group.members 必须是字符串列表（也可以留空，之后在表里插行）")
        names = [to_text(item).strip() for item in value]
        if any(not name for name in names):
            raise ValueError("variables.group.members 里有空名字")
        dupes = _duplicates(names)
        if dupes:
            raise ValueError(f"variables.group.members 名字重复: {', '.join(dupes)}")
        return names

    @model_validator(mode="after")
    def _check_group(self) -> GroupConfig:
        if not self.variables:
            raise ValueError("variables.group.variables 不能为空 —— 不想用第三层作用域就删掉整个 group 段")
        dupes = _duplicates([item.name for item in self.variables])
        if dupes:
            raise ValueError(f"group 变量名重复: {', '.join(dupes)}")
        for variable in self.variables:
            if variable.is_derived:
                raise ValueError(
                    f"group 变量 {variable.name!r} 用了 derived —— 派生参数的依赖图目前只覆盖 "
                    "global / local；请把它改成填写型，或挪到 local 里"
                )
        return self

    @property
    def names(self) -> list[str]:
        return [item.name for item in self.variables]


class VariablesConfig(BaseModel):
    """全局变量 + 局部变量 + （可选）成员表。"""

    model_config = ConfigDict(extra="forbid", populate_by_name=True)

    global_: list[VariableDef] = Field(default_factory=list, alias="global")
    local: list[VariableDef] = Field(default_factory=list)
    #: 第三层作用域：成员表（舱 / 设备）。见 :class:`GroupConfig`。
    group: GroupConfig | None = None

    @model_validator(mode="after")
    def _check_group_key(self) -> VariablesConfig:
        if self.group is None:
            return self
        if self.group.key not in self.local_names:
            raise ValueError(
                f"variables.group.key {self.group.key!r} 不是 local 变量 —— "
                f"它必须是「每个 Case 一列」的那种变量（当前 local: {', '.join(self.local_names) or '（空）'}）"
            )
        if self.group.sheet in {self.group.key}:
            raise ValueError("variables.group.sheet 与变量名冲突")
        clashes = sorted(set(self.group_names) & (set(self.global_names) | set(self.local_names)))
        if clashes:
            raise ValueError(
                f"group 变量与 global / local 重名: {', '.join(clashes)} —— "
                "同名会让人分不清用的是哪一个，请改名（第三层作用域的名字必须独立）"
            )
        return self

    @property
    def global_names(self) -> list[str]:
        return [item.name for item in self.global_]

    @property
    def local_names(self) -> list[str]:
        return [item.name for item in self.local]

    @property
    def group_variables(self) -> list[VariableDef]:
        return list(self.group.variables) if self.group else []

    @property
    def group_names(self) -> list[str]:
        return self.group.names if self.group else []

    @property
    def has_group(self) -> bool:
        return self.group is not None


# --------------------------------------------------------------------------- #
# Excel 结构
# --------------------------------------------------------------------------- #
class ExcelSheets(BaseModel):
    """工作表名称映射。"""

    model_config = ConfigDict(extra="forbid", populate_by_name=True)

    global_: str = Field("Global Parameter", alias="global")
    local: str = "Local Parameter"
    outputs: list[str] = Field(default_factory=lambda: ["Output"])

    @field_validator("global_", "local")
    @classmethod
    def _check_sheet_name(cls, value: str) -> str:
        return _valid_sheet_name(value)

    @field_validator("outputs")
    @classmethod
    def _check_outputs(cls, value: list[str]) -> list[str]:
        names = [to_text(item).strip() for item in value]
        if not names or any(not item for item in names):
            raise ValueError("excel.sheets.outputs 至少需要一个非空的工作表名")
        dupes = _duplicates(names)
        if dupes:
            raise ValueError(f"excel.sheets.outputs 中工作表名重复: {', '.join(dupes)}")
        return names


class ExcelConfig(BaseModel):
    """Excel 模板文件的整体配置。"""

    model_config = ConfigDict(extra="forbid")

    output: str = "template.xlsx"
    template_sheet: str | None = "Template"
    #: 工作簿里的"使用说明"表（放在第一张）。写清三步、命令、生成时间与指纹；
    #: 由 create_template 生成、write_results 刷新。设为 null 则不生成。
    howto_sheet: str | None = "HOWTO"
    # pydantic 的 default_factory 接受类本身；mypy 对它的签名判断过严，这里显式放行
    sheets: ExcelSheets = Field(default_factory=ExcelSheets)  # type: ignore[arg-type]
    #: **Local 表的工况排布**：``horizontal``（默认）= 一个工况一列；``vertical`` = 一个工况一行。
    #: 与每个模板自己的 ``direction``（**输出**排布）**互相独立** —— 输入竖着填、输出横着写都可以。
    #: 一行一个工况方便"整块粘贴"：有些软件里工况控制语句就是按行给的（见指南 §19）。
    local_direction: Direction = "horizontal"

    @field_validator("output")
    @classmethod
    def _check_output(cls, value: str) -> str:
        name = to_text(value).strip()
        if not name:
            raise ValueError("excel.output 不能为空")
        return name

    @field_validator("template_sheet", "howto_sheet")
    @classmethod
    def _check_optional_sheet_name(cls, value: str | None) -> str | None:
        if value is None or to_text(value).strip() == "":
            return None
        return _valid_sheet_name(value)


# --------------------------------------------------------------------------- #
# 模板定义
# --------------------------------------------------------------------------- #
class TemplateDef(BaseModel):
    """一个输出模板：源码 + 输出位置 + 布局方向。"""

    model_config = ConfigDict(extra="forbid")

    name: str
    description: str = ""
    output_sheet: str = "Output"
    start_cell: str = "B2"
    direction: Direction = "horizontal"
    write_case_headers: bool = True
    filename: str | None = None
    extension: str = ".txt"
    code: str | None = None
    template_file: str | None = None
    #: 输出引擎。**默认 ``excel``（写公式）** —— 本工具的用法是"生成一次工作簿，之后就用
    #: Excel 干活"：公式模式下改参数由 Excel 自己重算，不用再跑命令，工作簿**独立可用**。
    #: ``snapshot``（写文本快照）只在下面这些情况才需要，且必须**显式**声明：
    #:
    #: * 模板里有 ``{% for %}`` / 过滤器 / 多行 ``{% if %}`` / ``{% include %}``（公式模式表达不了）；
    #: * 想让导出的代码文件"改完参数自动同步"（公式模式的值只活在 Excel 里，
    #:   ``--outdir`` 仍然要走命令行）。
    engine: Engine = "excel"
    #: 可选：Jinja2 表达式，对每个 Case 的上下文求值；为假则该模板跳过这个 Case。
    #: 例如 ``case_filter: "kind == 'EXT'"``。用于"一本工作簿放两套规则"的场景。
    case_filter: str | None = None

    #: 声明这个模板（以及它的 ``template_file``）的文件所在目录。
    #: 只有 ``extends`` 合并进来的模板才需要它 —— 那些相对路径要相对**声明它的那个文件**解析，
    #: 而不是相对最终的项目 YAML。``None`` 表示"就是配置文件自己"，用 ``config.source_dir``。
    source_dir: Path | None = Field(default=None, exclude=True)

    @field_validator("name")
    @classmethod
    def _check_name(cls, value: str) -> str:
        name = to_text(value).strip()
        if not name:
            raise ValueError("模板名不能为空")
        return name

    @field_validator("description", mode="before")
    @classmethod
    def _description_to_str(cls, value: Any) -> str:
        return "" if value is None else to_text(value)

    @field_validator("case_filter", mode="before")
    @classmethod
    def _case_filter_to_str(cls, value: Any) -> str | None:
        if value is None:
            return None
        text = to_text(value).strip()
        return text or None

    @field_validator("extension")
    @classmethod
    def _check_extension(cls, value: str) -> str:
        ext = to_text(value).strip()
        if not ext:
            return ""
        return ext if ext.startswith(".") else f".{ext}"

    @field_validator("start_cell")
    @classmethod
    def _check_start_cell(cls, value: str) -> str:
        try:
            parse_cell(value)
        except ValueError as exc:  # pragma: no cover - 信息透传
            raise ValueError(str(exc)) from exc
        return to_text(value).strip().upper()

    @model_validator(mode="after")
    def _check_source(self) -> TemplateDef:
        if not (self.code and self.code.strip()) and not self.template_file:
            raise ValueError("必须提供 code（内联模板）或 template_file（外部模板文件）之一")
        if self.filename and not self.filename.strip():
            raise ValueError("filename 不能为空字符串（如需默认命名请删除该字段）")
        return self

    @property
    def source_code(self) -> str:
        """内联模板源码（``template_file`` 场景下为空字符串）。"""
        return self.code or ""


# --------------------------------------------------------------------------- #
# 顶层配置
# --------------------------------------------------------------------------- #
class ProjectConfig(BaseModel):
    """整个项目的配置根对象。"""

    model_config = ConfigDict(extra="forbid")

    version: int = 1
    excel: ExcelConfig = Field(default_factory=ExcelConfig)
    variables: VariablesConfig = Field(default_factory=VariablesConfig)
    templates: list[TemplateDef]

    #: **跨变量校验**：对**每个 Case** 求值的 Jinja 表达式，结果为假就报错（见指南 §3.6）。
    #: 单变量约束（``min`` / ``max`` / ``choices`` / ``pattern``）拦不住"吃水不能超过型深"
    #: 这类**组合**错误，这一层补的就是它。写法与 ``case_filter`` 一样：裸变量是组合值，
    #: **数值比较请写 ``.value``** —— ``asserts: ["draft.value <= d_tank.value"]``。
    asserts: list[str] = Field(default_factory=list)

    #: 配置文件所在目录（加载时自动填充，用于解析相对路径的 template_file）。
    source_dir: Path | None = Field(default=None, exclude=True)
    #: 配置文件本身的绝对路径（加载时自动填充）。生成的运行脚本与 HOWTO 表要用它。
    config_path: Path | None = Field(default=None, exclude=True)
    #: ``extends`` 指向的其他 YAML（相对本文件解析）。加载时会先合并它们，见 §16。
    #: 加载后这里保留的是**本文件写的原始列表**，方便调用方知道配置由哪些文件组成。
    extends: list[str] = Field(default_factory=list, exclude=True)
    #: 合并 ``extends`` 时的告警（例如同名变量的 ``default`` 不一致），由 CLI 打印出来。
    load_warnings: list[str] = Field(default_factory=list, exclude=True)
    #: 工作簿旁边**是否真的有一键刷新脚本**（由 create_template / write_results 填）。
    #: HOWTO 表据此决定要不要写「懒得开终端就双击那个脚本」。
    scripts_enabled: bool = Field(default=False, exclude=True)

    @field_validator("asserts", mode="before")
    @classmethod
    def _check_asserts(cls, value: Any) -> list[str]:
        if value is None:
            return []
        if isinstance(value, (str, bytes)) or not isinstance(value, (list, tuple)):
            raise ValueError('asserts 必须是表达式列表，例如 asserts: ["draft.value <= d_tank.value"]')
        out: list[str] = []
        for index, item in enumerate(value, start=1):
            text = to_text(item).strip()
            if not text:
                raise ValueError(f"asserts 第 {index} 条是空的（不想要就删掉它）")
            out.append(text)
        return out

    @model_validator(mode="after")
    def _validate_config(self) -> ProjectConfig:
        if self.version < 1:
            raise ValueError("version 必须 >= 1")
        if not self.templates:
            raise ValueError("至少需要定义一个 template")

        template_names = [item.name for item in self.templates]
        dupes = _duplicates(template_names)
        if dupes:
            raise ValueError(f"template 名称重复: {', '.join(dupes)}")

        for scope, names in (
            ("global", self.variables.global_names),
            ("local", self.variables.local_names),
        ):
            dupes = _duplicates(names)
            if dupes:
                raise ValueError(f"{scope} 变量名重复: {', '.join(dupes)}")

        outputs = set(self.excel.sheets.outputs)
        for template in self.templates:
            if template.output_sheet not in outputs:
                raise ValueError(
                    f"template {template.name!r} 的 output_sheet {template.output_sheet!r} "
                    f"未在 excel.sheets.outputs 中声明（当前: {', '.join(self.excel.sheets.outputs)}）"
                )

        used = {self.excel.sheets.global_, self.excel.sheets.local, *outputs}
        if self.group is not None:
            if self.group.sheet in used:
                raise ValueError(
                    f"variables.group.sheet {self.group.sheet!r} 与 Global / Local / Output 表名冲突，请改名"
                )
            used.add(self.group.sheet)
        for label, sheet_name in (
            ("excel.template_sheet", self.excel.template_sheet),
            ("excel.howto_sheet", self.excel.howto_sheet),
        ):
            if sheet_name and sheet_name in used:
                raise ValueError(f"{label} {sheet_name!r} 与 Global / Local / Output 表名冲突，请改名或设为 null")
        if self.excel.template_sheet and self.excel.howto_sheet and self.excel.template_sheet == self.excel.howto_sheet:
            raise ValueError(
                f"excel.template_sheet 与 excel.howto_sheet 不能同名（都是 {self.excel.template_sheet!r}）"
            )
        return self

    # -- 便捷访问 ---------------------------------------------------------- #
    @property
    def global_variables(self) -> list[VariableDef]:
        return list(self.variables.global_)

    @property
    def local_variables(self) -> list[VariableDef]:
        return list(self.variables.local)

    @property
    def global_names(self) -> list[str]:
        return self.variables.global_names

    @property
    def local_names(self) -> list[str]:
        return self.variables.local_names

    @property
    def group(self):
        """第三层作用域的成员表配置（没配就是 ``None``）。"""
        return self.variables.group

    @property
    def group_variables(self) -> list[VariableDef]:
        return self.variables.group_variables

    @property
    def group_names(self) -> list[str]:
        return self.variables.group_names

    @property
    def has_group(self) -> bool:
        return self.variables.has_group

    @property
    def defined_names(self) -> set[str]:
        """所有已定义变量名（用于校验模板引用的变量是否存在）。"""
        return set(self.global_names) | set(self.local_names) | set(self.group_names) | set(RESERVED_NAMES)

    def template_by_name(self, name: str) -> TemplateDef | None:
        for template in self.templates:
            if template.name == name:
                return template
        return None


class _StrictLoader(yaml.SafeLoader):
    """SafeLoader + 同一映射里的重复键检测。

    PyYAML 默认对重复键**不报错**（后者覆盖前者），于是 YAML 里写两个 ``variables:``
    会让第一个整块静默消失，错误要到渲染期才以"变量缺失"的形式暴露。
    """


def _construct_mapping_strict(loader: yaml.SafeLoader, node: yaml.MappingNode, deep: bool = False) -> dict:
    mapping: dict = {}
    for key_node, value_node in node.value:
        key = loader.construct_object(key_node, deep=deep)
        try:
            duplicated = key in mapping
        except TypeError as exc:  # 不可哈希的复杂键
            raise ConfigError(f"YAML 键非法（不可哈希）: 第 {key_node.start_mark.line + 1} 行") from exc
        if duplicated:
            raise ConfigError(
                f"YAML 键重复: {key!r}（第 {key_node.start_mark.line + 1} 行）"
                "—— 同一个映射里同一个键只能出现一次（重复会让前一份被静默覆盖）"
            )
        mapping[key] = loader.construct_object(value_node, deep=deep)
    return mapping


_StrictLoader.add_constructor(yaml.resolver.BaseResolver.DEFAULT_MAPPING_TAG, _construct_mapping_strict)


def _read_yaml_mapping(path: Path) -> dict:
    """读一个 YAML 文件并要求根节点是映射；失败时抛带位置的 :class:`ConfigError`。"""
    if not path.exists():
        raise ConfigError(f"配置文件不存在: {path}")
    if path.is_dir():
        raise ConfigError(f"配置路径是目录而不是文件: {path}")
    try:
        raw_text = path.read_text(encoding="utf-8")
    except OSError as exc:
        raise ConfigError(f"无法读取配置文件 {path}: {exc}") from exc
    try:
        raw = yaml.load(raw_text, Loader=_StrictLoader)
    except ConfigError:
        raise
    except yaml.YAMLError as exc:
        mark = getattr(exc, "problem_mark", None)
        where = f"（第 {mark.line + 1} 行，第 {mark.column + 1} 列）" if mark else ""
        raise ConfigError(f"YAML 解析失败 {path}{where}:\n  {exc}") from exc
    if raw is None:
        raise ConfigError(f"配置文件内容为空: {path}")
    if not isinstance(raw, dict):
        raise ConfigError(f"配置文件根节点必须是映射（mapping），实际是 {type(raw).__name__}: {path}")
    return raw


#: extends 合并时，同名变量之间**必须逐字一致**的字段 —— 它们决定生成出来的文本。
#: 和 ``abs_fpi/compose.py`` 的 STRICT_FIELDS 是同一套判据，只是把取值约束也纳入了。
EXTENDS_STRICT_FIELDS: tuple[str, ...] = (
    "prefix",
    "suffix",
    "type",
    "derived",
    "min",
    "max",
    "choices",
    "pattern",
)
#: 这些字段不一致只告警（不同规则集的示例工况本来就不同），保留先出现的那个。
EXTENDS_LOOSE_FIELDS: tuple[str, ...] = ("default", "description")


def _merge_variable(
    existing: dict, incoming: dict, *, scope: str, origin: Path, first_seen: Path, warnings: list[str]
) -> None:
    name = incoming.get("name")
    for key in EXTENDS_STRICT_FIELDS:
        before, after = existing.get(key), incoming.get(key)
        if before != after:
            raise ConfigError(
                f"变量 {name!r} 在 {scope} 里被 {first_seen} 与 {origin} 定义成不同的 {key}："
                f"{before!r} vs {after!r}。这些字段决定生成出来的文本，必须先统一。"
            )
    for key in EXTENDS_LOOSE_FIELDS:
        if existing.get(key) != incoming.get(key):
            warnings.append(
                f"变量 {name!r}（{scope}）的 {key} 在两处不一致：保留 {first_seen} 的 "
                f"{existing.get(key)!r}，忽略 {origin} 的 {incoming.get(key)!r}"
            )


def _merge_into(
    merged: dict,
    origins: dict[str, Path],
    raw: dict,
    *,
    origin: Path,
    is_root: bool,
    warnings: list[str],
) -> None:
    """把一个文件的 ``variables`` / ``templates`` 合并进累积结果。"""
    if not is_root and raw.get("excel"):
        warnings.append(f"忽略 {origin} 里的 excel 配置（工作簿布局以根配置文件为准）")
    variables_raw = raw.get("variables")
    if not is_root and isinstance(variables_raw, dict) and variables_raw.get("group"):
        warnings.append(f"忽略 {origin} 里的 variables.group（成员表只有一份，以根配置文件为准）")

    variables = raw.get("variables") or {}
    if not isinstance(variables, dict):
        variables = {}
    if is_root and isinstance(variables.get("group"), dict):
        merged["variables"]["group"] = dict(variables["group"])
    for scope in ("global", "local"):
        for item in variables.get(scope) or []:
            if not isinstance(item, dict) or "name" not in item:
                merged["variables"][scope].append(item)  # 交给 pydantic 报错，信息更统一
                continue
            name = item["name"]
            existing = next(
                (v for v in merged["variables"][scope] if isinstance(v, dict) and v.get("name") == name),
                None,
            )
            if existing is None:
                merged["variables"][scope].append(dict(item))
                origins.setdefault(f"{scope}:{name}", origin)
                continue
            _merge_variable(
                existing,
                item,
                scope=scope,
                origin=origin,
                first_seen=origins.get(f"{scope}:{name}", origin),
                warnings=warnings,
            )

    # 跨变量校验都保留下来（顺序：先被 extends 的在前），互不覆盖
    merged["asserts"].extend(raw.get("asserts") or [])

    for template in raw.get("templates") or []:
        if not isinstance(template, dict):
            merged["templates"].append(template)
            continue
        name = template.get("name")
        same = next((t for t in merged["templates"] if isinstance(t, dict) and t.get("name") == name), None)
        if same is not None:
            if yaml.safe_dump(same, sort_keys=True, allow_unicode=True) != yaml.safe_dump(
                template, sort_keys=True, allow_unicode=True
            ):
                raise ConfigError(
                    f"模板 {name!r} 在 {origins.get(f'template:{name}', origin)} 与 {origin} "
                    "里都定义了，但内容不同 —— 模板名必须唯一，请改名或统一内容"
                )
            continue
        merged["templates"].append(dict(template))
        origins.setdefault(f"template:{name}", origin)


def _load_with_extends(
    config_path: Path, *, chain: tuple[Path, ...] = (), warnings: list[str]
) -> tuple[dict, dict[str, Path]]:
    """递归展开 ``extends``，返回合并后的 raw dict 与"每项来自哪个目录"的索引。"""
    resolved = config_path.resolve()
    if resolved in chain:
        loop = " → ".join(str(p) for p in (*chain, resolved))
        raise ConfigError(f"extends 出现循环引用: {loop}")

    origin_dir = resolved.parent
    raw = _read_yaml_mapping(config_path)
    parents = raw.get("extends") or []
    if not isinstance(parents, list) or not all(isinstance(p, str) for p in parents):
        raise ConfigError(f"extends 必须是文件路径的列表，例如 extends: [rules/a.yaml, rules/b.yaml]: {config_path}")

    merged: dict = {"variables": {"global": [], "local": []}, "templates": [], "asserts": []}
    origins: dict[str, Path] = {}

    for relative in parents:
        parent_path = (origin_dir / relative).resolve()
        if not parent_path.exists():
            raise ConfigError(f"extends 指向的文件不存在: {parent_path}（写在 {config_path}）")
        parent_raw, parent_origins = _load_with_extends(parent_path, chain=(*chain, resolved), warnings=warnings)
        _merge_into(
            merged,
            origins,
            parent_raw,
            origin=parent_path.parent,
            is_root=False,
            warnings=warnings,
        )
        origins.update(parent_origins)

    _merge_into(merged, origins, raw, origin=origin_dir, is_root=True, warnings=warnings)
    merged["version"] = raw.get("version", 1)
    merged["excel"] = raw.get("excel", {})
    return merged, origins


def load_config(path: str | Path) -> ProjectConfig:
    """读取并校验 YAML 配置，失败时抛出带清晰提示的 :class:`ConfigError`。

    配置里写了 ``extends: [a.yaml, b.yaml]`` 时会先递归合并那些文件（见指南 §16）：
    同名变量的 ``prefix`` / ``suffix`` / ``type`` / 取值约束必须一致，``default`` 与
    ``description`` 不一致只告警；同名模板内容必须一致。
    """
    config_path = Path(path)
    warnings: list[str] = []

    probe = _read_yaml_mapping(config_path)

    if isinstance(probe.get("extends"), list) and probe["extends"]:
        raw, origins = _load_with_extends(config_path, warnings=warnings)
    else:
        raw, origins = probe, {}

    try:
        config = ProjectConfig.model_validate(raw)
    except ValidationError as exc:
        raise ConfigError(f"配置校验失败 {config_path}:\n{format_validation_error(exc)}") from exc

    config.source_dir = config_path.resolve().parent
    config.config_path = config_path.resolve()
    config.load_warnings = warnings
    config.extends = [str(item) for item in (probe.get("extends") or [])]
    # 每个模板记住"声明它的那个文件在哪"，这样 extends 进来的 template_file 相对路径仍然解析得对
    for template in config.templates:
        origin = origins.get(f"template:{template.name}")
        if origin is not None and origin != config.source_dir:
            template.source_dir = origin
    return config


# --------------------------------------------------------------------------- #
# 运行期数据
# --------------------------------------------------------------------------- #
@dataclass(frozen=True)
class CaseData:
    """Local Parameter 表里一个 Case 的取值。

    位置二选一：``horizontal`` 布局一个工况占**一列**（``column``），
    ``vertical`` 布局一个工况占**一行**（``row``）。
    """

    name: str
    #: 横向布局：这个 Case 的列号。
    column: int | None = None
    values: dict[str, VarValue] = field(default_factory=dict)
    #: 这一列/行在表里是否**至少填过一个**局部变量的值；
    #: 全空时所有变量都会回落 YAML ``default``（新插的空列/空行就此"悄悄"落进某个规则集）。
    explicit_values: bool = True
    #: 纵向布局：这个 Case 的行号。
    row: int | None = None

    @property
    def index(self) -> int:
        """Case 序号（Case1 -> 1）。"""
        if self.row is not None:
            return self.row - 1
        assert self.column is not None
        return self.column - FIRST_CASE_COLUMN + 1

    @property
    def where(self) -> str:
        """给人看的定位（横向说"第几列"，纵向说"第几行"）。"""
        if self.row is not None:
            return f"第 {self.row} 行"
        assert self.column is not None
        return f"第 {column_index_to_letter(self.column)} 列"


@dataclass(frozen=True)
class RenderResult:
    """某个模板在某个 Case 下的渲染结果。"""

    template_name: str
    case_name: str
    text: str
    #: 渲染这一份结果时用的上下文（变量 -> :class:`VarValue`）。
    #: 导出文件名（``template.filename``）里可以用到任意参数，比如只用来排序的 ``seq``。
    context: Mapping[str, Any] = field(default_factory=dict)

    @property
    def lines(self) -> list[str]:
        """渲染结果的行列表（不含结尾空行），用于写入 Excel。"""
        return split_lines(self.text)

    @property
    def line_count(self) -> int:
        return len(self.lines)
