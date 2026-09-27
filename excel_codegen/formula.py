"""公式引擎：把"纯替换"型模板编译成 Excel 公式（``engine: excel``）。

设计边界（A 档：只做占位符替换）
--------------------------------
支持的占位符（仅这些）：

=========================  ==========================================================
``{{ x }}``                ``Prefix & Value & Suffix``（与快照模式的组合值一致）
``{{ x.value }}``          取值列（原始值）
``{{ x.text }}``           同 ``.value``（对应 Python 侧 ``VarValue.text``）
``{{ x.prefix }}``         Prefix 列
``{{ x.suffix }}``         Suffix 列
``{{ case_name }}``        当前 Case 名（取 Local 表表头）
``{{ template_name }}``    模板名（常量）
=========================  ==========================================================

出现 ``{%`` / ``{#`` / ``|`` 过滤器 / 运算表达式 / 函数调用等一律**报错并指出行号内容**，
提示该模板继续用 ``engine: snapshot`` —— 控制流没法用"单元格引用"表达。

两条路径的一致性
----------------
* 变量用 ``INDEX/MATCH`` 按**变量名**定位，用户在 Global / Local 表里插行删行都不会指错。
* 空单元格回落 YAML ``default``：公式写成 ``IF(ref="", default, ref)``，与 Python 侧一致。
* ``type: int`` / ``float`` 用 ``TEXT()`` 规范化，避免出现 ``340.0`` 这种尾巴。
* ``{{ x }}`` 的组合顺序与 Python 侧一致：Prefix + Value + Suffix。

已知差异（见 docs/template_guide.md）
------------------------------------
* ``type: bool`` 在 Excel 里是 ``TRUE`` / ``FALSE``，Python 侧渲染成 ``true`` / ``false``。
* ``TEXT()`` 的格式串受区域设置影响。
* 公式上限 8192 字符/格，超过会报错并提示拆行或改用快照模式。
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any, Sequence

from .models import ProjectConfig, TemplateDef, VariableDef
from .utils import RenderError, column_index_to_letter, to_text

__all__ = [
    "LONG_FORMULA_WARN",
    "MAX_FORMULA_CHARS",
    "FormulaError",
    "compile_formulas",
    "compile_line",
    "excel_literal",
    "guarded_lookup",
    "lookup_expr",
]

#: Excel 单个公式的字符上限（留一点余量）。
MAX_FORMULA_CHARS = 8000

#: 公式长度**警告**阈值：一个 ``{{ x }}`` 大约展开 300–400 字符，
#: 超过这个量级的行虽然能编译，但人已经读不懂了，建议拆行。
LONG_FORMULA_WARN = 3000

_PLACEHOLDER_RE = re.compile(r"\{\{(.*?)\}\}", re.DOTALL)
_EXPR_RE = re.compile(r"^([A-Za-z_][A-Za-z0-9_]*)(?:\.([A-Za-z_][A-Za-z0-9_]*))?$")

#: 与 excel_io 的列约定一致
_GLOBAL_COLUMNS = {"value": 2, "prefix": 4, "suffix": 5}
_LOCAL_COLUMNS = {"prefix": 3, "suffix": 4}  # value 列随 Case 变化
_VALUE_ALIASES = frozenset({"value", "text"})
#: 数值形态**不用 TEXT()**：Excel 的 TEXT() 会把二进制尾巴原样打出来
#: （``TEXT(20.559,"0.###############")`` 是 ``20.559000000000001``），
#: 而单元格/拼接的**隐式转换**走的是 General 规则 = **最多 15 位有效数字** ——
#: 这正是 :func:`utils.to_text` 的规则。两边因此逐字一致（见 docs/template_guide.md §14.2）。


class FormulaError(RenderError):
    """模板无法编译成 Excel 公式（超出"纯替换"子集，或变量信息不足）。"""


# --------------------------------------------------------------------------- #
# 基础片段
# --------------------------------------------------------------------------- #
def _quote_sheet(sheet: str) -> str:
    return "'" + str(sheet).replace("'", "''") + "'"


def _quote_text(text: Any) -> str:
    return '"' + to_text(text).replace('"', '""') + '"'


def excel_literal(value: Any) -> str:
    """把 YAML 默认值写成 Excel 字面量。"""
    if value is None:
        return '""'
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, (int, float)):
        return repr(value)
    text = to_text(value)
    return '""' if text == "" else _quote_text(text)


def lookup_expr(sheet: str, variable: str, column: int, *, absolute: bool) -> str:
    """``INDEX('Local Parameter'!E:E,MATCH("port",'Local Parameter'!$A:$A,0))``。

    ``absolute=False`` 时列标不加 ``$`` —— 把输出单元格向右拖，Case 列会跟着走。
    """
    letter = column_index_to_letter(column)
    anchor = f"${letter}:${letter}" if absolute else f"{letter}:{letter}"
    quoted = _quote_sheet(sheet)
    return f"INDEX({quoted}!{anchor},MATCH({_quote_text(variable)},{quoted}!$A:$A,0))"


def guarded_lookup(
    sheet: str,
    variable: str,
    column: int,
    default: Any,
    *,
    absolute: bool = True,
) -> str:
    """``lookup_expr`` + "空单元格回落 YAML 默认值"的保护。

    这是**参数表里派生参数格**用的引用形态（``derived`` 求值也走同一套语义）。
    """
    return _with_default(lookup_expr(sheet, variable, column, absolute=absolute), default)


def _with_default(expr: str, default: Any) -> str:
    """空单元格回落 YAML 默认值（与 Python 侧一致）。

    ⚠ 必须用 ``ISBLANK`` 而不是 ``expr=""``：``INDEX`` 对**空单元格**返回数值 ``0``
    （不是空串），用 ``=""`` 判断会把空前缀变成 ``0``，生成 ``0STM32F1030`` 这种垃圾。
    另外再补一个 ``=""`` 分支，覆盖"长度为零的文本"。
    """
    if to_text(default) == "":
        return f'IF(ISBLANK({expr}),"",{expr})'
    literal = excel_literal(default)
    return f'IF(ISBLANK({expr}),{literal},IF({expr}="",{literal},{expr}))'


def _typed(expr: str, definition: VariableDef | None) -> str:
    """数值规范化：交给 Excel 的隐式转换（General = 最多 15 位有效数字）。

    与 :func:`utils.to_text` 的规则一致，所以 ``{{ x }}`` 在 Excel 里与导出文件里
    是同一个字符串；``type`` 只影响 Python 侧的取值转换，不再往公式里塞 ``TEXT()``。
    """
    return expr


@dataclass(frozen=True)
class _VarRef:
    """一个变量在公式里的定位信息。"""

    sheet: str
    name: str
    definition: VariableDef | None
    value_column: int
    prefix_column: int
    suffix_column: int
    value_column_is_relative: bool

    def _lookup(self, column: int, *, absolute: bool) -> str:
        return lookup_expr(self.sheet, self.name, column, absolute=absolute)

    def value(self) -> str:
        base = self._lookup(self.value_column, absolute=not self.value_column_is_relative)
        default = self.definition.default if self.definition else ""
        return _typed(_with_default(base, default), self.definition)

    def decor(self, which: str) -> str:
        column = self.prefix_column if which == "prefix" else self.suffix_column
        default = getattr(self.definition, which, "") if self.definition else ""
        return _with_default(self._lookup(column, absolute=True), default)

    def combined(self) -> str:
        return f"{self.decor('prefix')}&{self.value()}&{self.decor('suffix')}"


# --------------------------------------------------------------------------- #
# 编译器
# --------------------------------------------------------------------------- #
class _Compiler:
    """把某个模板的行编译成公式；缓存变量定位信息。"""

    def __init__(
        self,
        config: ProjectConfig,
        template_name: str,
        *,
        relative_case_column: bool = True,
    ) -> None:
        self.config = config
        self.template_name = template_name
        #: 横向布局：一个 Case 一列，向右拖公式应该跟着换 Case（相对列）。
        #: 纵向布局：向右拖是换"同一 Case 的下一行"，Case 列必须锁死（绝对列）。
        self.relative_case_column = relative_case_column
        self._globals = {item.name: item for item in config.global_variables}
        self._locals = {item.name: item for item in config.local_variables}
        self._cache: dict[tuple[str, int], _VarRef] = {}

    # -- 变量 -------------------------------------------------------------- #
    def resolve(self, name: str, case_column: int) -> _VarRef:
        key = (name, case_column)
        if key in self._cache:
            return self._cache[key]

        if name in self._globals:
            ref = _VarRef(
                sheet=self.config.excel.sheets.global_,
                name=name,
                definition=self._globals[name],
                value_column=_GLOBAL_COLUMNS["value"],
                prefix_column=_GLOBAL_COLUMNS["prefix"],
                suffix_column=_GLOBAL_COLUMNS["suffix"],
                value_column_is_relative=False,
            )
        elif name in self._locals:
            ref = _VarRef(
                sheet=self.config.excel.sheets.local,
                name=name,
                definition=self._locals[name],
                value_column=case_column,
                prefix_column=_LOCAL_COLUMNS["prefix"],
                suffix_column=_LOCAL_COLUMNS["suffix"],
                value_column_is_relative=self.relative_case_column,
            )
        else:
            raise FormulaError(
                f"模板 {self.template_name!r} 的公式模式引用了未在 YAML 中定义的变量 {name!r}："
                "公式需要知道它是 global 还是 local，请把它写进 variables"
                "（快照模式允许表里临时出现的变量，公式模式不允许）"
            )
        self._cache[key] = ref
        return ref

    # -- 表达式 ------------------------------------------------------------ #
    def expression(self, raw: str, *, case_column: int, line: str) -> str:
        expr = raw.strip()
        if "|" in expr:
            raise FormulaError(
                f"模板 {self.template_name!r} 的公式模式不支持过滤器：{{{{ {expr} }}}}；"
                "请把该模板改回 engine: snapshot，或把过滤器换成等价的单元格拼接"
            )
        if expr == "template_name":
            return _quote_text(self.template_name)
        if expr == "case_name":
            sheet = _quote_sheet(self.config.excel.sheets.local)
            column = column_index_to_letter(case_column)
            anchor = f"{column}$1" if self.relative_case_column else f"${column}$1"
            return f"{sheet}!{anchor}"

        match = _EXPR_RE.match(expr)
        if not match:
            raise FormulaError(
                f"模板 {self.template_name!r} 的公式模式只支持 "
                "变量 / 变量.value / .text / .prefix / .suffix / case_name / template_name，"
                f"这一行出现了 {expr!r}（行内容：{line.strip()!r}）"
            )

        name, attribute = match.group(1), match.group(2)
        ref = self.resolve(name, case_column)
        if attribute is None:
            return ref.combined()
        if attribute in _VALUE_ALIASES:
            return ref.value()
        if attribute == "prefix":
            return ref.decor("prefix")
        if attribute == "suffix":
            return ref.decor("suffix")
        raise FormulaError(
            f"模板 {self.template_name!r} 的公式模式不支持属性 .{attribute}（变量 {name!r}）："
            "只支持 .value / .text / .prefix / .suffix"
        )

    # -- 行 ---------------------------------------------------------------- #
    def line(self, line: str, *, case_column: int) -> str:
        if "{%" in line or "{#" in line:
            raise FormulaError(
                f"模板 {self.template_name!r} 的公式模式不支持 Jinja 控制流/注释（含 {{% 或 {{#）；"
                f"请把该模板改回 engine: snapshot；行内容：{line.strip()!r}"
            )

        pieces: list[str] = []
        cursor = 0
        for match in _PLACEHOLDER_RE.finditer(line):
            literal = line[cursor:match.start()]
            if literal:
                pieces.append(_quote_text(literal))
            try:
                pieces.append(
                    self.expression(match.group(1), case_column=case_column, line=line)
                )
            except FormulaError as exc:
                # 统一带上出错的行，方便在几十行的模板里定位
                raise FormulaError(f"{exc}；行内容：{line.strip()!r}") from exc
            cursor = match.end()
        tail = line[cursor:]
        if tail:
            pieces.append(_quote_text(tail))
        if cursor == 0:  # 整行没有占位符：常量文本
            pieces = [_quote_text(line)]
        elif "{{" in line[cursor:] or "}}" in line[cursor:]:
            raise FormulaError(
                f"模板 {self.template_name!r} 的占位符没有闭合：{line.strip()!r}"
            )

        formula = "=" + "&".join(pieces)
        if len(formula) > MAX_FORMULA_CHARS:
            raise FormulaError(
                f"模板 {self.template_name!r} 编译出的公式过长（{len(formula)} > "
                f"{MAX_FORMULA_CHARS} 字符）：{line.strip()[:60]!r}…；请拆成多行或改用 engine: snapshot"
            )
        return formula


# --------------------------------------------------------------------------- #
# 对外 API
# --------------------------------------------------------------------------- #
def compile_line(
    line: str,
    *,
    config: ProjectConfig,
    template_name: str,
    case_column: int,
    relative_case_column: bool = True,
) -> str:
    """编译单独一行（主要给测试用）。"""
    return _Compiler(
        config, template_name, relative_case_column=relative_case_column
    ).line(line, case_column=case_column)


def compile_formulas(
    template: TemplateDef,
    config: ProjectConfig,
    *,
    case_columns: Sequence[int],
    source: str | None = None,
) -> list[list[str]]:
    """按 Case 编译整个模板：``[case][line] = 公式``。

    :param source: 外部模板文件（``template_file``）的源码；不给则用内联 ``code``。
    """
    text = template.source_code if source is None else source
    if not text.strip():
        raise FormulaError(
            f"模板 {template.name!r} 没有可编译的源码（内联 code 为空且未提供 template_file 内容）"
        )
    compiler = _Compiler(
        config,
        template.name,
        relative_case_column=template.direction == "horizontal",
    )
    lines = text.splitlines()
    return [
        [compiler.line(line, case_column=column) for line in lines]
        for column in case_columns
    ]
