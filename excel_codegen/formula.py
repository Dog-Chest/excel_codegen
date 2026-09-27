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

出现 ``|`` 过滤器 / 运算表达式 / 函数调用等一律**报错并指出行号内容**，
提示该模板继续用 ``engine: snapshot``。

**行内 ``{% if %}``**（0.6.0 起）
--------------------------------
``{% if 条件 %}A{% else %}B{% endif %}`` 可以写，条件是**比较或逻辑表达式**，
会被翻译成 Excel 的 ``IF(...)``。但有一条硬约束：**必须整段写在同一行内** ——
公式模式是"一行模板 → 一个单元格"，跨行的分支会改变行数，没法映射到固定单元格。
``{% for %}`` / ``{% set %}`` 等仍然报错。

条件里的**变量名指它的取值**（相当于 Python 侧的 ``VarValue.value``），不是
Prefix+Value+Suffix 的组合值 —— 这样 ``draft > 20`` 才是数值比较。
``{{ }}`` 里那套组合值语义不变。

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

from jinja2 import TemplateSyntaxError, nodes

from .jinja_env import build_environment
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

#: 一次匹配出 ``{{ 表达式 }}`` / ``{% 标签 %}`` / ``{# 注释 #}``
_TAG_RE = re.compile(
    r"\{\{(?P<expr>.*?)\}\}|\{%(?P<tag>.*?)%\}|\{#(?P<comment>.*?)#\}", re.DOTALL
)
_EXPR_RE = re.compile(r"^([A-Za-z_][A-Za-z0-9_]*)(?:\.([A-Za-z_][A-Za-z0-9_]*))?$")


@dataclass(frozen=True)
class _Token:
    """模板一行的词法单元：``text``（字面文本）/ ``expr``（``{{ }}``）/ ``tag``（``{% %}``）/ ``comment``。"""

    kind: str
    value: str

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
        tokens = self._tokenize(line)
        body = self._compile_tokens(tokens, case_column=case_column, line=line)
        formula = "=" + body
        if len(formula) > MAX_FORMULA_CHARS:
            raise FormulaError(
                f"模板 {self.template_name!r} 编译出的公式过长（{len(formula)} > "
                f"{MAX_FORMULA_CHARS} 字符）：{line.strip()[:60]!r}…；请拆成多行或改用 engine: snapshot"
            )
        return formula

    # -- 分词 -------------------------------------------------------------- #
    def _tokenize(self, line: str) -> list[_Token]:
        tokens: list[_Token] = []
        cursor = 0
        for match in _TAG_RE.finditer(line):
            if match.start() > cursor:
                tokens.append(_Token("text", line[cursor:match.start()]))
            if match.group("expr") is not None:
                tokens.append(_Token("expr", match.group("expr")))
            elif match.group("tag") is not None:
                tokens.append(_Token("tag", match.group("tag").strip()))
            cursor = match.end()
        if cursor < len(line):
            tokens.append(_Token("text", line[cursor:]))

        for token in tokens:
            if token.kind == "text" and any(
                mark in token.value for mark in ("{{", "}}", "{%", "%}", "{#", "#}")
            ):
                raise FormulaError(
                    f"模板 {self.template_name!r} 的标记没有闭合：{token.value.strip()[:40]!r}；"
                    f"行内容：{line.strip()!r}"
                )
        return tokens

    # -- 条件（{% if %}）---------------------------------------------------- #
    #: 条件里可以直接写的名字（它们没有"取值/前缀/后缀"之分）
    _CONDITION_SPECIAL_NAMES = frozenset({"case_name", "template_name"})

    def condition(self, source: str, *, case_column: int, line: str) -> str:
        """把 ``{% if %}`` 的条件翻译成 Excel 的布尔表达式。

        **条件里的变量必须写属性**（``x.value`` / ``x.text`` / ``x.prefix`` / ``x.suffix``）。
        这不是洁癖：快照模式里裸变量是 ``VarValue`` 对象，拿它跟数字比会直接报 TypeError，
        而 ``x.value`` 在两边都是一个"纯值" —— 只有这种写法能让两种引擎逐字一致。
        规则与 ``case_filter`` 一致（那里也是"要数值比较请用 .value"）。
        """
        from .derived import DerivedError, to_excel  # 延迟导入：derived 依赖 formula，避免循环

        expression = source.strip()
        if not expression:
            raise FormulaError(
                f"模板 {self.template_name!r} 的 {{{{ if }}}} 后面没有条件；行内容：{line.strip()!r}"
            )

        env = build_environment()
        try:
            ast = env.parse("{{ " + expression + " }}")
        except TemplateSyntaxError as exc:
            raise FormulaError(
                f"模板 {self.template_name!r} 的 {{{{ if }}}} 条件语法错误：{exc.message}；"
                f"条件：{expression!r}"
            ) from exc
        body = getattr(ast, "body", [])
        if len(body) != 1 or not isinstance(body[0], nodes.Output) or len(body[0].nodes) != 1:
            raise FormulaError(
                f"模板 {self.template_name!r} 的 {{{{ if }}}} 条件必须是单个表达式：{expression!r}"
            )
        node = body[0].nodes[0]
        self._require_attribute_access(node, expression, line)

        def resolve(name: str, attribute: str | None = None) -> str:
            return self._condition_ref(name, attribute, case_column)

        try:
            excel = to_excel(
                expression,
                name=f"{self.template_name} 的 {{{{ if }}}} 条件",
                resolve=resolve,
                env=env,
                condition=True,
            )
        except DerivedError as exc:
            raise FormulaError(f"{exc}；行内容：{line.strip()!r}") from exc

        if isinstance(node, (nodes.Compare, nodes.And, nodes.Or, nodes.Not)):
            return excel
        if isinstance(node, (nodes.Name, nodes.Getattr)):
            # 按真假判断：数值比 0，其余比空串（与 Python 侧 bool(value) 的直觉一致）
            return f"({excel}<>{'0' if self._is_numeric_node(node) else '\"\"'})"
        if isinstance(node, nodes.Const):
            numeric = isinstance(node.value, (int, float)) and not isinstance(node.value, bool)
            return f"({excel}<>{'0' if numeric else '\"\"'})"
        raise FormulaError(
            f"模板 {self.template_name!r} 的 {{{{ if }}}} 条件必须是比较或逻辑表达式"
            f"（例如 draft.value > 20、kind.value == \"EXT\"、not flag.value），"
            f"或直接写一个值按真假判断；当前条件：{expression!r}"
        )

    def _condition_ref(self, name: str, attribute: str | None, case_column: int) -> str:
        if attribute is None and name in self._CONDITION_SPECIAL_NAMES:
            return self.expression(name, case_column=case_column, line=name)
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
            f"模板 {self.template_name!r} 的 {{{{ if }}}} 条件里不支持 .{attribute}"
            f"（变量 {name!r}）：只支持 .value / .text / .prefix / .suffix"
        )

    def _require_attribute_access(self, node, expression: str, line: str) -> None:
        """条件里的变量必须带属性；裸变量（除 ``case_name`` / ``template_name``）直接报错。"""

        def walk(current, parent) -> None:
            if isinstance(current, nodes.Name) and not isinstance(parent, nodes.Getattr):
                if current.name not in self._CONDITION_SPECIAL_NAMES:
                    raise FormulaError(
                        f"模板 {self.template_name!r} 的 {{{{ if }}}} 条件里裸写了变量 "
                        f"{current.name!r}：请指明要比较哪一部分 —— 数值比较写 "
                        f"{current.name}.value（快照模式里裸变量是 VarValue 对象，"
                        f"拿它跟数字比会直接报错）；条件：{expression!r}"
                    )
            for child in current.iter_child_nodes():
                walk(child, current)

        walk(node, None)

    def _is_numeric_node(self, node) -> bool:
        if isinstance(node, nodes.Getattr) and isinstance(node.node, nodes.Name):
            return self._is_numeric(node.node.name)
        return False

    def _is_numeric(self, name: str) -> bool:
        definition = self._globals.get(name) or self._locals.get(name)
        if definition is None:
            return False
        if definition.type in ("int", "float"):
            return True
        default = definition.default
        return isinstance(default, (int, float)) and not isinstance(default, bool)

    # -- 递归编译 ---------------------------------------------------------- #
    def _compile_tokens(self, tokens: list[_Token], *, case_column: int, line: str) -> str:
        index = 0

        def join(parts: list[str]) -> str:
            return "&".join(parts) if parts else '""'

        def parse_block(stop: tuple[str, ...]) -> str:
            nonlocal index
            parts: list[str] = []
            while index < len(tokens):
                token = tokens[index]
                if token.kind == "tag":
                    words = token.value.split()
                    keyword = words[0] if words else ""
                    if keyword in stop:
                        return join(parts)
                    if keyword == "if":
                        index += 1
                        parts.append(parse_if())
                        continue
                    if keyword in ("else", "elif", "endif"):
                        raise FormulaError(
                            f"模板 {self.template_name!r} 里的 {{% {keyword} %}} 没有对应的 "
                            f"{{% if %}}；行内容：{line.strip()!r}"
                        )
                    raise FormulaError(
                        f"模板 {self.template_name!r} 的公式模式不支持 {{% {keyword} %}}："
                        "行内只支持 {% if %} / {% else %} / {% endif %}（因为一行对应一个单元格，"
                        "循环与跨行分支没法用单元格引用表达）；请把该模板改回 engine: snapshot。"
                        f"行内容：{line.strip()!r}"
                    )
                index += 1
                if token.kind == "text":
                    if token.value:
                        parts.append(_quote_text(token.value))
                elif token.kind == "expr":
                    try:
                        parts.append(
                            self.expression(token.value, case_column=case_column, line=line)
                        )
                    except FormulaError as exc:
                        # 统一带上出错的行，方便在几十行的模板里定位
                        raise FormulaError(f"{exc}；行内容：{line.strip()!r}") from exc
                # kind == "comment"：整段丢掉
            if stop:
                raise FormulaError(
                    f"模板 {self.template_name!r} 的 {{{{ if }}}} 没有闭合（这一行少了 "
                    f"{{% endif %}}）。公式模式**每行对应一个单元格**，所以 {{% if %}} 必须"
                    f"写在同一行内；行内容：{line.strip()!r}"
                )
            return join(parts)

        def parse_if() -> str:
            nonlocal index
            condition_source = tokens[index - 1].value[len("if"):].strip()
            then_expr = parse_block(("else", "elif", "endif"))
            else_expr = '""'
            if tokens[index].value.split()[0] == "else":
                index += 1
                else_expr = parse_block(("endif",))
            index += 1  # 跳过 endif
            condition = self.condition(condition_source, case_column=case_column, line=line)
            return f"IF({condition},{then_expr},{else_expr})"

        return parse_block(())


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
