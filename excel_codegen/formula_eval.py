"""公式求值器：把 ``engine: excel`` 写出的公式在 Python 里算一遍。

为什么需要它
------------
公式模式的工作簿里，``Output`` 表每一格是 Excel 公式，而 ``openpyxl`` **不会算公式**。
如果 ``check`` 只比"公式文本是否与当前 YAML 一致"，就会留下一个空档：

* 工具侧：只比公式；
* 用户侧：打开 Excel 看到的那段代码，**谁也没验过**它等于 ``--outdir`` 导出的那份。

本模块实现公式子集（``&`` 拼接 / ``IF`` / ``ISBLANK`` / ``TEXT`` / ``INDEX``+``MATCH`` /
整列引用 / 单元格引用 / 字符串与数字字面量）的一个小求值器，把输出表里的公式逐格算成文本，
再与 :func:`excel_codegen.renderer.render_all` 的结果逐行比对。两边一致 ⇒
"Excel 里看到的"与"命令行导出的"是同一段代码。

它**解析的是工作簿里真实的公式文本**（不是重新生成一遍），所以能抓到"列标指错"
"该用 ISBLANK 却用了 =''" 这类**生成端**的 bug —— 这是它存在的意义。

局限
----
* 参数单元格如果本身是公式（用户手写 `=L/10`），离线读不到值 → 抛
  :class:`FormulaEvalError`，调用方应当降级为"只比公式文本"并给出提示。
* ``TEXT()`` 只实现本工具生成的两种格式串（``0`` / ``0.############``）。
* 不实现的函数/语法一律报错，不猜。
"""

from __future__ import annotations

import re
from typing import Any, Mapping, Protocol, Sequence

from .utils import RenderError, column_index_to_letter, to_text

__all__ = [
    "Cell",
    "ColRef",
    "Evaluator",
    "FormulaEvalError",
    "WorkbookReader",
    "evaluate_formula",
    "evaluate_template_values",
]


class FormulaEvalError(RenderError):
    """公式无法离线求值（语法超出子集，或引用了公式单元格）。"""


# --------------------------------------------------------------------------- #
# 词法
# --------------------------------------------------------------------------- #
_TOKEN_RE = re.compile(
    r"""
      (?P<string>"(?:[^"]|"")*")
    | (?P<sheet>'(?:[^']|'')*')
    | (?P<number>\d+(?:\.\d*)?)
    | (?P<ident>[A-Za-z_][A-Za-z0-9_.]*)
    | (?P<punct>[&(),!=:$+\-*/^<>])
    | (?P<space>\s+)
    """,
    re.VERBOSE,
)


def tokenize(text: str) -> list[tuple[str, str]]:
    out: list[tuple[str, str]] = []
    index = 0
    while index < len(text):
        match = _TOKEN_RE.match(text, index)
        if not match:
            raise FormulaEvalError(f"不能识别的公式片段: {text[index:index + 30]!r}")
        index = match.end()
        if match.lastgroup == "space":
            continue
        out.append((match.lastgroup or "", match.group()))
    return out


# --------------------------------------------------------------------------- #
# 值模型
# --------------------------------------------------------------------------- #
class Cell:
    """单元格引用（或 INDEX 的结果）：值 + 是否真空白（``ISBLANK`` 用）。"""

    __slots__ = ("value", "blank", "where")

    def __init__(self, value: Any, blank: bool, where: str = "") -> None:
        self.value = value
        self.blank = blank
        self.where = where

    def __repr__(self) -> str:  # pragma: no cover - 调试友好
        return f"Cell({self.value!r}, blank={self.blank}, {self.where})"


class ColRef:
    """整列引用（只作为 ``INDEX`` / ``MATCH`` 的参数出现）。"""

    __slots__ = ("sheet", "column")

    def __init__(self, sheet: str, column: str) -> None:
        self.sheet = sheet
        self.column = column

    def __repr__(self) -> str:  # pragma: no cover - 调试友好
        return f"ColRef({self.sheet}!{self.column})"


def as_text(value: Any) -> str:
    """Excel 把值拼进字符串时的形态。

    数值走 General 规则 = **最多 15 位有效数字**，这正是 :func:`utils.to_text` 的规则，
    所以这里直接复用它（``True``/``False`` 仍然按 Excel 的 ``TRUE``/``FALSE``）。
    """
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    return to_text(value)


def format_number(value: Any, fmt: str) -> str:
    """只实现本工具生成的两种格式串：``0`` 与 ``0.############``。"""
    if value is None or value == "":
        return ""
    try:
        number = float(value)
    except (TypeError, ValueError):
        return str(value)
    if fmt == "0":
        return str(int(round(number)))
    if fmt.startswith("0."):
        decimals = len(fmt) - 2
        text = f"{number:.{decimals}f}".rstrip("0").rstrip(".")
        return text if text not in ("", "-") else "0"
    raise FormulaEvalError(f"不支持的数字格式 {fmt!r}")


# --------------------------------------------------------------------------- #
# 数据源
# --------------------------------------------------------------------------- #
class SheetReader(Protocol):
    """求值器眼里的"工作簿"。"""

    def cell_value(self, sheet: str, column: int, row: int) -> tuple[Any, bool]:
        """返回 ``(值, 是否真空白)``。"""

    def name_row(self, sheet: str, variable: str) -> int:
        """在 A 列按**变量名**查行号（等同于 MATCH）。"""


class WorkbookReader:
    """从 ``openpyxl`` 工作簿读值。

    参数格里如果是**公式**（例如派生参数 ``derived`` 写成的公式），会递归求值 ——
    所以"公式引用公式"也能算出来；出现循环引用会明确报错。
    """

    def __init__(self, workbook, *, evaluate=None) -> None:
        self.workbook = workbook
        #: 由 :class:`Evaluator` 装上：公式文本 -> 文本
        self.evaluate = evaluate
        self._name_rows: dict[str, dict[str, int]] = {}
        self._cells: dict[tuple[str, int, int], tuple[Any, bool]] = {}
        self._active: set[tuple[str, int, int]] = set()

    def bind_evaluator(self, evaluate) -> None:
        """把求值器挂上来（嵌套公式要用）。"""
        self.evaluate = evaluate

    def cell_value(self, sheet: str, column: int, row: int) -> tuple[Any, bool]:
        key = (sheet, column, row)
        if key in self._cells:
            return self._cells[key]
        try:
            worksheet = self.workbook[sheet]
        except KeyError as exc:
            raise FormulaEvalError(f"公式引用了不存在的工作表 {sheet!r}") from exc

        value = worksheet.cell(row=row, column=column).value
        if isinstance(value, str) and value.startswith("="):
            if self.evaluate is None:
                raise FormulaEvalError(
                    f"{sheet}!{column_index_to_letter(column)}{row} 是公式，但没有可用的求值器"
                )
            if key in self._active:
                raise FormulaEvalError(
                    f"{sheet}!{column_index_to_letter(column)}{row} 出现循环引用"
                )
            self._active.add(key)
            try:
                result = (self.evaluate(value), False)
            finally:
                self._active.discard(key)
        else:
            result = (value, value is None or value == "")
        self._cells[key] = result
        return result

    def name_row(self, sheet: str, variable: str) -> int:
        cache = self._name_rows.setdefault(sheet, {})
        if not cache:
            try:
                worksheet = self.workbook[sheet]
            except KeyError as exc:
                raise FormulaEvalError(f"公式引用了不存在的工作表 {sheet!r}") from exc
            for row in range(1, worksheet.max_row + 1):
                text = worksheet.cell(row=row, column=1).value
                if text is not None and str(text).strip():
                    cache.setdefault(str(text).strip(), row)
        if variable not in cache:
            raise FormulaEvalError(f"{sheet} 表里没有变量 {variable!r}")
        return cache[variable]


# --------------------------------------------------------------------------- #
# 求值
# --------------------------------------------------------------------------- #
class Evaluator:
    """递归下降求值器（只覆盖本工具生成的公式子集）。"""

    def __init__(self, reader: SheetReader, sheet_names: Mapping[str, str]) -> None:
        self.reader = reader
        #: 小写表名 -> 真实表名（Excel 的表名匹配不区分大小写）
        self.sheets = {name.strip().lower(): name for name in sheet_names.values()}
        bind = getattr(reader, "bind_evaluator", None)
        if callable(bind):
            bind(self.evaluate)  # 让"公式格里的公式"也能递归求值

    # -- 入口 -------------------------------------------------------------- #
    def evaluate(self, formula: Any) -> str:
        if not isinstance(formula, str) or not formula.startswith("="):
            return as_text(formula)
        # 可重入：嵌套求值（参数格本身是公式）会递归调用本方法
        saved = (getattr(self, "tokens", None), getattr(self, "index", 0))
        self.tokens = tokenize(formula[1:])
        self.index = 0
        try:
            value = self.comparison()
            if self.index != len(self.tokens):
                raise FormulaEvalError(f"公式没有解析完: {self.tokens[self.index:]}")
            return as_text(self._scalar(value))
        finally:
            self.tokens, self.index = saved

    # -- 语法 -------------------------------------------------------------- #
    # Excel 的优先级：比较 < 拼接(&) < 加减 < 乘除 < 幂 < 一元 < 项
    def comparison(self):
        left = self.concat()
        while True:
            if self._peek("="):
                self._next()
                right = self.concat()
                left = as_text(self._scalar(left)) == as_text(self._scalar(right))
            elif self._peek("!"):
                self._next()
                self._expect("=")
                right = self.concat()
                left = as_text(self._scalar(left)) != as_text(self._scalar(right))
            elif self._peek("<") or self._peek(">"):
                operator = self._next()[1]
                if self._peek("="):
                    self._next()
                    operator += "="
                elif operator == "<" and self._peek(">"):
                    self._next()
                    operator = "<>"
                right = self.concat()
                left = self._compare(left, right, operator)
            else:
                break
        return left

    def _compare(self, left, right, operator: str) -> bool:
        """``<`` / ``>`` / ``<=`` / ``>=`` / ``<>``。

        Excel 对**数值**按数值比、对**文本**按字典序比 —— 这里照做（能读成数字就当数字）。
        行内 ``{% if %}`` 编译出来的条件会用到这些运算符。
        """
        a, b = self._scalar(left), self._scalar(right)
        try:
            x, y = self._number(a), self._number(b)
        except FormulaEvalError:
            x, y = as_text(a), as_text(b)
        if operator == "<":
            return x < y
        if operator == ">":
            return x > y
        if operator == "<=":
            return x <= y
        if operator == ">=":
            return x >= y
        if operator == "<>":
            return x != y
        raise FormulaEvalError(f"不支持的比较运算符 {operator!r}")

    def concat(self):
        left = self.additive()
        while self._peek("&"):
            self._next()
            right = self.additive()
            left = as_text(self._scalar(left)) + as_text(self._scalar(right))
        return left

    def additive(self):
        left = self.multiplicative()
        while self._peek("+") or self._peek("-"):
            operator = self._next()[1]
            right = self.multiplicative()
            left = self._number(left) + self._number(right) if operator == "+" else self._number(left) - self._number(right)
        return left

    def multiplicative(self):
        left = self.power()
        while self._peek("*") or self._peek("/"):
            operator = self._next()[1]
            right = self.power()
            if operator == "*":
                left = self._number(left) * self._number(right)
            else:
                divisor = self._number(right)
                if divisor == 0:
                    raise FormulaEvalError("公式里出现除零")
                left = self._number(left) / divisor
        return left

    def power(self):
        left = self.unary()
        while self._peek("^"):
            self._next()
            right = self.unary()
            left = self._number(left) ** self._number(right)
        return left

    def unary(self):
        if self._peek("-"):
            self._next()
            return -self._number(self.unary())
        if self._peek("+"):
            self._next()
            return self._number(self.unary())
        if self._peek("("):
            self._next()
            value = self.comparison()
            self._expect(")")
            return value
        return self.term()

    def _number(self, value):
        """取数值（Cell 取值；文本按 Excel 的隐式转换规则）。"""
        raw = self._scalar(value)
        if raw is None or raw == "":
            return 0  # Excel 把空单元格当 0 参与算术
        if isinstance(raw, bool):
            return 1 if raw else 0
        try:
            return float(raw)
        except (TypeError, ValueError):
            raise FormulaEvalError(f"{raw!r} 不是数值，无法参与算术") from None

    def term(self):
        kind, text = self._next()
        if kind == "string":
            return text[1:-1].replace('""', '"')
        if kind == "number":
            return float(text) if "." in text else int(text)
        if kind == "sheet" or (kind == "ident" and self._peek("!")):
            return self.reference(kind, text)
        if kind == "ident":
            upper = text.upper()
            if not self._peek("("):
                if upper in {"TRUE", "FALSE"}:
                    return upper == "TRUE"
                raise FormulaEvalError(f"裸标识符 {text!r} 无法求值")
            self._next()  # 吃掉 "("
            if upper == "IF":
                condition = self.comparison()
                self._expect(",")
                yes = self.comparison()
                self._expect(",")
                no = self.comparison()
                self._expect(")")
                return yes if self.truthy(condition) else no
            if upper in {"AND", "OR"}:
                values = [self.comparison()]
                while self._peek(","):
                    self._next()
                    values.append(self.comparison())
                self._expect(")")
                flags = [self.truthy(item) for item in values]
                return all(flags) if upper == "AND" else any(flags)
            if upper == "NOT":
                value = self.comparison()
                self._expect(")")
                return not self.truthy(value)
            if upper == "ISBLANK":
                argument = self.comparison()
                self._expect(")")
                return bool(argument.blank) if isinstance(argument, Cell) else False
            if upper == "TEXT":
                value = self.comparison()
                self._expect(",")
                fmt = self.comparison()
                self._expect(")")
                return format_number(self._scalar(value), self._scalar(fmt))
            if upper == "INDEX":
                ref = self.comparison()
                self._expect(",")
                row = self.comparison()
                self._expect(")")
                if not isinstance(ref, ColRef):
                    raise FormulaEvalError("INDEX 的第一个参数不是整列引用")
                value, blank = self.reader.cell_value(ref.sheet, _column_number(ref.column), int(self._scalar(row)))
                return Cell(value, blank, f"{ref.sheet}!{ref.column}{int(self._scalar(row))}")
            if upper == "MATCH":
                needle = self.comparison()
                self._expect(",")
                ref = self.comparison()
                self._expect(",")
                self.comparison()  # 0（精确匹配）
                self._expect(")")
                if not isinstance(ref, ColRef):
                    raise FormulaEvalError("MATCH 的第二个参数不是整列引用")
                return self.reader.name_row(ref.sheet, str(self._scalar(needle)))
            if upper in {"MIN", "MAX", "ABS", "TRUNC", "MOD"}:
                arguments = [self.comparison()]
                while self._peek(","):
                    self._next()
                    arguments.append(self.comparison())
                self._expect(")")
                numbers = [self._number(argument) for argument in arguments]
                if upper == "MIN":
                    return min(numbers)
                if upper == "MAX":
                    return max(numbers)
                if upper == "ABS":
                    return abs(numbers[0])
                if upper == "TRUNC":
                    return int(numbers[0])
                divisor = numbers[1]
                if divisor == 0:
                    raise FormulaEvalError("MOD 的除数为 0")
                return numbers[0] - divisor * int(numbers[0] / divisor)
            raise FormulaEvalError(f"不支持的函数 {upper}()")
        raise FormulaEvalError(f"看不懂的记号 {kind}={text!r}")

    def reference(self, kind: str, text: str):
        sheet: str | None = None
        if kind == "sheet":
            sheet = text[1:-1].replace("''", "'")
            self._expect("!")
        if self._peek("$"):
            self._next()
        token_kind, column = self._next()
        if token_kind != "ident":
            raise FormulaEvalError(f"引用里缺列标: {column!r}")
        if self._peek(":"):
            self._next()
            if self._peek("$"):
                self._next()
            _, second = self._next()
            if second.upper() != column.upper():
                raise FormulaEvalError(f"只支持整列引用，不支持列区间 {column}:{second}")
            if sheet is None:
                raise FormulaEvalError("公式里出现了没有表名的整列引用")
            return ColRef(sheet, column.upper())
        if self._peek("$"):
            self._next()
        row_kind, row_text = self._next()
        if row_kind != "number":
            raise FormulaEvalError(f"引用里缺行号: {row_text!r}")
        if sheet is None:
            raise FormulaEvalError("公式里出现了没有表名的单元格引用")
        row = int(float(row_text))
        value, blank = self.reader.cell_value(sheet, _column_number(column), row)
        return Cell(value, blank, f"{sheet}!{column}{row}")

    # -- 小工具 ------------------------------------------------------------ #
    def _next(self) -> tuple[str, str]:
        if self.index >= len(self.tokens):
            raise FormulaEvalError("公式意外结束")
        token = self.tokens[self.index]
        self.index += 1
        return token

    def _peek(self, punct: str) -> bool:
        return self.index < len(self.tokens) and self.tokens[self.index] == ("punct", punct)

    def _expect(self, punct: str) -> None:
        if not self._peek(punct):
            got = self.tokens[self.index] if self.index < len(self.tokens) else "结束"
            raise FormulaEvalError(f"期望 {punct!r}，实际 {got}")
        self.index += 1

    @staticmethod
    def _scalar(value: Any) -> Any:
        return value.value if isinstance(value, Cell) else value

    @staticmethod
    def truthy(value: Any) -> bool:
        if isinstance(value, Cell):
            return value.blank is False and value.value not in (0, "", None)
        if isinstance(value, str):
            return value.upper() not in ("", "FALSE", "0")
        return bool(value)


def _column_number(letter: str) -> int:
    number = 0
    for char in letter.upper():
        if not char.isalpha():
            raise FormulaEvalError(f"非法列标 {letter!r}")
        number = number * 26 + (ord(char) - ord("A") + 1)
    return number


# --------------------------------------------------------------------------- #
# 面向工作簿的便捷入口
# --------------------------------------------------------------------------- #
def sheet_names_of(config) -> list[str]:
    names = [
        config.excel.sheets.global_,
        config.excel.sheets.local,
        *config.excel.sheets.outputs,
    ]
    for optional in (config.excel.template_sheet, config.excel.howto_sheet):
        if optional:
            names.append(optional)
    return names


def evaluate_formula(formula: Any, *, reader: SheetReader, config) -> str:
    """求值单格公式（主要给测试用）。"""
    evaluator = Evaluator(reader, {name: name for name in sheet_names_of(config)})
    return evaluator.evaluate(formula)


def evaluate_template_values(
    workbook,
    config,
    template,
    case_names: Sequence[str],
) -> dict[str, list[str]]:
    """把某个模板在输出表里的公式逐格算出来：``{case 名: [每行文本]}``。

    读的是**工作簿里真实的公式文本**，所以能发现"公式本身写错了"（列标、引用、函数）。
    """
    from .utils import parse_cell

    evaluator = Evaluator(WorkbookReader(workbook), {name: name for name in sheet_names_of(config)})
    worksheet = workbook[template.output_sheet]
    column, row = parse_cell(template.start_cell)
    blocks: dict[str, list[str]] = {}

    for offset, case_name in enumerate(case_names):
        lines: list[str] = []
        if template.direction == "horizontal":
            current = row
            while True:
                value = worksheet.cell(row=current, column=column + offset).value
                if value is None:
                    break
                lines.append(evaluator.evaluate(value))
                current += 1
        else:
            current = column
            target_row = row + offset
            while True:
                value = worksheet.cell(row=target_row, column=current).value
                if value is None:
                    break
                lines.append(evaluator.evaluate(value))
                current += 1
        blocks[case_name] = lines
    return blocks
