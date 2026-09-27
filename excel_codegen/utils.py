"""基础工具：错误类型、Excel 单元格地址运算、Prefix/Value/Suffix 组合值。

本模块刻意不依赖任何第三方库，方便单独单元测试与复用。
"""

from __future__ import annotations

import hashlib
import re
from collections.abc import Iterable
from dataclasses import dataclass
from typing import Any

__all__ = [
    "CodeGenError",
    "ConfigError",
    "ExcelError",
    "RenderError",
    "VarValue",
    "cell_ref",
    "column_index_to_letter",
    "column_letter_to_index",
    "fingerprint",
    "is_identifier",
    "parse_cell",
    "safe_filename",
    "slugify",
    "split_cell",
    "split_lines",
    "to_text",
]


# --------------------------------------------------------------------------- #
# 错误类型
# --------------------------------------------------------------------------- #
class CodeGenError(Exception):
    """excel_codegen 所有面向用户的错误的基类（CLI 会友好地打印它们）。"""


class ConfigError(CodeGenError):
    """YAML 配置缺失、无法解析或语义非法。"""


class ExcelError(CodeGenError):
    """Excel 文件缺失、无法读取，或工作表结构与配置不一致。"""


class RenderError(CodeGenError):
    """Jinja2 模板语法错误或渲染期错误（例如变量缺失）。"""


# --------------------------------------------------------------------------- #
# 单元格 / 列标运算
# --------------------------------------------------------------------------- #
_CELL_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?([0-9]+)$")
_IDENTIFIER_RE = re.compile(r"^[A-Za-z_][A-Za-z0-9_]*$")
_UNSAFE_FILENAME_RE = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


def column_letter_to_index(letter: str) -> int:
    """``"A" -> 1``、``"Z" -> 26``、``"AA" -> 27``。非法输入抛 ``ValueError``。"""
    letters = str(letter).strip().upper()
    if not letters or not letters.isalpha():
        raise ValueError(f"非法列标: {letter!r}")
    index = 0
    for char in letters:
        index = index * 26 + (ord(char) - ord("A") + 1)
    if index < 1:
        raise ValueError(f"非法列标: {letter!r}")
    return index


def column_index_to_letter(index: int) -> str:
    """``1 -> "A"``、``27 -> "AA"``。"""
    value = int(index)
    if value < 1:
        raise ValueError(f"列号必须 >= 1，收到 {index!r}")
    letters = ""
    while value:
        value, remainder = divmod(value - 1, 26)
        letters = chr(ord("A") + remainder) + letters
    return letters


def parse_cell(cell: str) -> tuple[int, int]:
    """把 ``"B2"`` 解析为 ``(列号, 行号)``（均从 1 开始）。"""
    match = _CELL_RE.match(str(cell).strip())
    if not match:
        raise ValueError(f"非法单元格引用: {cell!r}（示例: B2）")
    return column_letter_to_index(match.group(1)), int(match.group(2))


def split_cell(cell: str) -> tuple[str, int]:
    """把 ``"B2"`` 解析为 ``("B", 2)``。"""
    column, row = parse_cell(cell)
    return column_index_to_letter(column), row


def cell_ref(column: int, row: int) -> str:
    """``(2, 2) -> "B2"``。"""
    return f"{column_index_to_letter(column)}{row}"


# --------------------------------------------------------------------------- #
# 值处理
# --------------------------------------------------------------------------- #
#: 超过这个量级的浮点不再做"整数化"处理（避免 1e20 变成 21 位数字串）。
_INT_SAFE_LIMIT = 1e16


def to_text(value: Any) -> str:
    """把 Excel / YAML 值转换成生成代码中应出现的文本形式。

    * ``None`` -> ``""``
    * ``True`` / ``False`` -> ``"true"`` / ``"false"``
    * ``115200.0`` -> ``"115200"``（整数浮点去掉小数点，避免 Excel 数值失真）
    * ``155.85637499999999`` -> ``"155.856375"``（保留 15 位有效数字，与 Excel/WPS 的显示一致）
    * ``1e20`` -> ``"1e+20"``（量级过大时保留科学计数法）

    15 位有效数字这一条很关键：派生参数（``derived:``）是算出来的，二进制浮点的尾巴
    （``0.1 + 0.2`` 那种）如果原样输出，就会出现"Excel 里显示 155.856375、导出文件里写
    155.85637499999999"的假差异；Excel 的 General 格式本身就是 15 位有效数字，
    所以这里对齐它。Jinja2 侧还通过 ``finalize`` 对 ``{{ x.value }}`` 施加同一规则。

    > 残留差异：有效数字超过 15 位的量（例如 17 位整数部分）两边仍可能不同 ——
    > ``excel-codegen check`` 会在公式模式下把公式算一遍来发现它。
    """
    if value is None:
        return ""
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, float):
        if value.is_integer() and abs(value) < _INT_SAFE_LIMIT:
            return str(int(value))
        # %.15g：15 位有效数字（<= 0 时保留科学计数法），再取最短往返表示
        return repr(float(f"{value:.15g}"))
    if isinstance(value, str):
        return value
    return str(value)


def fingerprint(*parts: Any) -> str:
    """对若干段文本求一个短指纹（sha1 前 12 位），用于工作簿里的"参数/输出"记录。"""
    rows: list[str] = []
    for part in parts:
        if isinstance(part, str):
            rows.append(part)
        elif isinstance(part, Iterable):
            rows.extend(to_text(item) for item in part)
        else:
            rows.append(to_text(part))
    return hashlib.sha1("\n".join(rows).encode("utf-8")).hexdigest()[:12]


def split_lines(text: str) -> list[str]:
    """按行拆分渲染结果，且不保留结尾空行。"""
    return str(text).splitlines()


def is_identifier(name: str) -> bool:
    """判断是否为合法 Python/Jinja2 标识符（变量名必须满足）。"""
    return bool(_IDENTIFIER_RE.match(str(name)))


def slugify(name: str) -> str:
    """把任意文本转成适合做文件名/标识符的片段。"""
    cleaned = re.sub(r"[^0-9A-Za-z_.-]+", "_", str(name).strip()).strip("_.")
    return cleaned or "unnamed"


def safe_filename(name: str) -> str:
    """去掉 Windows / POSIX 文件名中的非法字符。"""
    cleaned = _UNSAFE_FILENAME_RE.sub("_", str(name).strip()).strip(" .")
    return cleaned or "output"


# --------------------------------------------------------------------------- #
# Prefix + Value + Suffix
# --------------------------------------------------------------------------- #
@dataclass(frozen=True)
class VarValue:
    """变量值及其前缀/后缀包装对象。

    渲染上下文中的每个变量都是 ``VarValue``，因此在 Jinja2 模板里::

        {{ port }}          ->  "GPIOA_PORT"   (str(VarValue)，即 prefix+value+suffix)
        {{ port.value }}    ->  "A"            (纯值)
        {{ port.prefix }}   ->  "GPIO"
        {{ port.suffix }}   ->  "_PORT"

    ``value`` 保留原始类型（int / float / str 均可），只有转成字符串时才做
    "整数浮点去零" 的规范化，因此 ``{{ baud }}`` 输出 ``115200`` 而不是 ``115200.0``。
    """

    value: Any = ""
    prefix: str = ""
    suffix: str = ""

    def __post_init__(self) -> None:
        object.__setattr__(self, "prefix", "" if self.prefix is None else str(self.prefix))
        object.__setattr__(self, "suffix", "" if self.suffix is None else str(self.suffix))

    @property
    def text(self) -> str:
        """纯值的文本形式（``None`` -> ``""``）。"""
        return to_text(self.value)

    @property
    def is_empty(self) -> bool:
        """值、前缀、后缀全为空时为 ``True``。"""
        return not self.text and not self.prefix and not self.suffix

    def __str__(self) -> str:
        return f"{self.prefix}{self.text}{self.suffix}"

    def __bool__(self) -> bool:
        return not self.is_empty

    def __format__(self, format_spec: str) -> str:
        return format(str(self), format_spec)

    def as_dict(self) -> dict[str, Any]:
        """便于调试与测试的字典形式。"""
        return {"value": self.value, "prefix": self.prefix, "suffix": self.suffix}
