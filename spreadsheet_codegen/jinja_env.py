"""Jinja2 环境与自定义过滤器（模板渲染与派生参数求值共用一套语义）。

单独成模块的原因：``renderer`` 依赖 ``excel_io``，而 ``excel_io`` 又要能求值派生参数
（``derived:``）；把环境构造放在最底层，两边都 import 它，避免循环依赖。

**模板渲染与派生参数必须用同一个环境** —— 否则 ``{{ x }}`` 与 ``derived: x`` 的
数值形态（整数浮点、bool、None）会不一致。
"""

from __future__ import annotations

from collections.abc import Sequence
from pathlib import Path
from typing import Any

from jinja2 import Environment, FileSystemLoader, StrictUndefined, Undefined

from .utils import to_text

__all__ = ["build_environment", "pvs", "wrap"]


# --------------------------------------------------------------------------- #
# 过滤器
# --------------------------------------------------------------------------- #
def pvs(value: Any, prefix: Any = "", suffix: Any = "") -> str:
    """Prefix + Value + Suffix 过滤器：``{{ "A" | pvs("GPIO", "_PORT") }}`` -> ``GPIOA_PORT``。"""
    return f"{to_text(prefix)}{to_text(value)}{to_text(suffix)}"


def wrap(value: Any, prefix: Any = "", suffix: Any = "") -> str:
    """``pvs`` 的别名：``{{ "A" | wrap("GPIO", "_PORT") }}`` -> ``GPIOA_PORT``。"""
    return pvs(value, prefix, suffix)


# --------------------------------------------------------------------------- #
# 环境
# --------------------------------------------------------------------------- #
def finalize_value(value: Any) -> Any:
    """Jinja2 输出收尾：让 ``{{ x }}`` 与 ``{{ x.value }}`` 的文本形态一致。

    只在 ``bool`` / ``float`` / ``None`` 上做规范化，其余类型原样交给 Jinja2：

    * ``{{ x.value }}``（float ``340.0``）-> ``340``，与 ``{{ x }}`` 相同
    * ``{{ flag.value }}``（bool）-> ``true`` / ``false``，与 :func:`utils.to_text` 相同
    * ``{{ none_value }}`` -> ``""``（而不是 ``None``）
    """
    if value is None or isinstance(value, (bool, float)):
        return to_text(value)
    return value


def build_environment(
    *,
    strict: bool = True,
    trim_blocks: bool = False,
    lstrip_blocks: bool = False,
    keep_trailing_newline: bool = True,
    search_path: Sequence[str | Path] | None = None,
    **options: Any,
) -> Environment:
    """构建带 ``pvs`` / ``wrap`` 过滤器的 Jinja2 环境。

    默认使用 ``StrictUndefined``：模板引用了不存在的变量会立刻报错，而不是静默渲染成空串。
    ``finalize`` 负责把数值统一成文本形态（见 :func:`finalize_value`）。

    :param search_path: 给了就装一个 ``FileSystemLoader``，模板里的 ``{% include "片段.j2" %}``
        会按这些目录去找（相对**声明模板的那个文件**解析，见指南 §17）。
    """
    finalize = options.pop("finalize", finalize_value)
    loader = FileSystemLoader([str(item) for item in search_path]) if search_path else None
    environment = Environment(
        loader=loader,
        undefined=StrictUndefined if strict else Undefined,
        trim_blocks=trim_blocks,
        lstrip_blocks=lstrip_blocks,
        keep_trailing_newline=keep_trailing_newline,
        autoescape=False,
        finalize=finalize,
        **options,
    )
    environment.filters["pvs"] = pvs
    environment.filters["wrap"] = wrap
    # 常用数学函数：Jinja 只把 min/max 做成过滤器，表达式里 `max(a, b)` 会报 undefined。
    # 派生参数（derived:）与模板都用同一套全局函数，语义是 Python 的。
    environment.globals.setdefault("min", min)
    environment.globals.setdefault("max", max)
    environment.globals.setdefault("abs", abs)
    environment.globals.setdefault("int", int)
    environment.globals.setdefault("float", float)
    environment.globals.setdefault("round", round)
    # len 同时给"函数调用"与"过滤器"两种写法：`len(x)` / `x | len` / `x | length`。
    # 三种在 Python 侧都是 len()，Excel 侧都是 LEN() —— 补齐它是因为它最常用
    # （此前 `derived: "len(secret.value)"` 会报"引用了未定义的变量 'len'"）。
    environment.globals.setdefault("len", len)
    environment.filters.setdefault("len", len)
    environment.filters.setdefault("length", len)
    return environment
