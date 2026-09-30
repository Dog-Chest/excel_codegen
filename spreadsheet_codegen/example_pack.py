"""随包发布的内置示例（``spreadsheet_codegen/examples/``）：定位、列出、复制。

示例是**包数据** —— wheel 与 sdist 里都有一份，所以 ``pip install spreadsheet-codegen`` /
``uv tool install spreadsheet-codegen`` 装完、没有克隆仓库的人也能拿到：

    spreadsheet-codegen examples                  # 看有哪些示例、各自演示什么
    spreadsheet-codegen examples --copy ./demo    # 拷出来直接用

它同时是仓库里的回归夹具（``abs_fpi/`` 的现场脚本与 ``tests/`` 都直接读这几个 YAML），
所以仓库根目录**不再**单独放一份 ``examples/``，免得两份漂移。
"""

from __future__ import annotations

import shutil
from dataclasses import dataclass
from pathlib import Path

from .utils import CodeGenError

__all__ = ["EXAMPLES", "Example", "copy_examples", "examples_root"]


@dataclass(frozen=True)
class Example:
    """一个内置示例：目录名 + 给人看的说明 + 两个常用路径（都相对示例根目录）。"""

    name: str
    title: str
    summary: str
    entry: str
    workbook: str


#: 示例清单。新增一个示例 = 在 ``examples/<name>/`` 放好文件 + 在这里登记；
#: ``tests/test_examples_command.py`` 会检查两边一一对应，漏登记会红。
EXAMPLES: tuple[Example, ...] = (
    Example(
        name="basic",
        title="入门：公式模式与快照模式",
        summary=(
            "STM32 风格的 UART 配置。example_formula.yaml 是默认的公式模式"
            "（改参数在 Excel 里自己重算，不用跑命令）；example.yaml 显式写 engine: snapshot，"
            "演示过滤器 / 循环 / 导出文件。"
        ),
        entry="basic/example_formula.yaml",
        workbook="basic/template_formula.xlsx",
    ),
    Example(
        name="nastran",
        title="NASTRAN 工况控制：一行一个工况",
        summary=(
            "local_direction: vertical —— Local 表一行一个工况，下拉即增行；"
            "语句留空就不输出，最后拼成 .inc / .deck 工况控制语句（指南 §19 / §20）。"
        ),
        entry="nastran/nastran_case_control.yaml",
        workbook="nastran/nastran_case_control.xlsx",
    ),
    Example(
        name="abs_fpi",
        title="现场用例：ABS FPI 内外压 → GeniE",
        summary=(
            "两个规则集（外压 5A-3-2/5.5、内压 5A-3-2/5.7）可以各自单独生成一本工作簿，"
            "也可以合成一本项目工作簿共用一张 Global 表；内压用成员表 Tank Data 把舱参数只写一遍。"
        ),
        entry="abs_fpi/abs_fpi_internal.yaml",
        workbook="abs_fpi/abs_fpi_internal.xlsx",
    ),
)


def examples_root() -> Path:
    """内置示例的根目录（可编辑安装与 wheel 安装都在同一个相对位置）。"""
    return Path(__file__).resolve().parent / "examples"


def find(name: str) -> Example:
    for example in EXAMPLES:
        if example.name == name:
            return example
    available = "、".join(item.name for item in EXAMPLES)
    raise CodeGenError(f"没有这个示例：{name}（可选：{available}）")


def copy_examples(dest: Path, *, only: str | None = None, force: bool = False) -> list[Path]:
    """把示例拷进 ``dest/<示例名>/``，返回拷出来的顶层目录列表。

    :param only: 只拷这一个（默认全部）。
    :param force: 目标目录已存在且非空时是否覆盖。
    """
    selected = (find(only),) if only else EXAMPLES
    root = examples_root()
    if not root.is_dir():  # pragma: no cover - 只有被裁剪过的安装才会走到
        raise CodeGenError(f"找不到内置示例目录：{root}（安装包可能不完整，重装一次试试）")

    dest.mkdir(parents=True, exist_ok=True)
    written: list[Path] = []
    for example in selected:
        src = root / example.name
        if not src.is_dir():  # pragma: no cover - 由 tests/test_examples_command.py 守住
            raise CodeGenError(f"内置示例 {example.name} 缺文件：{src}")
        target = dest / example.name
        if target.exists() and any(target.iterdir()) and not force:
            raise CodeGenError(f"{target} 已存在且非空：加 --force 覆盖，或换一个目录")
        shutil.copytree(src, target, dirs_exist_ok=True)
        written.append(target)
    return written
