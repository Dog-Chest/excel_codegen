"""回归测试：`abs_fpi` 现场项目用**成员表**表达舱数据（指南 §18 / §19）。

`abs_fpi/` 是"拿这个工具干真活"的项目，0.8.0 起把 13 个舱参数从 `local` 挪进了
`variables.group`（工作表 `Tank Data`，一行一个舱），`local.tank_ref` 只是指针。
本文件把这次迁移**钉在 CI 里**（项目侧的逐条实测在 `abs_fpi/probes/probe_group_table.py`）：

1. 配置层面：舱参数都在 group 里、Local 表里没有它们、`tank_ref` 是带 choices 的指针；
2. 数据层面：`Tank Data` 表里三个舱各自只有一份数据，两个 WBT6 工况共用它；
3. 产物层面：`examples/abs_fpi/generated/` 的导出文件与当前 YAML + 工作簿一致
   （防止"改了 YAML 忘了重新生成"）。

0.9.0 起这些示例资产随包发布（`spreadsheet_codegen/examples/abs_fpi/`，见 `examples/README.md`），
所以 sdist 里也有；下面的"文件不在就跳过"只是给被裁剪过的安装留条活路。
"""

from __future__ import annotations

from pathlib import Path

import pytest

from spreadsheet_codegen import read_group_members
from spreadsheet_codegen.example_pack import examples_root
from spreadsheet_codegen.excel_io import load_workbook_file
from spreadsheet_codegen.models import load_config
from spreadsheet_codegen.renderer import render_all

#: 示例资产随包发布（0.9.0 起从 abs_fpi/ 挪到 spreadsheet_codegen/examples/abs_fpi/）；
#: 仓库根的 abs_fpi/ 只留现场脚本与实测报告（FINDINGS.md / probes/）
ABS_FPI = examples_root() / "abs_fpi"
INTERNAL_YAML = ABS_FPI / "abs_fpi_internal.yaml"
PROJECT_YAML = ABS_FPI / "abs_fpi.yaml"
EXTERNAL_YAML = ABS_FPI / "abs_fpi_external.yaml"

#: 舱参数（都在成员表里，不在 local 里）
TANK_VARIABLES = [
    "rho_tank",
    "l_tank",
    "b_tank",
    "h_tank",
    "eta_deck",
    "eta_overflow",
    "C_dp",
    "C_ru",
    "p_vp",
    "GM_full_in",
    "k_r_in",
    "tank_is_ballast",
    "member_11_17",
]


def _require(path: Path) -> Path:
    if not path.exists():  # pragma: no cover - 示例随包发布，只有被裁剪的安装才会缺
        pytest.skip(f"{path.name} 不在（示例资产缺失）")
    return path


@pytest.fixture(scope="module")
def internal():
    return load_config(_require(INTERNAL_YAML))


# --------------------------------------------------------------------------- #
# 1. 配置：舱数据在 group 里，local 里只剩指针
# --------------------------------------------------------------------------- #
def test_internal_declares_the_member_table(internal) -> None:
    group = internal.group
    assert group is not None, "舱数据应该放在 variables.group 里"
    assert group.sheet == "Tank Data"
    assert group.key == "tank_ref"
    assert group.members == ["WBT6", "WBT7", "COT1"]
    assert [item.name for item in internal.group_variables] == TANK_VARIABLES


def test_tank_variables_are_no_longer_local(internal) -> None:
    """迁移的核心断言：同一个舱不该在 Local 表里再写一遍。"""
    local_names = set(internal.local_names)
    assert not (set(TANK_VARIABLES) & local_names)
    # 指针本身仍然是 local（要"一个工况一份"才能指向不同舱）
    assert "tank_ref" in local_names
    key = next(item for item in internal.local_variables if item.name == "tank_ref")
    assert key.choices == ["WBT6", "WBT7", "COT1"]


def test_external_ruleset_has_no_member_table() -> None:
    """外压不需要舱，不该凭空多一张成员表。"""
    assert load_config(_require(EXTERNAL_YAML)).group is None


def test_composed_project_carries_the_member_table() -> None:
    """compose.py 要把成员表声明合并进项目工作簿 YAML。"""
    project = load_config(_require(PROJECT_YAML))
    assert project.group is not None
    assert project.group.sheet == "Tank Data"
    assert project.group.key == "tank_ref"
    assert project.group.members == ["WBT6", "WBT7", "COT1"]
    assert [item.name for item in project.group_variables] == TANK_VARIABLES
    # 外压工况也在同一本工作簿里，它们用默认的 tank_ref（WBT6）—— 不影响外压模板
    assert not (set(TANK_VARIABLES) & set(project.local_names))


# --------------------------------------------------------------------------- #
# 2. 数据：一个舱一份，多个工况共用
# --------------------------------------------------------------------------- #
def test_member_sheet_holds_one_row_per_tank(internal) -> None:
    book = _require(ABS_FPI / "abs_fpi_internal.xlsx")
    workbook = load_workbook_file(book)
    try:
        members = read_group_members(workbook, internal)
    finally:
        workbook.close()

    assert list(members) == ["WBT6", "WBT7", "COT1"]
    assert members["WBT6"]["l_tank"].text == "42"
    assert members["WBT7"]["l_tank"].text == "14.5"
    assert members["COT1"]["rho_tank"].text == "900"


def test_two_cases_share_one_tank(internal) -> None:
    """WBT6 被两个工况用到 —— 改一处就该两处都变（这里只验证它们同源）。"""
    book = _require(ABS_FPI / "abs_fpi_internal.xlsx")
    lines = {result.case_name: result.lines for result in render_all(internal, book).results["genie_int"]}

    def tank_line(case: str, prefix: str) -> str:
        return next(item for item in lines[case] if item.startswith(prefix))

    assert tank_line("WBT6-d8-mu90", "var l_tank") == "var l_tank = 42 m;"
    assert tank_line("WBT6-d15.059-mu0", "var l_tank") == "var l_tank = 42 m;"
    assert tank_line("WBT7-d20.559-mu90", "var l_tank") == "var l_tank = 14.5 m;"
    assert tank_line("COT1-d20.559-mu60", "var rho_tank") == "var rho_tank = 900 kg/m^3;"


# --------------------------------------------------------------------------- #
# 3. 产物：仓库里的导出文件与当前 YAML + 工作簿一致
# --------------------------------------------------------------------------- #
@pytest.mark.parametrize(
    ("case", "line"),
    [
        ("WBT6-d8-mu90", "var l_tank = 42 m;"),
        ("WBT7-d20.559-mu90", "var h_tank = 32 m;"),
        ("COT1-d20.559-mu60", "var rho_tank = 900 kg/m^3;"),
    ],
)
def test_generated_files_are_current(case: str, line: str) -> None:
    """`examples/abs_fpi/generated/` 是提交进仓库的产物 —— 改了 YAML 忘了重跑就该红。"""
    path = ABS_FPI / "generated" / f"{case}.js"
    if not path.exists():  # pragma: no cover - 示例随包发布，只有被裁剪的安装才会缺
        pytest.skip("examples/abs_fpi/generated/ 不在")
    assert line in path.read_text(encoding="utf-8")
