"""取值约束（``min`` / ``max`` / ``choices`` / ``pattern``）的配置期与运行期校验。

覆盖三件事：
1. 配置期：YAML 里写错约束（min>max、给字符串加范围、default 越界、派生参数加约束……）要立刻报错；
2. 生成期：约束要写成 Excel 的数据有效性（下拉 / 数值范围），表达不了时报错；
3. 运行期：表里填的值越界要在 render / validate / check 三处都被拦住。
"""

from __future__ import annotations

from pathlib import Path

import pytest
from openpyxl import load_workbook

from spreadsheet_codegen import (
    ExcelError,
    create_template,
    load_config,
    render_all,
)
from spreadsheet_codegen.excel_io import check_value_constraints
from spreadsheet_codegen.models import CaseData, VariableDef
from spreadsheet_codegen.utils import ConfigError, VarValue

BASE = """\
version: 1

excel:
  output: "template.xlsx"
  template_sheet: null
  howto_sheet: null
  sheets:
    global: "Global Parameter"
    local: "Local Parameter"
    outputs: ["Output"]

variables:
  global:
__GLOBALS__
  local:
__LOCALS__

templates:
  # 这里引用了只写在 Excel 表里、没进 variables 的 kind（快照模式允许），
  # 所以显式声明 engine: snapshot（项目默认为公式模式 excel，它要求变量都定义过）
  - name: demo
    output_sheet: "Output"
    engine: "snapshot"
    code: |
      // {{ kind }} L={{ draft }}
"""


def _config(tmp_path: Path, globals_: str, locals_: str = "    []"):
    text = BASE.replace("__GLOBALS__", globals_).replace("__LOCALS__", locals_)
    path = tmp_path / "demo.yaml"
    path.write_text(text, encoding="utf-8")
    return load_config(path)


# --------------------------------------------------------------------------- #
# 1. 配置期：约束本身写错了要在 load_config 就报出来
# --------------------------------------------------------------------------- #
@pytest.mark.parametrize(
    ("snippet", "keyword"),
    [
        ("    - name: v\n      min: 5\n      max: 1\n", "大于"),
        ("    - name: v\n      type: string\n      min: 1\n", "不能加 min"),
        ("    - name: v\n      min: 1\n      default: 0\n", "default"),
        ('    - name: v\n      derived: "1 + 1"\n      min: 0\n', "派生参数"),
        ("    - name: v\n      choices: []\n", "不能是空列表"),
        ("    - name: v\n      choices: [A, A]\n", "重复"),
        ("    - name: v\n      choices: A\n", "必须是列表"),
        ('    - name: v\n      pattern: "["\n', "正则"),
        ("    - name: v\n      min: abc\n", "必须是数字"),
        ("    - name: v\n      min: true\n", "必须是数字"),
    ],
)
def test_bad_constraint_rejected_at_config_time(tmp_path: Path, snippet: str, keyword: str) -> None:
    with pytest.raises(ConfigError) as excinfo:
        _config(tmp_path, snippet)
    assert keyword in str(excinfo.value)


def test_constraint_text_is_human_readable() -> None:
    assert VariableDef(name="a", min=0, max=50).constraint_text == "范围: 0 ~ 50"
    assert VariableDef(name="a", min=0).constraint_text == "范围: >= 0"
    assert VariableDef(name="a", max=0).constraint_text == "范围: <= 0"
    assert VariableDef(name="a", choices=["EXT", "INT"]).constraint_text == "可选: EXT / INT"
    assert VariableDef(name="a", pattern="[A-Z]+").constraint_text == "格式: [A-Z]+"


def test_number_bounds_accept_numeric_strings() -> None:
    variable = VariableDef(name="a", type="float", min="0.5", max="1.5")
    assert (variable.min, variable.max) == (0.5, 1.5)


def test_integer_bounds_are_shown_without_decimal_point() -> None:
    assert VariableDef(name="a", min=0.0, max=20.0).constraint_text == "范围: 0 ~ 20"


def test_empty_default_is_allowed_but_must_be_filled_later() -> None:
    """声明了约束但没给默认值是允许的 —— 表示"这一格必须去表里填"。"""
    variable = VariableDef(name="kind", choices=["EXT", "INT"])
    assert variable.value_problem("") is not None  # 空值不合格


# --------------------------------------------------------------------------- #
# 2. value_problem 的判定细节
# --------------------------------------------------------------------------- #
def test_value_problem_matches_by_text() -> None:
    variable = VariableDef(name="n", choices=[1, 2])
    assert variable.value_problem(1) is None
    assert variable.value_problem("1") is None
    assert variable.value_problem(1.0) is None  # 1.0 与 1 视为同一个值
    assert variable.value_problem(3) is not None


def test_value_problem_numeric_edges() -> None:
    variable = VariableDef(name="d", type="float", min=0, max=50)
    assert variable.value_problem(0) is None  # 边界值算通过
    assert variable.value_problem(50) is None
    assert variable.value_problem(-0.1) is not None
    assert "下限" in (variable.value_problem(-1) or "")
    assert "上限" in (variable.value_problem(51) or "")
    assert "不是数字" in (variable.value_problem("abc") or "")


def test_bool_is_not_a_number() -> None:
    variable = VariableDef(name="d", type="float", min=0, max=1)
    assert variable.value_problem(True) is not None


# --------------------------------------------------------------------------- #
# 3. 生成期：数据有效性
# --------------------------------------------------------------------------- #
def test_create_template_writes_data_validation(tmp_path: Path) -> None:
    config = _config(
        tmp_path,
        "    - name: draft\n      type: float\n      min: 0\n      max: 50\n      default: 20\n",
        "    - name: kind\n      choices: [EXT, INT]\n      default: EXT\n"
        "    - name: n_cyl\n      type: int\n      min: 1\n      max: 8\n      default: 4\n",
    )
    path = create_template(config, tmp_path / "template.xlsx", cases=2, overwrite=True)

    book = load_workbook(path)
    try:
        global_dv = book["Global Parameter"].data_validations.dataValidation
        assert len(global_dv) == 1
        assert global_dv[0].type == "decimal"
        assert (global_dv[0].formula1, global_dv[0].formula2) == ("0", "50")
        assert str(global_dv[0].sqref) == "B2"

        local_dv = book["Local Parameter"].data_validations.dataValidation
        by_type = {dv.type: dv for dv in local_dv}
        assert set(by_type) == {"list", "whole"}
        # 下拉列表：两个 Case 列都覆盖
        assert by_type["list"].formula1 == '"EXT,INT"'
        assert str(by_type["list"].sqref) == "E2 F2"
        # 整数范围用 whole 而不是 decimal
        assert (by_type["whole"].formula1, by_type["whole"].formula2) == ("1", "8")
        assert str(by_type["whole"].sqref) == "E3 F3"

        # 提示语带上约束，鼠标悬停能看到
        assert "可选: EXT / INT" in (by_type["list"].prompt or "")
        assert "范围: 0 ~ 50" in (global_dv[0].prompt or "")
    finally:
        book.close()


def test_variables_without_constraints_get_no_validation(tmp_path: Path) -> None:
    config = _config(tmp_path, "    - name: draft\n      type: float\n      default: 20\n")
    path = create_template(config, tmp_path / "template.xlsx", cases=2, overwrite=True)
    book = load_workbook(path)
    try:
        assert not book["Global Parameter"].data_validations.dataValidation
        assert not book["Local Parameter"].data_validations.dataValidation
    finally:
        book.close()


def test_derived_variables_get_no_validation(tmp_path: Path) -> None:
    config = _config(
        tmp_path,
        "    - name: draft\n      type: float\n      default: 20\n"
        '    - name: half\n      type: float\n      derived: "draft / 2"\n',
    )
    path = create_template(config, tmp_path / "template.xlsx", cases=2, overwrite=True)
    book = load_workbook(path)
    try:
        # 只有 draft 那一行有校验，派生格（第 3 行）没有
        assert [str(dv.sqref) for dv in book["Global Parameter"].data_validations.dataValidation] == []
    finally:
        book.close()


def test_pattern_only_variable_gets_no_validation(tmp_path: Path) -> None:
    """Excel 的数据有效性没有正则 —— 只有 pattern 时不写校验（但工具侧照查）。"""
    config = _config(tmp_path, '    - name: mcu\n      default: STM32F103\n      pattern: "STM32.*"\n')
    path = create_template(config, tmp_path / "template.xlsx", cases=2, overwrite=True)
    book = load_workbook(path)
    try:
        assert not book["Global Parameter"].data_validations.dataValidation
    finally:
        book.close()
    with pytest.raises(ExcelError):
        check_value_constraints(config, {"mcu": VarValue(value="ATMEGA328")}, [])


def test_choices_with_comma_is_rejected(tmp_path: Path) -> None:
    config = _config(tmp_path, '    - name: v\n      choices: ["A,B", C]\n      default: C\n')
    with pytest.raises(ExcelError) as excinfo:
        create_template(config, tmp_path / "template.xlsx", cases=2, overwrite=True)
    assert "逗号" in str(excinfo.value)


def test_too_many_choices_is_rejected(tmp_path: Path) -> None:
    many = ", ".join(f"V{i:03d}" for i in range(60))  # 60 * 5 = 300 字符 > 255
    config = _config(tmp_path, f"    - name: v\n      choices: [{many}]\n      default: V000\n")
    with pytest.raises(ExcelError) as excinfo:
        create_template(config, tmp_path / "template.xlsx", cases=2, overwrite=True)
    assert "255" in str(excinfo.value)


# --------------------------------------------------------------------------- #
# 4. 运行期：check_value_constraints
# --------------------------------------------------------------------------- #
def test_check_value_constraints_reports_global_and_local(tmp_path: Path) -> None:
    config = _config(
        tmp_path,
        "    - name: draft\n      type: float\n      min: 0\n      max: 50\n      default: 20\n",
        "    - name: kind\n      choices: [EXT, INT]\n      default: EXT\n",
    )
    globals_ = {"draft": VarValue(value=205.59)}
    cases = [
        CaseData(name="C1", column=5, values={"kind": VarValue(value="EXT")}),
        CaseData(name="C2", column=6, values={"kind": VarValue(value="FOO")}),
    ]
    with pytest.raises(ExcelError) as excinfo:
        check_value_constraints(config, globals_, cases)
    message = str(excinfo.value)
    assert "共 2 处" in message
    assert "第 2 行 'draft'" in message
    assert "第 F 列 'C2'" in message


def test_check_value_constraints_passes_when_all_good(tmp_path: Path) -> None:
    config = _config(
        tmp_path,
        "    - name: draft\n      type: float\n      min: 0\n      max: 50\n      default: 20\n",
        "    - name: kind\n      choices: [EXT, INT]\n      default: EXT\n",
    )
    check_value_constraints(
        config,
        {"draft": VarValue(value=20)},
        [CaseData(name="C1", column=5, values={"kind": VarValue(value="INT")})],
    )


def test_empty_value_violates_declared_constraints(tmp_path: Path) -> None:
    """空值也算不合格 —— 这正是"新插一个空 Case 列悄悄落进某个规则集"的解药。"""
    config = _config(
        tmp_path,
        "    - name: draft\n      type: float\n      default: 20\n",
        "    - name: kind\n      choices: [EXT, INT]\n",
    )
    with pytest.raises(ExcelError) as excinfo:
        check_value_constraints(
            config,
            {"draft": VarValue(value=20)},
            [CaseData(name="C3", column=7, values={"kind": VarValue(value="")})],
        )
    assert "（空）" in str(excinfo.value)


# --------------------------------------------------------------------------- #
# 5. 端到端：render / validate / check 都拦得住
# --------------------------------------------------------------------------- #
def _end_to_end(tmp_path: Path):
    config = _config(
        tmp_path,
        "    - name: draft\n      type: float\n      min: 0\n      max: 50\n      default: 20\n",
        "    - name: kind\n      choices: [EXT, INT]\n      default: EXT\n",
    )
    excel = tmp_path / "template.xlsx"
    create_template(config, excel, cases=2, overwrite=True)
    return config, excel


def _set_cell(path: Path, sheet: str, coordinate: str, value: object) -> None:
    book = load_workbook(path)
    try:
        book[sheet][coordinate] = value
        book.save(path)
    finally:
        book.close()


def test_render_rejects_out_of_range_value(tmp_path: Path) -> None:
    config, excel = _end_to_end(tmp_path)
    _set_cell(excel, "Global Parameter", "B2", 205.59)
    with pytest.raises(ExcelError) as excinfo:
        render_all(config, excel)
    assert "大于上限 50" in str(excinfo.value)


def test_render_accepts_values_within_range(tmp_path: Path) -> None:
    config, excel = _end_to_end(tmp_path)
    output = render_all(config, excel)
    assert output.results["demo"][0].lines[0] == "// EXT L=20"


def test_cli_validate_and_check_exit_nonzero(tmp_path: Path) -> None:
    from typer.testing import CliRunner

    from spreadsheet_codegen.cli import app
    from tests.test_cli import output_of

    _config, excel = _end_to_end(tmp_path)
    _set_cell(excel, "Local Parameter", "E2", "FOO")  # kind 在第 2 行、E 列

    runner = CliRunner()
    for command in ("validate", "check"):
        result = runner.invoke(app, [command, "-c", str(tmp_path / "demo.yaml"), "-x", str(excel)])
        assert result.exit_code == 1, f"{command} 应当以退出码 1 结束"
        assert "不在允许列表" in output_of(result)
