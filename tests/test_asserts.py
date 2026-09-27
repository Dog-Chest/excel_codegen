"""跨变量校验 ``asserts``：拦"吃水不能超过型深"这类**组合**错误。

单变量约束（``min`` / ``max`` / ``choices`` / ``pattern``，见 test_constraints.py）只能看一列；
``asserts`` 对**每个 Case** 求值一段表达式，能引用同 Case 的多个变量。
写法与 ``case_filter`` 完全一致：裸变量是组合值，**数值比较请写 ``.value``**。
"""

from __future__ import annotations

from pathlib import Path

import pytest
from openpyxl import load_workbook
from typer.testing import CliRunner

from excel_codegen import create_template, load_config, render_all
from excel_codegen.cli import app
from excel_codegen.excel_io import check_value_constraints
from excel_codegen.models import CaseData
from excel_codegen.renderer import check_asserts, compile_asserts
from excel_codegen.utils import ConfigError, ExcelError, RenderError, VarValue

runner = CliRunner()

BASE = """\
version: 1

excel:
  output: "a.xlsx"
  template_sheet: null
  howto_sheet: null
  sheets:
    global: "Global Parameter"
    local: "Local Parameter"
    outputs: ["Output"]

variables:
  global:
    - name: d_tank
      type: float
      default: 24
  local:
    - name: draft
      type: float
      default: 20
    - name: kind
      choices: ["EXT", "INT"]
      default: "EXT"

__ASSERTS__

templates:
  - name: demo
    output_sheet: "Output"
    code: |
      // {{ kind }} {{ draft }}
"""


def make_config(tmp_path: Path, asserts: list[str] | str | None = None):
    if asserts is None:
        block = "asserts: []"
    elif isinstance(asserts, str):
        block = asserts
    else:
        block = "asserts:\n" + "".join(f"  - '{item.replace(chr(39), chr(39) * 2)}'\n" for item in asserts)
    path = tmp_path / "a.yaml"
    path.write_text(BASE.replace("__ASSERTS__", block), encoding="utf-8")
    return load_config(path)


def cases(*specs: tuple[str, int, float, str]) -> list[CaseData]:
    return [
        CaseData(
            name=name,
            column=column,
            values={"draft": VarValue(value=draft), "kind": VarValue(value=kind)},
        )
        for name, column, draft, kind in specs
    ]


GLOBALS = {"d_tank": VarValue(value=24)}


# --------------------------------------------------------------------------- #
# 配置期
# --------------------------------------------------------------------------- #
def test_asserts_default_to_empty(tmp_path: Path) -> None:
    assert make_config(tmp_path).asserts == []


def test_asserts_are_stripped(tmp_path: Path) -> None:
    config = make_config(tmp_path, 'asserts:\n  - "  draft.value <= d_tank.value  "\n')
    assert config.asserts == ["draft.value <= d_tank.value"]


@pytest.mark.parametrize(
    "block",
    [
        "asserts: not-a-list",
        "asserts:\n  - ''\n",
        "asserts:\n  - '   '\n",
    ],
)
def test_bad_asserts_are_rejected(tmp_path: Path, block: str) -> None:
    with pytest.raises(ConfigError):
        make_config(tmp_path, block)


def test_syntax_error_is_reported(tmp_path: Path) -> None:
    config = make_config(tmp_path, ["draft.value <="])
    with pytest.raises(RenderError) as excinfo:
        compile_asserts(config)
    assert "asserts 第 1 条语法错误" in str(excinfo.value)


# --------------------------------------------------------------------------- #
# 求值
# --------------------------------------------------------------------------- #
def test_passing_asserts_are_silent(tmp_path: Path) -> None:
    config = make_config(tmp_path, ["draft.value <= d_tank.value"])
    check_asserts(config, GLOBALS, cases(("A", 5, 20, "EXT"), ("B", 6, 24, "INT")))


def test_violation_names_the_case_and_column(tmp_path: Path) -> None:
    config = make_config(tmp_path, ["draft.value <= d_tank.value"])
    with pytest.raises(ExcelError) as excinfo:
        check_asserts(config, GLOBALS, cases(("A", 5, 20, "EXT"), ("B", 6, 30, "INT")))
    message = str(excinfo.value)
    assert "共 1 处" in message
    assert "Case 'B'" in message and "第 F 列" in message
    assert "draft.value <= d_tank.value" in message


def test_multiple_asserts_and_cases_are_all_reported(tmp_path: Path) -> None:
    config = make_config(
        tmp_path,
        ["draft.value <= d_tank.value", "not (kind.value == 'INT' and draft.value > 22)"],
    )
    bad = cases(("A", 5, 30, "EXT"), ("B", 6, 23, "INT"))
    with pytest.raises(ExcelError) as excinfo:
        check_asserts(config, GLOBALS, bad)
    message = str(excinfo.value)
    assert "共 2 处" in message
    assert "draft.value <= d_tank.value" in message
    assert "not (kind.value == 'INT'" in message


def test_bare_name_is_the_combined_value(tmp_path: Path) -> None:
    """与 case_filter 同一套语义：裸变量是组合值，数值比较必须写 .value。"""
    config = make_config(tmp_path, ['kind == "EXT"'])
    check_asserts(config, GLOBALS, cases(("A", 5, 20, "EXT")))
    with pytest.raises(ExcelError):
        check_asserts(config, GLOBALS, cases(("A", 5, 20, "INT")))


def test_unknown_variable_gives_a_clear_error(tmp_path: Path) -> None:
    config = make_config(tmp_path, ["nope.value > 1"])
    with pytest.raises(RenderError) as excinfo:
        check_asserts(config, GLOBALS, cases(("A", 5, 20, "EXT")))
    assert "求值失败" in str(excinfo.value)


def test_no_asserts_is_a_noop(tmp_path: Path) -> None:
    config = make_config(tmp_path)
    assert compile_asserts(config) == []
    check_asserts(config, GLOBALS, cases(("A", 5, 999, "EXT")))  # 不报错


# --------------------------------------------------------------------------- #
# 与 extends 的关系
# --------------------------------------------------------------------------- #
def test_asserts_from_extended_files_are_kept(tmp_path: Path) -> None:
    (tmp_path / "rules").mkdir()
    (tmp_path / "rules" / "a.yaml").write_text(
        "version: 1\nasserts:\n  - 'draft.value <= d_tank.value'\n"
        "variables:\n  global: []\n  local: []\ntemplates: []\n",
        encoding="utf-8",
    )
    (tmp_path / "rules" / "b.yaml").write_text(
        "version: 1\nasserts:\n  - 'draft.value >= 0'\nvariables:\n  global: []\n  local: []\ntemplates: []\n",
        encoding="utf-8",
    )
    config = make_config(tmp_path, ["draft.value <= d_tank.value"])
    (tmp_path / "a.yaml").write_text(
        BASE.replace("__ASSERTS__", "asserts:\n  - 'draft.value <= d_tank.value'\n").replace(
            "version: 1\n", "version: 1\nextends:\n  - rules/a.yaml\n  - rules/b.yaml\n", 1
        ),
        encoding="utf-8",
    )
    config = load_config(tmp_path / "a.yaml")
    assert config.asserts == [
        "draft.value <= d_tank.value",
        "draft.value >= 0",
        "draft.value <= d_tank.value",
    ]


# --------------------------------------------------------------------------- #
# 端到端：render / validate / check 都拦得住
# --------------------------------------------------------------------------- #
def test_render_and_validate_reject_violation(tmp_path: Path) -> None:
    config = make_config(tmp_path, ["draft.value <= d_tank.value"])
    excel = tmp_path / "a.xlsx"
    create_template(config, excel, cases=["A", "B"], overwrite=True, include_scripts=False)

    book = load_workbook(excel)
    try:
        book["Local Parameter"]["F2"] = 30  # B 列的 draft 超过 d_tank
        book.save(excel)
    finally:
        book.close()

    with pytest.raises(ExcelError):
        render_all(config, excel)
    for command in ("validate", "check"):
        result = runner.invoke(app, [command, "-c", str(tmp_path / "a.yaml"), "-x", str(excel)])
        assert result.exit_code == 1, f"{command} 应当以退出码 1 结束"
        text = ((result.output or "") + (getattr(result, "stderr", "") or "")).replace("\n", "")
        assert "asserts" in text


def test_constraint_error_wins_over_assert(tmp_path: Path) -> None:
    """同一个值同时违反单变量约束与跨变量规则时，先报更具体的那条（约束）。"""
    path = tmp_path / "a.yaml"
    path.write_text(
        BASE.replace("__ASSERTS__", "asserts:\n  - 'draft.value <= d_tank.value'\n").replace(
            "      type: float\n      default: 20", "      type: float\n      default: 20\n      min: 0\n      max: 22"
        ),
        encoding="utf-8",
    )
    config = load_config(path)
    bad = [CaseData(name="A", column=5, values={"draft": VarValue(value=30), "kind": VarValue(value="EXT")})]
    with pytest.raises(ExcelError) as excinfo:
        check_value_constraints(config, GLOBALS, bad)
    assert "大于上限" in str(excinfo.value)
