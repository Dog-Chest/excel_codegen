"""公式模式的行内 ``{% if %}``：编译成 Excel ``IF()``，并且**两种引擎的结果必须一致**。

这个文件里的端到端用例（``test_both_engines_agree`` / ``test_check_verifies_if_formulas``）
是重点：`check` 会把 Excel 里的公式在 Python 里算一遍再与快照渲染逐行比对 ——
公式编译器与公式求值器是两套独立实现，两边都过才算"Excel 里看到的"与"导出的"一致。
"""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_codegen import create_template, load_config, render_all
from excel_codegen.excel_io import write_results
from excel_codegen.formula import FormulaError, compile_formulas
from excel_codegen.models import ProjectConfig

IF_YAML = """\
version: 1

excel:
  output: "if.xlsx"
  template_sheet: null
  howto_sheet: null
  sheets:
    global: "Global Parameter"
    local: "Local Parameter"
    outputs: ["Output"]

variables:
  global:
    - name: L
      type: float
      default: 300
      suffix: " m"
  local:
    - name: flag
      type: int
      default: 1
    - name: kind
      default: "EXT"

templates:
  - name: demo
    output_sheet: "Output"
    engine: "excel"
    code: |
__CODE__
"""


def make_config(tmp_path: Path, code: str) -> ProjectConfig:
    """把 ``code`` 塞进模板（每行统一缩进 6 格，块标量的缩进由第一行决定）。"""
    indented = "\n".join("      " + line if line.strip() else "" for line in code.splitlines())
    path = tmp_path / "if.yaml"
    path.write_text(IF_YAML.replace("__CODE__", indented), encoding="utf-8")
    return load_config(path)


def compile_code(config: ProjectConfig, code: str | None = None) -> list[str]:
    template = config.templates[0]
    if code is not None:
        template = template.model_copy(update={"code": code})
    return compile_formulas(template, config, case_columns=[5])[0]


# --------------------------------------------------------------------------- #
# 编译形态
# --------------------------------------------------------------------------- #
def test_if_else_is_compiled_to_if(tmp_path: Path) -> None:
    config = make_config(tmp_path, "x = {% if flag.value == 1 %}A{% else %}B{% endif %};")
    formula = compile_code(config)[0]
    assert formula.startswith('="x = "&IF(')
    assert formula.endswith('"A","B")&";"')


def test_if_without_else_falls_back_to_empty(tmp_path: Path) -> None:
    config = make_config(tmp_path, "x = {% if flag.value == 1 %}A{% endif %};")
    assert compile_code(config)[0].endswith('"A","")&";"')


def test_condition_uses_raw_value_column(tmp_path: Path) -> None:
    """条件里写 ``.value`` 就该取"取值"列（Local 表的 Case 列），而不是前缀 / 后缀列。"""
    config = make_config(tmp_path, 'x = {% if kind.value == "EXT" %}A{% endif %}')
    formula = compile_code(config)[0]
    assert "INDEX('Local Parameter'!E:E," in formula
    assert "INDEX('Local Parameter'!C:C," not in formula
    assert "INDEX('Local Parameter'!D:D," not in formula


def test_condition_can_use_prefix(tmp_path: Path) -> None:
    config = make_config(tmp_path, 'x = {% if kind.prefix == "" %}A{% endif %}')
    assert "INDEX('Local Parameter'!$C:$C," in compile_code(config)[0]


def test_nested_if(tmp_path: Path) -> None:
    config = make_config(
        tmp_path,
        'x{% if flag.value == 1 %}{% if kind.value == "EXT" %}both{% else %}flag only{% endif %}{% endif %}',
    )
    formula = compile_code(config)[0]
    assert '"both","flag only"' in formula  # 内层 IF 的两个分支
    assert formula.count("<>") == 0  # 条件都是显式比较，没有真假包装


def test_jinja_comment_is_dropped(tmp_path: Path) -> None:
    config = make_config(tmp_path, "x = 1{# 这是注释，不该出现在结果里 #};")
    formula = compile_code(config)[0]
    assert "注释" not in formula
    assert formula == '="x = 1"&";"'


def test_case_name_and_template_name_are_allowed_bare(tmp_path: Path) -> None:
    config = make_config(
        tmp_path,
        'a = {% if case_name == "C1" %}first{% endif %}\nb = {% if template_name == "demo" %}yes{% endif %}',
    )
    formulas = compile_code(config)
    assert "'Local Parameter'!E$1=\"C1\"" in formulas[0]
    assert '"demo"="demo"' in formulas[1]


@pytest.mark.parametrize(
    "condition",
    [
        "flag.value == 1",
        "flag.value != 1",
        "flag.value > 0",
        "flag.value < 2",
        "flag.value >= 1",
        "flag.value <= 1",
    ],
)
def test_comparison_operators(tmp_path: Path, condition: str) -> None:
    config = make_config(tmp_path, f"x = {{% if {condition} %}}A{{% endif %}}")
    assert "IF(" in compile_code(config)[0]


def test_and_or_not_are_translated(tmp_path: Path) -> None:
    config = make_config(
        tmp_path,
        'a = {% if flag.value > 0 and kind.value == "EXT" %}A{% endif %}\n'
        'b = {% if flag.value == 0 or kind.value == "EXT" %}B{% endif %}\n'
        "c = {% if not flag.value %}C{% endif %}",
    )
    formulas = compile_code(config)
    assert "AND(" in formulas[0]
    assert "OR(" in formulas[1]
    assert "NOT(" in formulas[2]


def test_truthiness_wrap_depends_on_type(tmp_path: Path) -> None:
    """数值变量比 0，文本变量比空串 —— 与 Python 侧 ``bool(value)`` 的直觉一致。"""
    config = make_config(
        tmp_path,
        "a = {% if flag.value %}x{% endif %}\nb = {% if kind.value %}x{% endif %}",
    )
    formulas = compile_code(config)
    assert "<>0)" in formulas[0]
    assert '<>""' in formulas[1]


# --------------------------------------------------------------------------- #
# 报错（都要指出问题在哪）
# --------------------------------------------------------------------------- #
def test_bare_variable_in_condition_is_rejected(tmp_path: Path) -> None:
    config = make_config(tmp_path, "x = {% if flag %}A{% endif %}")
    with pytest.raises(FormulaError) as excinfo:
        compile_code(config)
    message = str(excinfo.value)
    assert ".value" in message
    assert "裸写" in message


def test_multiline_if_is_rejected(tmp_path: Path) -> None:
    """一行模板 = 一个单元格，跨行分支会改变行数，映射不到固定单元格。"""
    config = make_config(tmp_path, "{% if flag.value == 1 %}\nX\n{% endif %}")
    with pytest.raises(FormulaError) as excinfo:
        compile_code(config)
    assert "没有闭合" in str(excinfo.value)
    assert "同一行" in str(excinfo.value)


def test_for_loop_is_still_rejected(tmp_path: Path) -> None:
    config = make_config(tmp_path, "{% for x in [1] %}X{% endfor %}")
    with pytest.raises(FormulaError) as excinfo:
        compile_code(config)
    assert "不支持" in str(excinfo.value)


def test_stray_endif_is_rejected(tmp_path: Path) -> None:
    config = make_config(tmp_path, "x{% endif %}")
    with pytest.raises(FormulaError) as excinfo:
        compile_code(config)
    assert "没有对应的" in str(excinfo.value)


def test_unclosed_tag_is_rejected(tmp_path: Path) -> None:
    config = make_config(tmp_path, "x = {{ flag.value")
    with pytest.raises(FormulaError) as excinfo:
        compile_code(config)
    assert "没有闭合" in str(excinfo.value)


def test_unsupported_attribute_in_condition(tmp_path: Path) -> None:
    config = make_config(tmp_path, "x = {% if flag.nope %}A{% endif %}")
    with pytest.raises(FormulaError) as excinfo:
        compile_code(config)
    assert "不支持 .nope" in str(excinfo.value)


def test_non_boolean_condition_shape_is_rejected(tmp_path: Path) -> None:
    config = make_config(tmp_path, "x = {% if flag.value + 1 %}A{% endif %}")
    with pytest.raises(FormulaError) as excinfo:
        compile_code(config)
    assert "必须是比较或逻辑表达式" in str(excinfo.value)


def test_empty_condition_is_rejected(tmp_path: Path) -> None:
    config = make_config(tmp_path, "x = {% if %}A{% endif %}")
    with pytest.raises(FormulaError) as excinfo:
        compile_code(config)
    assert "没有条件" in str(excinfo.value)


# --------------------------------------------------------------------------- #
# 端到端：两种引擎必须给出同一段代码
# --------------------------------------------------------------------------- #
IF_BODY = """\
// Case {{ case_name }}
y = {% if flag.value == 1 %}{{ L }}{% else %}- {{ L }}{% endif %};
x = {% if kind.value == "EXT" %}+{% else %}-{% endif %}t{{ L.value }};
z = {% if flag.value %}has{% else %}none{% endif %};
w = {% if flag.value > 0 and kind.value == "EXT" %}both{% else %}neither{% endif %};
v = {% if not flag.value %}off{% endif %};
"""


def test_both_engines_agree(tmp_path: Path) -> None:
    """同一份模板：快照渲染的文本，与公式模式写进 Excel 的公式，逐行一致。"""
    config = make_config(tmp_path, IF_BODY)
    excel = create_template(config, tmp_path / "if.xlsx", cases=["EXT-1", "INT-1"], overwrite=True)

    # 第二个 Case 走 else 分支
    from openpyxl import load_workbook

    book = load_workbook(excel)
    try:
        book["Local Parameter"]["F2"] = 0
        book["Local Parameter"]["F3"] = "INT"
        book.save(excel)
    finally:
        book.close()

    snapshot = render_all(config, excel)
    expected_first = [
        "// Case EXT-1",
        "y = 300 m;",
        "x = +t300;",
        "z = has;",
        "w = both;",
        "v = ;",
    ]
    expected_second = [
        "// Case INT-1",
        "y = - 300 m;",
        "x = -t300;",
        "z = none;",
        "w = neither;",
        "v = off;",
    ]
    assert snapshot.results["demo"][0].lines == expected_first
    assert snapshot.results["demo"][1].lines == expected_second


def test_check_verifies_if_formulas(tmp_path: Path) -> None:
    """``check --values`` 把 Excel 公式算一遍再比 —— 编译器与求值器是两套独立实现。"""
    config = make_config(tmp_path, IF_BODY)
    excel = create_template(config, tmp_path / "if.xlsx", cases=["EXT-1", "INT-1"], overwrite=True)

    from openpyxl import load_workbook

    book = load_workbook(excel)
    try:
        book["Local Parameter"]["F2"] = 0
        book["Local Parameter"]["F3"] = "INT"
        book.save(excel)
    finally:
        book.close()

    write_results(excel, config, render_all(config, excel).results)

    from typer.testing import CliRunner

    from excel_codegen.cli import app

    result = CliRunner().invoke(app, ["check", "-c", str(tmp_path / "if.yaml"), "-x", str(excel)])
    assert result.exit_code == 0, result.output
