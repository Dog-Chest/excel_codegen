"""派生表达式（``derived:``）的一个**静默失效**家族：``x.value`` 从未真正生效过。

背景（0.11.0 修的，见 CHANGELOG）
--------------------------------
``derived:`` 表达式在两种情形下的解析器曾经都只接受一个参数 —— 而
``derived.to_excel`` 翻译 ``x.value`` 这类属性访问时**必须**把属性名传下去
（``resolve(name, attribute)``）。解析器只收一个参数就抛 ``TypeError``，
翻译器把它当成"上下文不支持属性访问"、降级成"往格子里写算好的值"：

* 公式模式的核心承诺没了 —— 参数格子里的派生值**改输入不会自动重算**；
* 而且**一句提示都没有**，只有 ``--write-excel`` 重新算一次才跟得上。

同一件事还让 ``is_translatable()`` 把任何带 ``.value`` 的表达式判成"翻译不了"。

顺带把 ``len`` 支持上（最常用、最容易补齐）：此前 ``derived: "len(x.value)"``
会被报成"引用了未定义的变量 'len'" —— 那句话把人的心智带向完全不同的修法。
"""

from __future__ import annotations

from pathlib import Path

import pytest
from openpyxl import load_workbook

from spreadsheet_codegen.derived import (
    DerivedError,
    is_translatable,
    to_excel,
    untranslatable_names,
)
from spreadsheet_codegen.excel_io import create_template, write_results
from spreadsheet_codegen.models import ProjectConfig, load_config
from spreadsheet_codegen.renderer import render_all

#: ``secret`` 是文本、``secret_len`` 量它的长度 —— 正好卡在"必须写 .value"的位置
LEN_YAML = """\
version: 1

excel:
  output: "len.xlsx"
  template_sheet: null
  howto_sheet: null
  sheets:
    global: "Global Parameter"
    local: "Local Parameter"
    outputs: ["Output"]

variables:
  local:
    - name: secret
      default: "abcdef"
      type: string
    - name: secret_len
      derived: "len(secret.value)"
      type: int

templates:
  - name: row
    output_sheet: "Output"
    start_cell: "B2"
    code: |
      len={{ secret_len }}
"""


@pytest.fixture()
def len_config(tmp_path: Path) -> ProjectConfig:
    path = tmp_path / "len.yaml"
    path.write_text(LEN_YAML, encoding="utf-8")
    return load_config(path)


# --------------------------------------------------------------------------- #
# len：能用，而且要能翻成 Excel 公式
# --------------------------------------------------------------------------- #
def test_len_is_translatable(len_config: ProjectConfig) -> None:
    """``len(x.value)`` 要能翻译成 Excel 的 ``LEN()``，不是"只能写算好的值"。"""
    variable = len_config.local_variables[1]
    assert variable.derived == "len(secret.value)"
    assert is_translatable(variable) is True, "len() 应当能翻成 Excel 的 LEN()"
    assert untranslatable_names(len_config) == []


def test_len_compiles_to_len_call() -> None:
    """翻译形态：``len(secret.value)`` -> ``LEN(<取值>)``。"""
    assert to_excel("len(secret.value)", name="t", resolve=lambda name, attribute=None: f"<{name}.{attribute}>") == (
        "LEN(<secret.value>)"
    )
    # 过滤器写法等价
    assert to_excel("secret.value | len", name="t", resolve=lambda name, attribute=None: "<v>") == "LEN(<v>)"
    assert to_excel("secret.value | length", name="t", resolve=lambda name, attribute=None: "<v>") == "LEN(<v>)"
    # 参与算术
    assert (
        to_excel("len(a.value) + 1", name="t", resolve=lambda name, attribute=None: "<a.value>") == "(LEN(<a.value>)+1)"
    )


def test_len_cell_is_a_live_formula(len_config: ProjectConfig, tmp_path: Path) -> None:
    """参数表里的派生格必须是**公式**（``data_type == "f"``）。

    这条正是那个静默 bug 的守门人：以前这里会是一个算好的数字，
    于是"改了 secret 的长度，格子不会自己变"——而且没有任何提示。
    """
    path = create_template(len_config, tmp_path / "len.xlsx", cases=1)
    workbook = load_workbook(path)
    try:
        cell = workbook["Local Parameter"]["E3"]  # secret_len 的格子
        assert cell.data_type == "f", f"应当是公式，实际是 {cell.data_type}: {cell.value!r}"
        assert cell.value.upper().startswith("=LEN(")
    finally:
        workbook.close()


def test_len_formula_follows_input_change(len_config: ProjectConfig, tmp_path: Path) -> None:
    """改了 secret，Python 侧与 Excel 公式侧**同时**跟上（这才是公式模式的意义）。"""
    from spreadsheet_codegen.formula_eval import Evaluator, WorkbookReader, sheet_names_of

    path = create_template(len_config, tmp_path / "len.xlsx", cases=1)
    workbook = load_workbook(path)
    try:
        sheet = workbook["Local Parameter"]
        row = next(r for r in range(2, sheet.max_row + 1) if sheet.cell(row=r, column=1).value == "secret")
        sheet.cell(row=row, column=5).value = "abcdefghij"  # 10 个字符
        workbook.save(path)
    finally:
        workbook.close()

    output = render_all(len_config, path)
    assert output.cases[0].values["secret_len"].value == 10
    assert "len=10" in output.results["row"][0].text

    # 公式文本没变，但把**公式在 Python 里算一遍**应当得到新的长度（证明它是活公式）
    workbook = load_workbook(path)
    try:
        sheet = workbook["Local Parameter"]
        derived_row = next(r for r in range(2, sheet.max_row + 1) if sheet.cell(row=r, column=1).value == "secret_len")
        formula = sheet.cell(row=derived_row, column=5).value
        reader = WorkbookReader(workbook)
        evaluator = Evaluator(reader, {name: name for name in sheet_names_of(len_config)})
        assert float(evaluator.evaluate(formula)) == 10
    finally:
        workbook.close()


def test_len_rejects_wrong_arity(len_config: ProjectConfig) -> None:
    """``len()`` 的参数个数不对时，配置期给出可操作的报错。"""
    from spreadsheet_codegen.derived import validate_config

    variable = len_config.local_variables[1]
    variable.derived = "len()"
    with pytest.raises(DerivedError, match="只接受一个参数"):
        validate_config(len_config)
    variable.derived = "len(secret.value, secret.value)"
    with pytest.raises(DerivedError, match="只接受一个参数"):
        validate_config(len_config)


def test_len_bare_variable_measures_the_pure_value(value_config: ProjectConfig, tmp_path: Path) -> None:
    """``len(x)`` 与 ``len(x.value)`` **等价** —— 都是纯值的长度，与 Excel 逐字对齐。

    指南 §15.1 写的是"引用拿到的是**纯值**（``x.value`` 语义）"，Excel 侧的
    ``resolve(name)`` 也确实指向取值列。所以 ``len(port)`` 量的是 ``"A"``（1），
    不是组合值 ``"GPIOA_PORT"``（10）。

    回归：0.11.0 早期把派生上下文包成了模板侧的 ``VarValue``（``str(x)`` 是组合值），
    于是同一个表达式两边给出两个答案 —— Python 10 / Excel 1。这条测试**真的去把
    生成的 Excel 公式算一遍**，两边的数必须相等（只断言 Pythbn 侧那个数就是过拟合，
    正是它让不一致溜了过去）。
    """
    from spreadsheet_codegen.formula_eval import Evaluator, WorkbookReader, sheet_names_of

    value_config.local_variables[1].derived = "len(port)"
    path = create_template(value_config, tmp_path / "bare.xlsx", cases=1)
    output = render_all(value_config, path)
    write_results(path, value_config, output.results)

    python_side = output.cases[0].values["port_name"].value
    assert python_side == 1, "port='A'（纯值）的长度是 1；组合值 GPIOA_PORT 才是 10"

    workbook = load_workbook(path)
    try:
        sheet = workbook["Local Parameter"]
        row = next(r for r in range(2, sheet.max_row + 1) if sheet.cell(row=r, column=1).value == "port_name")
        cell = sheet.cell(row=row, column=5)
        assert cell.data_type == "f", f"应当是活公式，实际 {cell.data_type}: {cell.value!r}"

        evaluator = Evaluator(WorkbookReader(workbook), {name: name for name in sheet_names_of(value_config)})
        excel_side = float(evaluator.evaluate(cell.value))
        assert excel_side == python_side, f"Excel 算出 {excel_side}，Python 算出 {python_side} —— 两边必须一致"
    finally:
        workbook.close()


def test_bare_reference_and_dot_value_agree_in_derived(value_config: ProjectConfig, tmp_path: Path) -> None:
    """``port ~ port`` 用的是**纯值**：Python 与 Excel 都得到 ``"AA"``。

    回归：包成模板侧 ``VarValue`` 时 Python 给 ``"GPIOA_PORTGPIOA_PORT"``，
    而 Excel 的 ``&`` 拼的是取值列（``"A"``）—— 同一个表达式两个答案。
    """
    from spreadsheet_codegen.formula_eval import Evaluator, WorkbookReader, sheet_names_of

    value_config.local_variables[1].derived = "port ~ port"
    path = create_template(value_config, tmp_path / "cat.xlsx", cases=1)
    output = render_all(value_config, path)
    write_results(path, value_config, output.results)

    assert output.cases[0].values["port_name"].value == "AA"

    workbook = load_workbook(path)
    try:
        sheet = workbook["Local Parameter"]
        row = next(r for r in range(2, sheet.max_row + 1) if sheet.cell(row=r, column=1).value == "port_name")
        cell = sheet.cell(row=row, column=5)
        assert cell.data_type == "f", f"应当是活公式，实际 {cell.data_type}: {cell.value!r}"
        evaluator = Evaluator(WorkbookReader(workbook), {name: name for name in sheet_names_of(value_config)})
        assert evaluator.evaluate(cell.value) == "AA"
    finally:
        workbook.close()


def test_len_of_literal_and_expression(len_config: ProjectConfig) -> None:
    """常量与"表达式结果"的长度都能翻译（只有裸变量名被拒）。"""
    resolve = lambda name, attribute=None: f"<{name}.{attribute}>"  # noqa: E731
    assert to_excel("len('abc')", name="t", resolve=resolve) == 'LEN("abc")'
    assert to_excel("len(a.value + b.value)", name="t", resolve=resolve) == "LEN((<a.value>+<b.value>))"


# --------------------------------------------------------------------------- #
# x.value / x.prefix / x.suffix 在派生表达式里必须真的生效
# --------------------------------------------------------------------------- #
VALUE_YAML = """\
version: 1

excel:
  output: "v.xlsx"
  template_sheet: null
  howto_sheet: null
  sheets:
    global: "Global Parameter"
    local: "Local Parameter"
    outputs: ["Output"]

variables:
  local:
    - name: port
      default: "A"
      prefix: "GPIO"
      suffix: "_PORT"
    - name: port_name
      derived: "port.value ~ '_PIN'"
    - name: port_len
      derived: "len(port.value)"
      type: int
    - name: port_prefixed_len
      derived: "len(port.value) + len(port.prefix)"
      type: int

templates:
  - name: row
    output_sheet: "Output"
    start_cell: "B2"
    code: |
      {{ port_name }} len={{ port_len }} tot={{ port_prefixed_len }}
"""


@pytest.fixture()
def value_config(tmp_path: Path) -> ProjectConfig:
    path = tmp_path / "value.yaml"
    path.write_text(VALUE_YAML, encoding="utf-8")
    return load_config(path)


def test_attribute_access_in_derived_is_translatable(value_config: ProjectConfig) -> None:
    """带 ``.value`` / ``.prefix`` 的派生表达式都能写成公式（回归：以前全被判成翻译不了）。"""
    assert untranslatable_names(value_config) == []
    for variable in value_config.local_variables[1:]:
        assert is_translatable(variable) is True, variable.derived


def test_attribute_access_in_derived_produces_formulas(value_config: ProjectConfig, tmp_path: Path) -> None:
    """三个派生格都应当是活公式，并且值算得对。"""
    path = create_template(value_config, tmp_path / "v.xlsx", cases=1)
    output = render_all(value_config, path)
    write_results(path, value_config, output.results)

    values = output.cases[0].values
    assert str(values["port"]) == "GPIOA_PORT"
    assert values["port_name"].value == "A_PIN"
    assert values["port_len"].value == 1
    assert values["port_prefixed_len"].value == 5  # 1 + len("GPIO")

    workbook = load_workbook(path)
    try:
        sheet = workbook["Local Parameter"]
        rows = {sheet.cell(row=r, column=1).value: r for r in range(2, sheet.max_row + 1)}
        for name in ("port_name", "port_len", "port_prefixed_len"):
            cell = sheet.cell(row=rows[name], column=5)
            assert cell.data_type == "f", f"{name} 应当是公式，实际 {cell.data_type}: {cell.value!r}"
    finally:
        workbook.close()


def test_derived_formula_varies_with_prefix(value_config: ProjectConfig, tmp_path: Path) -> None:
    """改前缀，``len(port.value) + len(port.prefix)`` 在公式里要跟着变（不是写死的值）。"""
    from spreadsheet_codegen.formula_eval import evaluate_template_values

    path = create_template(value_config, tmp_path / "v.xlsx", cases=1)
    output = render_all(value_config, path)
    write_results(path, value_config, output.results)
    workbook = load_workbook(path)
    try:
        sheet = workbook["Local Parameter"]
        row = next(r for r in range(2, sheet.max_row + 1) if sheet.cell(row=r, column=1).value == "port")
        sheet.cell(row=row, column=3).value = "GPIOX"  # Prefix 列 GPIO -> GPIOX
        workbook.save(path)
    finally:
        workbook.close()

    output = render_all(value_config, path)
    assert output.cases[0].values["port_prefixed_len"].value == 6  # 1 + 5

    workbook = load_workbook(path)
    try:
        got = evaluate_template_values(workbook, value_config, value_config.templates[0], ["Case1"])
    finally:
        workbook.close()
    assert got["Case1"] == output.results["row"][0].lines
    assert "tot=6" in got["Case1"][0]


def test_unknown_attribute_is_a_config_error(value_config: ProjectConfig, tmp_path: Path) -> None:
    """属性拼错要**报错**，不能静默降级成"写算好的值"（那会悄悄丢掉自动重算）。"""
    value_config.local_variables[1].derived = "port.valu"
    with pytest.raises(DerivedError, match=r"\.valu"):
        create_template(value_config, tmp_path / "bad.xlsx", cases=1)
    # 属性名拼错时不能"看起来成功"：报错路径不留下工作簿
    assert not (tmp_path / "bad.xlsx").exists()


# --------------------------------------------------------------------------- #
# 两个引擎必须给同一个答案：跨作用域的前后缀、以及数值比较
# --------------------------------------------------------------------------- #
CROSS_SCOPE_YAML = """\
version: 1

excel:
  output: "x.xlsx"
  template_sheet: null
  howto_sheet: null
  sheets:
    global: "Global Parameter"
    local: "Local Parameter"
    outputs: ["Output"]

variables:
  global:
    - name: baud
      default: 115200
      prefix: "GP"
  local:
    - name: port
      default: "A"
    - name: prefix_len
      derived: "len(baud.prefix)"
      type: int

templates:
  - name: row
    output_sheet: "Output"
    start_cell: "B2"
    code: |
      {{ prefix_len }}
"""


def test_local_derived_reads_global_prefix_from_the_sheet(tmp_path: Path) -> None:
    """local 派生引用 global 的 ``.prefix`` 时，取的是 **Global 表里那一格**，不是 YAML。

    回归：``_resolve_derived_locals`` 只把 local 的 values 交给 ``_derived_context``，
    global 名字在里面取不到就回落到 YAML 的前缀 —— 而公式侧 ``resolve`` 读的是
    Global 表的 D 列。表里改过前缀，Python 与 Excel 就各算一个数。
    """
    config_path = tmp_path / "x.yaml"
    config_path.write_text(CROSS_SCOPE_YAML, encoding="utf-8")
    config = load_config(config_path)
    path = create_template(config, tmp_path / "x.xlsx", cases=1)

    workbook = load_workbook(path)
    try:
        sheet = workbook["Global Parameter"]
        row = next(r for r in range(2, sheet.max_row + 1) if sheet.cell(row=r, column=1).value == "baud")
        sheet.cell(row=row, column=4).value = "LONGPREFIX"  # D 列 = Prefix
        workbook.save(path)
    finally:
        workbook.close()

    output = render_all(config, path)
    assert output.cases[0].values["prefix_len"].value == len("LONGPREFIX"), "要按表里读到的前缀算，不是 YAML 的 'GP'"


NUMERIC_YAML = """\
version: 1

excel:
  output: "n.xlsx"
  template_sheet: null
  howto_sheet: null
  sheets:
    global: "Global Parameter"
    local: "Local Parameter"
    outputs: ["Output"]

variables:
  global:
    - name: a
      type: int
      default: 10
    - name: b
      type: int
      default: 9
  local:
    - name: port
      default: "A"
    - name: smaller
      derived: "min(a, b)"
      type: int
    - name: bigger
      derived: "max(a, b)"
      type: int
    - name: is_less
      derived: "int(a < b)"
      type: int

templates:
  - name: row
    output_sheet: "Output"
    start_cell: "B2"
    code: |
      {{ smaller }} {{ bigger }} {{ is_less }}
"""


def test_numeric_comparison_matches_excel(tmp_path: Path) -> None:
    """数值比较要**按值**比，不是按文本 —— 否则 ``min`` / ``max`` / ``<`` 全是错的。

    回归：排序运算符早先用 ``to_text`` 逐字符比字符串，于是 ``10`` 与 ``9`` 比出
    ``"10" < "9"`` → ``min(10, 9)`` 得到 **10**，而 Excel 的 ``MIN(10,9)`` 是 **9**。
    ``min`` / ``max`` / ``<`` 都在公式翻译白名单里，两边必须一致。

    这条测试**真的把生成的 Excel 公式算一遍**再比（只断言 Python 侧会漏掉这类分叉）。

    这里用 ``int(a < b)`` 而不是裸的 ``a < b``：布尔的**文本形式**两边还不一致
    （Python ``to_text(False)`` = ``"false"``、Excel 拼出来是 ``"FALSE"``），
    那是另一个独立的问题，不该混进这条回归里。
    """
    from spreadsheet_codegen.formula_eval import Evaluator, WorkbookReader, sheet_names_of

    config_path = tmp_path / "n.yaml"
    config_path.write_text(NUMERIC_YAML, encoding="utf-8")
    config = load_config(config_path)
    path = create_template(config, tmp_path / "n.xlsx", cases=1)
    output = render_all(config, path)
    write_results(path, config, output.results)

    assert output.cases[0].values["smaller"].value == 9
    assert output.cases[0].values["bigger"].value == 10
    assert output.cases[0].values["is_less"].value == 0, "a < b 按值比是假（10 < 9）"

    workbook = load_workbook(path)
    try:
        sheet = workbook["Local Parameter"]
        rows = {sheet.cell(row=r, column=1).value: r for r in range(2, sheet.max_row + 1)}
        evaluator = Evaluator(WorkbookReader(workbook), {name: name for name in sheet_names_of(config)})
        assert float(evaluator.evaluate(sheet.cell(row=rows["smaller"], column=5).value)) == 9.0
        assert float(evaluator.evaluate(sheet.cell(row=rows["bigger"], column=5).value)) == 10.0
        assert float(evaluator.evaluate(sheet.cell(row=rows["is_less"], column=5).value)) == 0.0
    finally:
        workbook.close()
