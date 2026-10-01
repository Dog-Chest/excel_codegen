"""``prefill`` / ``fallback``：把"预填提示值"与"空单元格回落默认值"分开。

要解决的问题（改进建议 §2.1，附录 A4 有复现）
--------------------------------------------
作者写 ``default`` 的本意常常只是"给新表一个提示值"，但运行时语义是"空 = 用默认值"。
于是一旦某个字段是"某类型才有"的（比如网站的 ``url``、银行卡的 ``card_type``），
``default`` 就从"体贴"变成"污染"：银行卡那列明明清空了，导出结果里却回落到
``https://example.com/login`` —— **没填**与**填了这个值**在结果里不可区分，
生成的是错数据，而且不报错。

两个开关：
* ``prefill: false``  —— 新表里这一格留空（不写 ``default``）；
* ``fallback: false`` —— 单元格为空时**不回落到 default**（空就是空）。

**默认两个都开**（与 0.10.0 完全一致，已有 YAML 一行都不用改）。
"""

from __future__ import annotations

from pathlib import Path

import pytest
from openpyxl import load_workbook

from spreadsheet_codegen.excel_io import case_axis_map, create_template, template_source
from spreadsheet_codegen.formula import compile_formulas
from spreadsheet_codegen.formula_eval import Evaluator, WorkbookReader, sheet_names_of
from spreadsheet_codegen.models import ProjectConfig, VariableDef, load_config
from spreadsheet_codegen.renderer import render_all

#: 一个"按类型可选"的字段：url 只有网站登录才有
OPTIONAL_FIELD_YAML = """\
version: 1

excel:
  output: "vault.xlsx"
  template_sheet: null
  howto_sheet: null
  sheets:
    global: "Global Parameter"
    local: "Local Parameter"
    outputs: ["Output"]

variables:
  local:
    - name: kind
      default: "网站登录"
      type: string
    - name: url
      default: "https://example.com/login"
      type: string
{extra}

templates:
  - name: row
    output_sheet: "Output"
    start_cell: "B2"
    filename: "{{{{ case_name }}}}.txt"
    code: |
      {{{{ kind }}}} => url=[{{{{ url }}}}]
"""


def _write_config(tmp_path: Path, extra: str) -> ProjectConfig:
    path = tmp_path / "vault.yaml"
    path.write_text(OPTIONAL_FIELD_YAML.format(extra=extra), encoding="utf-8")
    return load_config(path)


def _fill_and_clear(path: Path, *, sheet_url_for_first: str = "https://github.com/login") -> None:
    """site_a 填 url，bank_b **清空** —— 表示"这条记录没有这个字段"。"""
    workbook = load_workbook(path)
    try:
        sheet = workbook["Local Parameter"]
        row = next(r for r in range(2, sheet.max_row + 1) if sheet.cell(row=r, column=1).value == "url")
        sheet.cell(row=row, column=5).value = sheet_url_for_first
        sheet.cell(row=row, column=6).value = None
        workbook.save(path)
    finally:
        workbook.close()


# --------------------------------------------------------------------------- #
# 默认：老行为一字不变
# --------------------------------------------------------------------------- #
def test_defaults_keep_the_old_behaviour(tmp_path: Path) -> None:
    """什么都不写 = 预填 + 回落（0.10.0 的行为），已有配置不受影响。"""
    variable = VariableDef.model_validate({"name": "url", "default": "x"})
    assert variable.prefill is True
    assert variable.fallback is True
    assert variable.effective_default == "x"

    config = _write_config(tmp_path, "")
    path = create_template(config, tmp_path / "vault.xlsx", cases=["site_a", "bank_b"])
    workbook = load_workbook(path)
    try:
        sheet = workbook["Local Parameter"]
        row = next(r for r in range(2, sheet.max_row + 1) if sheet.cell(row=r, column=1).value == "url")
        assert sheet.cell(row=row, column=5).value == "https://example.com/login"  # 预填了
    finally:
        workbook.close()

    _fill_and_clear(path)
    output = render_all(config, path)
    texts = [result.text.strip() for result in output.results["row"]]
    assert texts[0] == "网站登录 => url=[https://github.com/login]"
    assert texts[1] == "网站登录 => url=[https://example.com/login]"  # 清空了也回落 —— 老行为


# --------------------------------------------------------------------------- #
# fallback: false —— "清空 = 没有这个字段"
# --------------------------------------------------------------------------- #
def test_fallback_false_keeps_blank_blank(tmp_path: Path) -> None:
    """清空的格子就真的空：不再回落默认值（Python 侧）。"""
    config = _write_config(tmp_path, "      fallback: false")
    path = create_template(config, tmp_path / "vault.xlsx", cases=["site_a", "bank_b"])
    _fill_and_clear(path)

    output = render_all(config, path)
    texts = [result.text.strip() for result in output.results["row"]]
    assert texts[0] == "网站登录 => url=[https://github.com/login]"
    assert texts[1] == "网站登录 => url=[]", "bank_b 没有 url，不该被默认值填上"


def test_fallback_false_still_prefills_the_suggestion(tmp_path: Path) -> None:
    """``fallback: false`` 只关"回落"，``default`` 仍然预填进新表当提示值。"""
    config = _write_config(tmp_path, "      fallback: false")
    path = create_template(config, tmp_path / "vault.xlsx", cases=["site_a"])
    workbook = load_workbook(path)
    try:
        sheet = workbook["Local Parameter"]
        row = next(r for r in range(2, sheet.max_row + 1) if sheet.cell(row=r, column=1).value == "url")
        assert sheet.cell(row=row, column=5).value == "https://example.com/login"
    finally:
        workbook.close()


def test_fallback_false_agrees_with_the_excel_formula(tmp_path: Path) -> None:
    """**两种引擎必须一致**：Excel 公式对空单元格同样算成空串。

    这一条是关键 —— 只改 Python 侧会让 check 报"公式算出来的文本与 Python 渲染不一致"，
    或者更糟：打开 Excel 看到的是默认值、导出的文件里是空的。
    """
    config = _write_config(tmp_path, "      fallback: false")
    path = create_template(config, tmp_path / "vault.xlsx", cases=["site_a", "bank_b"])
    _fill_and_clear(path)

    output = render_all(config, path)
    python_lines = {result.case_name: result.lines for result in output.results["row"]}

    workbook = load_workbook(path)
    try:
        axes = case_axis_map(workbook, config)
        per_case = compile_formulas(
            config.templates[0],
            config,
            case_axes=[axes["site_a"], axes["bank_b"]],
            source=template_source(config.templates[0], None),
        )
        reader = WorkbookReader(workbook)
        evaluator = Evaluator(reader, {name: name for name in sheet_names_of(config)})
        for case_name, formulas in zip(["site_a", "bank_b"], per_case, strict=False):
            got = [evaluator.evaluate(formula) for formula in formulas]
            assert got == python_lines[case_name], f"{case_name}: Excel={got} Python={python_lines[case_name]}"
    finally:
        workbook.close()


# --------------------------------------------------------------------------- #
# prefill: false —— 新表里不留提示值
# --------------------------------------------------------------------------- #
def test_prefill_false_leaves_the_cell_empty(tmp_path: Path) -> None:
    """``prefill: false``：新表格里那一格是空的，而且**空就是空**。

    ``fallback`` 默认跟着 ``prefill``：既然没往格子里写过默认值，"空单元格回落到默认值"
    就没有意义（用户会以为表里那个值来自某处）。想要"不预填、但空时仍回落"这种组合，
    必须显式写 ``fallback: true``。
    """
    config = _write_config(tmp_path, "      prefill: false")
    path = create_template(config, tmp_path / "vault.xlsx", cases=["site_a"])
    workbook = load_workbook(path)
    try:
        sheet = workbook["Local Parameter"]
        row = next(r for r in range(2, sheet.max_row + 1) if sheet.cell(row=r, column=1).value == "url")
        assert sheet.cell(row=row, column=5).value is None, "prefill: false 不该往格子里写默认值"
    finally:
        workbook.close()

    output = render_all(config, path)
    assert output.results["row"][0].text.strip() == "网站登录 => url=[]"


def test_prefill_false_and_fallback_true_is_rejected_as_contradictory(tmp_path: Path) -> None:
    """``prefill: false`` + ``fallback: true`` 自相矛盾 —— 配置期就报错并说清改法。

    （"不预填、空时却回落"本身没有意义：格子里从没写过那个默认值，
    用户会以为表里那个值来自某处。想要提示值就删掉 ``prefill``。）
    """
    from spreadsheet_codegen.utils import ConfigError

    path = tmp_path / "bad.yaml"
    path.write_text(
        OPTIONAL_FIELD_YAML.format(extra="      prefill: false\n      fallback: true"),
        encoding="utf-8",
    )
    with pytest.raises(ConfigError, match="自相矛盾"):
        load_config(path)


def test_prefill_false_fallback_false_ignores_the_default(tmp_path: Path) -> None:
    """两个都关：``default`` 只对"新表的提示"没意义，值完全由表里决定。"""
    config = _write_config(tmp_path, "      prefill: false\n      fallback: false")
    path = create_template(config, tmp_path / "vault.xlsx", cases=["site_a", "bank_b"])
    workbook = load_workbook(path)
    try:
        sheet = workbook["Local Parameter"]
        row = next(r for r in range(2, sheet.max_row + 1) if sheet.cell(row=r, column=1).value == "url")
        assert sheet.cell(row=row, column=5).value is None
    finally:
        workbook.close()

    output = render_all(config, path)
    texts = [result.text.strip() for result in output.results["row"]]
    assert texts == ["网站登录 => url=[]", "网站登录 => url=[]"]


# --------------------------------------------------------------------------- #
# 配置期：自相矛盾的写法要拦住
# --------------------------------------------------------------------------- #
def test_contradictory_prefill_and_fallback_is_rejected(tmp_path: Path) -> None:
    """``prefill: false`` + ``fallback: true`` 自相矛盾（既不预填、又要空单元格回落）。"""
    from spreadsheet_codegen.utils import ConfigError

    path = tmp_path / "bad.yaml"
    path.write_text(
        OPTIONAL_FIELD_YAML.format(extra="      prefill: false\n      fallback: true"),
        encoding="utf-8",
    )
    with pytest.raises(ConfigError, match="自相矛盾"):
        load_config(path)


def test_derived_variable_rejects_prefill_switches(tmp_path: Path) -> None:
    """派生参数的值是算出来的，``prefill`` / ``fallback`` 对它没有意义。"""
    with pytest.raises(ValueError, match="派生参数"):
        VariableDef.model_validate({"name": "n", "derived": "a + 1", "prefill": False})


def test_comment_explains_blank_semantics(tmp_path: Path) -> None:
    """变量名格子的批注要写清"留空是什么意思" —— 这是这套语义唯一的说明场所。"""
    from spreadsheet_codegen.excel_io import _variable_comment

    fallback_off = VariableDef.model_validate({"name": "url", "default": "x", "fallback": False})
    assert "留空 = 这个参数没有值" in _variable_comment(fallback_off, where="Local 表")

    normal = VariableDef.model_validate({"name": "url", "default": "x"})
    assert "留空 = 使用上面的默认值" in _variable_comment(normal, where="Local 表")

    no_prefill = VariableDef.model_validate({"name": "url", "default": "x", "prefill": False, "fallback": False})
    assert "不预填" in _variable_comment(no_prefill, where="Local 表")
