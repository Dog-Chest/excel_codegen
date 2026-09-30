"""``spreadsheet-codegen doctor``：一条命令体检环境 / 配置 / 工作簿。

它的价值在"把散在各处的常见坑一次说清"，所以测试关注三件事：
① 正常项目退出码 0 且该报的"值得留意"都报了；② 真有问题时退出码 1 且**不甩 traceback**；
③ 只该是警告的事（没 uv、没渲染记录、没约束）不能升级成错误。
"""

from __future__ import annotations

from pathlib import Path

import pytest
from openpyxl import load_workbook
from typer.testing import CliRunner

from spreadsheet_codegen import create_template, load_config, render_all
from spreadsheet_codegen.cli import app
from spreadsheet_codegen.excel_io import write_results

runner = CliRunner()

CONFIG = """\
version: 1

excel:
  output: "t.xlsx"
  template_sheet: "Template"
  howto_sheet: "HOWTO"
  sheets:
    global: "Global Parameter"
    local: "Local Parameter"
    outputs: ["Output"]

variables:
  global:
    - name: d_tank
      type: float
      default: 24
      min: 1
  local:
    - name: draft
      type: float
      default: 20
      min: 0
    - name: unused_one
      default: "x"

asserts:
  - "draft.value <= d_tank.value"

templates:
  - name: demo
    output_sheet: "Output"
    code: |
      // {{ draft }}
"""


def text_of(result) -> str:
    return ((result.output or "") + (getattr(result, "stderr", "") or "")).replace("\n", "")


@pytest.fixture()
def config_path(tmp_path: Path) -> Path:
    path = tmp_path / "t.yaml"
    path.write_text(CONFIG, encoding="utf-8")
    return path


def run_doctor(path: Path, excel: Path | None = None):
    args = ["doctor", "-c", str(path)]
    if excel is not None:
        args += ["-x", str(excel)]
    return runner.invoke(app, args)


def make_book(config_path: Path, *, render: bool = False) -> Path:
    config = load_config(config_path)
    excel = config_path.parent / "t.xlsx"
    create_template(config, excel, cases=["A", "B"], overwrite=True, include_scripts=False)
    if render:
        write_results(excel, config, render_all(config, excel).results)
    return excel


# --------------------------------------------------------------------------- #
# 正常路径
# --------------------------------------------------------------------------- #
def test_doctor_on_healthy_project(config_path: Path) -> None:
    excel = make_book(config_path, render=True)
    result = run_doctor(config_path, excel)
    assert result.exit_code == 0, text_of(result)
    text = text_of(result)
    assert "Python" in text and "依赖" in text
    assert "加载成功" in text
    assert "与当前一致" in text
    assert "体检完成" in text


def test_doctor_excel_defaults_to_config_output(config_path: Path, monkeypatch) -> None:
    """不传 -x 时用配置里的 excel.output；它是相对 CWD 解析的（与 init / render 一致）。"""
    make_book(config_path, render=True)
    monkeypatch.chdir(config_path.parent)
    result = run_doctor(config_path)  # 不传 -x
    assert result.exit_code == 0
    assert "与当前一致" in text_of(result)


def test_missing_workbook_is_only_a_warning(config_path: Path) -> None:
    result = run_doctor(config_path)  # 还没 init
    assert result.exit_code == 0
    assert "不存在" in text_of(result)


def test_missing_render_record_is_only_a_warning(config_path: Path) -> None:
    excel = make_book(config_path, render=False)
    result = run_doctor(config_path, excel)
    assert result.exit_code == 0
    assert "还没跑过" in text_of(result)


def test_missing_constraints_is_reported_as_a_hint(tmp_path: Path) -> None:
    path = tmp_path / "bare.yaml"
    path.write_text(
        CONFIG.replace("      min: 1\n", "")
        .replace("      min: 0\n", "")
        .replace('asserts:\n  - "draft.value <= d_tank.value"\n', "asserts: []\n"),
        encoding="utf-8",
    )
    excel = make_book(path, render=True)
    result = run_doctor(path, excel)
    assert result.exit_code == 0
    assert "一条约束都没有" in text_of(result)


def test_unused_variable_is_reported(tmp_path: Path) -> None:
    excel = make_book(config_path=tmp_path / "t.yaml") if False else None  # 占位，见下
    path = tmp_path / "u.yaml"
    path.write_text(CONFIG, encoding="utf-8")
    excel = make_book(path, render=True)
    result = run_doctor(path, excel)
    assert result.exit_code == 0
    text = text_of(result)
    assert "unused_one" in text
    # 保留名（case_name / template_name）不该被当成"定义了没人用"
    assert "template_name" not in text


# --------------------------------------------------------------------------- #
# 真问题：退出码 1，但必须是可读结论而不是 traceback
# --------------------------------------------------------------------------- #
def test_value_violation_is_an_error_not_a_crash(tmp_path: Path) -> None:
    path = tmp_path / "t.yaml"
    path.write_text(CONFIG, encoding="utf-8")
    excel = make_book(path, render=True)

    book = load_workbook(excel)
    try:
        book["Local Parameter"]["E2"] = 99  # draft 超过 d_tank
        book.save(excel)
    finally:
        book.close()

    result = run_doctor(path, excel)
    assert result.exit_code == 1
    text = text_of(result)
    assert "asserts" in text
    assert "Traceback" not in text
    assert "不满足" in text


def test_broken_yaml_is_reported_cleanly(tmp_path: Path) -> None:
    path = tmp_path / "bad.yaml"
    path.write_text("version: 1\nvariables: {global: [], local: []}\n", encoding="utf-8")
    result = run_doctor(path)
    assert result.exit_code == 1
    text = text_of(result)
    assert "YAML" in text and "ERROR" in text
    assert "Traceback" not in text


def test_missing_fragment_is_reported_as_template_error(tmp_path: Path) -> None:
    path = tmp_path / "f.yaml"
    path.write_text(CONFIG.replace("      // {{ draft }}\n", '      {% include "nope.j2" %}\n'), encoding="utf-8")
    result = run_doctor(path)
    assert result.exit_code == 1
    text = text_of(result)
    assert "找不到片段" in text
    assert "Traceback" not in text
