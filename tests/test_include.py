"""``{% include %}`` 模板片段复用（快照模式）。

``extends`` 解决的是"跨 **YAML 文件**复用变量与模板"；片段复用解决的是"跨 **模板文件**
复用一段代码"。片段路径**相对声明模板的那个文件**解析 —— 与 ``template_file`` 同一套规则，
所以 ``extends`` 进来的模板也能正确找到自己旁边的片段。

公式模式不支持（一行模板 = 一个单元格，片段会让行数不可预测）。
"""

from __future__ import annotations

from pathlib import Path

import pytest
from typer.testing import CliRunner

from excel_codegen import create_template, load_config, render_all
from excel_codegen.cli import app
from excel_codegen.excel_io import write_results
from excel_codegen.formula import FormulaError, compile_formulas
from excel_codegen.jinja_env import build_environment
from excel_codegen.renderer import collect_variables, render_template, validate_template
from excel_codegen.utils import RenderError

runner = CliRunner()

HEAD = """\
version: 1

excel:
  output: "p.xlsx"
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
    - name: unit_name
      default: "m"
  local:
    - name: kind
      default: "EXT"
"""


def write(path: Path, text: str) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")
    return path


def project(tmp_path: Path, code_body: str, *, engine: str | None = None) -> Path:
    engine_line = f"    engine: {engine}\n" if engine else ""
    return write(
        tmp_path / "p.yaml",
        HEAD
        + 'templates:\n  - name: demo\n    output_sheet: "Output"\n'
        + engine_line
        + "    code: |\n"
        + "".join(f"      {line}\n" for line in code_body.splitlines()),
    )


# --------------------------------------------------------------------------- #
# 基本行为
# --------------------------------------------------------------------------- #
def test_include_expands_fragment(tmp_path: Path) -> None:
    write(tmp_path / "frag" / "body.j2", "// 片段：{{ L }}\n")
    path = project(tmp_path, '{% include "frag/body.j2" %}\nvar x = {{ L }};')
    config = load_config(path)

    text = render_template(config.templates[0], {"L": 300}, base_dir=config.source_dir)
    assert "// 片段：300" in text
    assert "var x = 300;" in text


def test_nested_include(tmp_path: Path) -> None:
    write(tmp_path / "frag" / "outer.j2", 'outer\n{% include "frag/inner.j2" %}\n')
    write(tmp_path / "frag" / "inner.j2", "inner {{ L }}\n")
    path = project(tmp_path, '{% include "frag/outer.j2" %}')
    config = load_config(path)
    text = render_template(config.templates[0], {"L": 7}, base_dir=config.source_dir)
    assert "outer" in text and "inner 7" in text


def test_variables_inside_fragments_are_collected(tmp_path: Path) -> None:
    """片段里的变量也算"被引用" —— 否则 validate 会误报"定义了没人用"、也不查定义。"""
    write(tmp_path / "frag.j2", "{{ unit_name }}\n")
    path = project(tmp_path, '{% include "frag.j2" %}\n{{ L }}')
    config = load_config(path)
    used = collect_variables(config.templates[0], base_dir=config.source_dir)
    assert {"unit_name", "L"} <= used


def test_render_all_works_end_to_end(tmp_path: Path) -> None:
    write(tmp_path / "frag.j2", "// {{ case_name }} {{ L }}")
    path = project(tmp_path, '{% include "frag.j2" %}')
    config = load_config(path)
    excel = create_template(config, tmp_path / "p.xlsx", cases=["A", "B"], overwrite=True, include_scripts=False)
    output = render_all(config, excel)
    assert output.results["demo"][0].lines == ["// A 300"]
    assert output.results["demo"][1].lines == ["// B 300"]


# --------------------------------------------------------------------------- #
# 与 extends / template_file 的关系
# --------------------------------------------------------------------------- #
def test_fragment_resolves_relative_to_declaring_file(tmp_path: Path) -> None:
    """``extends`` 进来的模板，片段相对**它自己**所在目录解析。"""
    write(tmp_path / "rules" / "frag" / "body.j2", "// from rules: {{ L }}\n")
    write(
        tmp_path / "rules" / "rule.yaml",
        "version: 1\nvariables:\n  global:\n    - name: L\n      type: float\n      default: 300\n  local: []\n"
        'templates:\n  - name: from_rule\n    output_sheet: "Output"\n    code: |\n'
        '      {% include "frag/body.j2" %}\n',
    )
    write(
        tmp_path / "p.yaml",
        HEAD.replace("version: 1\n", "version: 1\nextends:\n  - rules/rule.yaml\n", 1) + "templates: []\n",
    )
    config = load_config(tmp_path / "p.yaml")
    template = next(t for t in config.templates if t.name == "from_rule")
    assert template.source_dir == (tmp_path / "rules").resolve()
    excel = create_template(config, tmp_path / "p.xlsx", cases=["A"], overwrite=True, include_scripts=False)
    assert render_all(config, excel).results["from_rule"][0].lines[0] == "// from rules: 300"


# --------------------------------------------------------------------------- #
# 报错
# --------------------------------------------------------------------------- #
def test_missing_fragment_is_reported(tmp_path: Path) -> None:
    path = project(tmp_path, '{% include "frag/nope.j2" %}')
    config = load_config(path)
    with pytest.raises(RenderError) as excinfo:
        validate_template(config.templates[0], base_dir=config.source_dir)
    message = str(excinfo.value)
    assert "找不到片段" in message
    assert "frag/nope.j2" in message
    assert "§17" in message


def test_fragment_syntax_error_points_at_the_fragment(tmp_path: Path) -> None:
    write(tmp_path / "bad.j2", "{% if %}\n")
    path = project(tmp_path, '{% include "bad.j2" %}')
    config = load_config(path)
    with pytest.raises(RenderError) as excinfo:
        validate_template(config.templates[0], base_dir=config.source_dir)
    assert "'bad.j2'" in str(excinfo.value)


def test_formula_mode_rejects_include(tmp_path: Path) -> None:
    path = project(tmp_path, '{% include "frag.j2" %}', engine="excel")
    config = load_config(path)
    with pytest.raises(FormulaError) as excinfo:
        compile_formulas(config.templates[0], config, case_axes=[5])
    message = str(excinfo.value)
    assert "include" in message
    assert "engine: snapshot" in message


def test_loaderless_environment_says_what_to_do(tmp_path: Path) -> None:
    """显式传一个没有搜索路径的环境时，要给得出"该怎么办"。"""
    write(tmp_path / "frag.j2", "x\n")
    path = project(tmp_path, '{% include "frag.j2" %}')
    config = load_config(path)
    with pytest.raises(RenderError) as excinfo:
        validate_template(config.templates[0], env=build_environment())
    assert "没有搜索路径" in str(excinfo.value)


# --------------------------------------------------------------------------- #
# 端到端：check 会把片段算进指纹
# --------------------------------------------------------------------------- #
def test_check_notices_fragment_change(tmp_path: Path) -> None:
    """改了片段 → 输出过期 —— 说明片段确实参与了渲染与指纹。"""
    fragment = write(tmp_path / "frag.j2", "// v1 {{ L }}")
    path = project(tmp_path, '{% include "frag.j2" %}')
    config = load_config(path)
    excel = create_template(config, tmp_path / "p.xlsx", cases=["A"], overwrite=True, include_scripts=False)
    write_results(excel, config, render_all(config, excel).results)

    fresh = runner.invoke(app, ["check", "-c", str(path), "-x", str(excel)])
    assert fresh.exit_code == 0

    fragment.write_text("// v2 {{ L }}", encoding="utf-8")
    stale = runner.invoke(app, ["check", "-c", str(path), "-x", str(excel)])
    assert stale.exit_code == 1
    assert "第 1 行不同" in ((stale.output or "") + (getattr(stale, "stderr", "") or "")).replace("\n", "")
