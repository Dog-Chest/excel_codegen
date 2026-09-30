"""``extends``：把多个 YAML 合成一份项目配置。

判据与 ``abs_fpi/compose.py`` 保持一致：决定生成文本的字段（``prefix`` / ``suffix`` /
``type`` / ``derived`` / 取值约束）必须逐字一致；``default`` 与 ``description`` 不一致只告警；
同名模板内容必须一致。这里额外覆盖 compose.py 没做的事 —— **``template_file`` 相对声明它的
那个文件解析**（被 extends 的文件常常在子目录里）。
"""

from __future__ import annotations

from pathlib import Path

import pytest

from spreadsheet_codegen import create_template, load_config, render_all
from spreadsheet_codegen.utils import ConfigError

HEAD = "version: 1\n"
GLOBAL_L = '  global:\n    - name: L\n      type: float\n      default: 300\n      suffix: " m"\n'
NO_VARS = "variables:\n  global: []\n  local: []\n"
NO_TEMPLATES = "templates: []\n"

#: 项目文件里默认给一个局部变量和一个模板 —— 合并后的配置必须至少各有一个才合法
DEFAULT_LOCAL = "  local:\n    - name: kind\n      default: EXT\n"
DEFAULT_TEMPLATES = '  - name: own\n    output_sheet: "Code"\n    code: |\n      // own {{ kind }}\n'


def write(path: Path, text: str) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")
    return path


def project(
    path: Path,
    *,
    extends: list[str],
    global_block: str = "  global: []\n",
    local_block: str = DEFAULT_LOCAL,
    templates: str = DEFAULT_TEMPLATES,
) -> Path:
    lines = [
        "version: 1",
        "extends:",
        *(f"  - {item}" for item in extends),
        "excel:",
        '  output: "project.xlsx"',
        "  template_sheet: null",
        "  howto_sheet: null",
        "  sheets:",
        '    global: "Global Parameter"',
        '    local: "Local Parameter"',
        '    outputs: ["Code"]',
        "variables:",
        global_block.rstrip("\n"),
        local_block.rstrip("\n"),
        "templates:",
        templates.rstrip("\n"),
    ]
    return write(path, "\n".join(lines) + "\n")


def count_warnings(config, keyword: str) -> int:
    return sum(1 for w in config.load_warnings if keyword in w)


# --------------------------------------------------------------------------- #
# 合并语义
# --------------------------------------------------------------------------- #
def test_merges_variables_and_templates(tmp_path: Path) -> None:
    write(
        tmp_path / "a.yaml",
        HEAD + "variables:\n" + GLOBAL_L + "  local:\n    - name: kind\n      default: EXT\n"
        'templates:\n  - name: code_a\n    output_sheet: "Code"\n    code: |\n      // a {{ L }}\n',
    )
    write(
        tmp_path / "b.yaml",
        HEAD + "variables:\n" + GLOBAL_L + "    - name: g\n      type: float\n      default: 9.81\n"
        '  local: []\ntemplates:\n  - name: code_b\n    output_sheet: "Code"\n    code: |\n      // b {{ g }}\n',
    )
    path = project(
        tmp_path / "project.yaml",
        extends=["a.yaml", "b.yaml"],
        global_block=GLOBAL_L + "    - name: extra\n      default: X\n",
        templates=DEFAULT_TEMPLATES,
    )

    config = load_config(path)
    assert [v.name for v in config.global_variables] == ["L", "g", "extra"]
    assert [v.name for v in config.local_variables] == ["kind"]
    assert [t.name for t in config.templates] == ["code_a", "code_b", "own"]
    assert config.load_warnings == []


def test_identical_variable_in_two_files_is_deduped(tmp_path: Path) -> None:
    for name in ("a.yaml", "b.yaml"):
        write(tmp_path / name, HEAD + "variables:\n" + GLOBAL_L + "  local: []\n" + NO_TEMPLATES)
    config = load_config(project(tmp_path / "project.yaml", extends=["a.yaml", "b.yaml"]))
    assert [v.name for v in config.global_variables] == ["L"]


def test_identical_template_in_two_files_is_deduped(tmp_path: Path) -> None:
    body = 'templates:\n  - name: same\n    output_sheet: "Code"\n    code: |\n      // x\n'
    for name in ("a.yaml", "b.yaml"):
        write(tmp_path / name, HEAD + NO_VARS + body)
    config = load_config(project(tmp_path / "project.yaml", extends=["a.yaml", "b.yaml"]))
    assert [t.name for t in config.templates] == ["same", "own"]


def test_nested_extends_is_recursive(tmp_path: Path) -> None:
    write(tmp_path / "base.yaml", HEAD + "variables:\n" + GLOBAL_L + "  local: []\n" + NO_TEMPLATES)
    write(
        tmp_path / "mid.yaml",
        HEAD + "extends:\n  - base.yaml\nvariables:\n  global:\n    - name: g\n      default: 9.81\n"
        "  local: []\ntemplates: []\n",
    )
    config = load_config(project(tmp_path / "project.yaml", extends=["mid.yaml"]))
    assert [v.name for v in config.global_variables] == ["L", "g"]


# --------------------------------------------------------------------------- #
# 冲突判定
# --------------------------------------------------------------------------- #
def test_strict_field_conflict_is_rejected(tmp_path: Path) -> None:
    write(tmp_path / "a.yaml", HEAD + "variables:\n" + GLOBAL_L + "  local: []\n" + NO_TEMPLATES)
    write(
        tmp_path / "b.yaml",
        HEAD + "variables:\n  global:\n    - name: L\n      type: float\n      default: 300\n"
        '      suffix: " [m]"\n  local: []\ntemplates: []\n',
    )
    with pytest.raises(ConfigError) as excinfo:
        load_config(project(tmp_path / "project.yaml", extends=["a.yaml", "b.yaml"]))
    assert "suffix" in str(excinfo.value)


def test_constraint_conflict_is_rejected(tmp_path: Path) -> None:
    write(
        tmp_path / "a.yaml",
        HEAD + "variables:\n  global:\n    - name: d\n      min: 0\n      max: 50\n  local: []\n" + NO_TEMPLATES,
    )
    write(
        tmp_path / "b.yaml",
        HEAD + "variables:\n  global:\n    - name: d\n      min: 0\n      max: 60\n  local: []\n" + NO_TEMPLATES,
    )
    with pytest.raises(ConfigError) as excinfo:
        load_config(project(tmp_path / "project.yaml", extends=["a.yaml", "b.yaml"]))
    assert "max" in str(excinfo.value)


def test_loose_field_conflict_only_warns(tmp_path: Path) -> None:
    write(
        tmp_path / "a.yaml",
        HEAD
        + "variables:\n  global:\n    - name: d\n      default: 1\n      description: A\n  local: []\n"
        + NO_TEMPLATES,
    )
    write(
        tmp_path / "b.yaml",
        HEAD
        + "variables:\n  global:\n    - name: d\n      default: 2\n      description: B\n  local: []\n"
        + NO_TEMPLATES,
    )
    config = load_config(project(tmp_path / "project.yaml", extends=["a.yaml", "b.yaml"]))
    assert count_warnings(config, "default") == 1
    assert count_warnings(config, "description") == 1
    assert config.global_variables[0].default == 1  # 保留先出现的


def test_template_name_clash_with_different_body_is_rejected(tmp_path: Path) -> None:
    write(
        tmp_path / "a.yaml",
        HEAD + NO_VARS + 'templates:\n  - name: t\n    output_sheet: "Code"\n    code: |\n      // one\n',
    )
    write(
        tmp_path / "b.yaml",
        HEAD + NO_VARS + 'templates:\n  - name: t\n    output_sheet: "Code"\n    code: |\n      // two\n',
    )
    with pytest.raises(ConfigError) as excinfo:
        load_config(project(tmp_path / "project.yaml", extends=["a.yaml", "b.yaml"]))
    assert "模板" in str(excinfo.value)


# --------------------------------------------------------------------------- #
# 错误的 extends 写法
# --------------------------------------------------------------------------- #
def test_missing_file_is_rejected(tmp_path: Path) -> None:
    with pytest.raises(ConfigError) as excinfo:
        load_config(project(tmp_path / "project.yaml", extends=["nope.yaml"]))
    assert "不存在" in str(excinfo.value)


def test_cycle_is_rejected(tmp_path: Path) -> None:
    write(tmp_path / "a.yaml", HEAD + "extends:\n  - b.yaml\n" + NO_VARS + NO_TEMPLATES)
    write(tmp_path / "b.yaml", HEAD + "extends:\n  - a.yaml\n" + NO_VARS + NO_TEMPLATES)
    with pytest.raises(ConfigError) as excinfo:
        load_config(project(tmp_path / "project.yaml", extends=["a.yaml"]))
    assert "循环引用" in str(excinfo.value)


def test_extends_must_be_a_list_of_paths(tmp_path: Path) -> None:
    path = write(tmp_path / "project.yaml", HEAD + "extends: not-a-list\n" + NO_VARS + NO_TEMPLATES)
    with pytest.raises(ConfigError) as excinfo:
        load_config(path)
    assert "extends" in str(excinfo.value)


def test_included_excel_block_is_ignored_with_warning(tmp_path: Path) -> None:
    write(
        tmp_path / "a.yaml",
        HEAD + 'excel:\n  output: "other.xlsx"\n  sheets:\n    global: "G"\n    local: "L"\n'
        '    outputs: ["Code"]\n' + NO_VARS + NO_TEMPLATES,
    )
    config = load_config(project(tmp_path / "project.yaml", extends=["a.yaml"]))
    assert count_warnings(config, "excel") == 1
    assert Path(config.excel.output).name == "project.xlsx"


# --------------------------------------------------------------------------- #
# template_file 相对**声明它的文件**解析（compose.py 做不到的那件事）
# --------------------------------------------------------------------------- #
def test_template_file_resolves_relative_to_declaring_file(tmp_path: Path) -> None:
    write(tmp_path / "rules" / "tpl" / "body.j2", "// from subdir: {{ case_name }} L={{ L }}\n")
    write(
        tmp_path / "rules" / "a.yaml",
        HEAD + "variables:\n" + GLOBAL_L + "  local: []\ntemplates:\n  - name: sub\n"
        '    output_sheet: "Code"\n    template_file: "tpl/body.j2"\n',
    )
    config = load_config(project(tmp_path / "project.yaml", extends=["rules/a.yaml"]))

    assert config.templates[0].name == "sub"
    assert config.templates[0].source_dir == (tmp_path / "rules").resolve()
    excel = create_template(config, tmp_path / "project.xlsx", cases=["C1"], overwrite=True)
    assert render_all(config, excel).results["sub"][0].lines == ["// from subdir: C1 L=300 m"]


def test_root_template_file_still_resolves_relative_to_root(tmp_path: Path) -> None:
    write(tmp_path / "top.j2", "// root file {{ case_name }}\n")
    write(tmp_path / "a.yaml", HEAD + NO_VARS + NO_TEMPLATES)
    path = project(
        tmp_path / "project.yaml",
        extends=["a.yaml"],
        templates='  - name: own\n    output_sheet: "Code"\n    template_file: "top.j2"\n',
    )
    config = load_config(path)
    assert config.templates[0].source_dir is None  # 根自己的模板不特殊标记，走 config.source_dir
    excel = create_template(config, tmp_path / "project.xlsx", cases=["C1"], overwrite=True)
    assert render_all(config, excel).results["own"][0].lines == ["// root file C1"]


# --------------------------------------------------------------------------- #
# 不带 extends 时行为不变
# --------------------------------------------------------------------------- #
def test_without_extends_nothing_changes(tmp_path: Path, config_text: str) -> None:
    path = write(tmp_path / "plain.yaml", config_text)
    config = load_config(path)
    assert config.extends == []
    assert config.load_warnings == []
    assert config.templates[0].source_dir is None
    assert config.source_dir == tmp_path.resolve()


def test_extends_list_is_kept_for_introspection(tmp_path: Path) -> None:
    write(tmp_path / "a.yaml", HEAD + NO_VARS + NO_TEMPLATES)
    config = load_config(project(tmp_path / "project.yaml", extends=["a.yaml"]))
    assert config.extends == ["a.yaml"]
    assert config.source_dir == tmp_path.resolve()
