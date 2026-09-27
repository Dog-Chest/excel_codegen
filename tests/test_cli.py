"""CLI 冒烟测试：init / render / validate / check 四条命令的完整流程，含 FINDINGS 回归。"""

from __future__ import annotations

import io
import re
import sys
from contextlib import suppress
from pathlib import Path

import pytest
from openpyxl import load_workbook
from typer.testing import CliRunner

from excel_codegen.cli import app

runner = CliRunner()


def output_of(result) -> str:
    """兼容不同 click 版本的 stdout/stderr 合并方式。"""
    text = result.output or ""
    # click 8.1 在 mix_stderr 模式下不允许访问 stderr
    with suppress(ValueError, AttributeError):
        text += result.stderr or ""
    return text


def write_config(tmp_path: Path, config_text: str) -> Path:
    path = tmp_path / "example.yaml"
    path.write_text(config_text, encoding="utf-8")
    return path


def test_cli_version() -> None:
    result = runner.invoke(app, ["--version"])
    assert result.exit_code == 0
    assert "excel-codegen" in output_of(result)


def test_cli_init_render_validate(tmp_path: Path, config_text: str) -> None:
    config_path = write_config(tmp_path, config_text)
    excel_path = tmp_path / "template.xlsx"
    outdir = tmp_path / "generated"

    result = runner.invoke(
        app,
        ["init", "--config", str(config_path), "--output", str(excel_path), "--cases", "2"],
    )
    assert result.exit_code == 0, output_of(result)
    assert excel_path.exists()

    # 已存在且未加 --force -> 失败
    result = runner.invoke(app, ["init", "--config", str(config_path), "--output", str(excel_path)])
    assert result.exit_code == 1
    assert "已存在" in output_of(result)

    # 重新生成（覆盖）+ 不生成隐藏 Template 表
    result = runner.invoke(
        app,
        [
            "init",
            "--config",
            str(config_path),
            "--output",
            str(excel_path),
            "--force",
            "--no-template-sheet",
        ],
    )
    assert result.exit_code == 0, output_of(result)

    result = runner.invoke(
        app,
        [
            "render",
            "--config",
            str(config_path),
            "--excel",
            str(excel_path),
            "--write-excel",
            "--outdir",
            str(outdir),
            "--no-show",
        ],
    )
    assert result.exit_code == 0, output_of(result)
    assert (outdir / "uart_init_Case1.c").exists()
    assert (outdir / "uart_summary_Case2.md").exists()
    assert "渲染完成" in output_of(result)

    result = runner.invoke(
        app,
        ["validate", "--config", str(config_path), "--excel", str(excel_path)],
    )
    assert result.exit_code == 0, output_of(result)
    assert "配置校验通过" in output_of(result)


def test_cli_render_missing_excel(tmp_path: Path, config_text: str) -> None:
    config_path = write_config(tmp_path, config_text)
    result = runner.invoke(
        app,
        ["render", "--config", str(config_path), "--excel", str(tmp_path / "nope.xlsx")],
    )
    assert result.exit_code == 1
    assert "不存在" in output_of(result)


def test_cli_render_single_case(tmp_path: Path, config_text: str) -> None:
    config_path = write_config(tmp_path, config_text)
    excel_path = tmp_path / "template.xlsx"
    runner.invoke(app, ["init", "--config", str(config_path), "--output", str(excel_path)])

    result = runner.invoke(
        app,
        [
            "render",
            "--config",
            str(config_path),
            "--excel",
            str(excel_path),
            "--case",
            "Case2",
            "--no-show",
        ],
    )
    assert result.exit_code == 0, output_of(result)
    assert "Case2" in output_of(result)


def test_cli_validate_invalid_yaml(tmp_path: Path) -> None:
    bad = tmp_path / "bad.yaml"
    bad.write_text("version: 1\ntemplates: [\n", encoding="utf-8")
    result = runner.invoke(app, ["validate", "--config", str(bad)])
    assert result.exit_code == 1
    assert "YAML" in output_of(result)


def test_cli_validate_unknown_variable_warns(tmp_path: Path, config_text: str) -> None:
    config_path = write_config(tmp_path, config_text)
    text = config_path.read_text(encoding="utf-8").replace("{{ mcu }}", "{{ not_defined }}")
    config_path.write_text(text, encoding="utf-8")

    result = runner.invoke(app, ["validate", "--config", str(config_path)])
    assert result.exit_code == 0, output_of(result)
    assert "not_defined" in output_of(result)


# --------------------------------------------------------------------------- #
# 回归：FINDINGS.md 里逐条修复的问题
# --------------------------------------------------------------------------- #
def test_cli_render_never_silently_skips_writeback(tmp_path: Path, config_text: str) -> None:
    """FINDINGS #8.1：不带 --write-excel 时摘要必须说明"没有写回"。"""
    config_path = write_config(tmp_path, config_text)
    excel_path = tmp_path / "template.xlsx"
    runner.invoke(app, ["init", "--config", str(config_path), "--output", str(excel_path)])

    result = runner.invoke(
        app,
        ["render", "--config", str(config_path), "--excel", str(excel_path), "--no-show"],
    )
    assert result.exit_code == 0, output_of(result)
    text = output_of(result)
    assert "否（需要 --write-excel）" in text
    assert "只预览" in text

    result = runner.invoke(
        app,
        [
            "render",
            "--config",
            str(config_path),
            "--excel",
            str(excel_path),
            "--write-excel",
            "--no-show",
        ],
    )
    assert result.exit_code == 0, output_of(result)
    assert "写回 Excel" in output_of(result)


def test_cli_success_marker_is_ascii(tmp_path: Path, config_text: str) -> None:
    """FINDINGS #1：收尾打印不许出现 GBK 编不出来的字符。"""
    config_path = write_config(tmp_path, config_text)
    excel_path = tmp_path / "template.xlsx"
    runner.invoke(app, ["init", "--config", str(config_path), "--output", str(excel_path)])

    result = runner.invoke(
        app,
        ["render", "--config", str(config_path), "--excel", str(excel_path), "--no-show"],
    )
    text = output_of(result)
    assert result.exit_code == 0, text
    assert "OK" in text
    assert "\u2713" not in text and "\u2717" not in text and "✓" not in text


def test_cli_survives_gbk_stdout(tmp_path: Path, config_text: str, monkeypatch) -> None:
    """FINDINGS #1 的核心回归：把 stdout 换成只支持 GBK 的流，退出码仍必须是 0。

    直接调用命令函数（绕开 CliRunner 的捕获），这样 rich 真正写到我们换上去的流上。
    """
    from excel_codegen import cli

    config_path = write_config(tmp_path, config_text)
    raw = io.BytesIO()
    gbk = io.TextIOWrapper(raw, encoding="gbk", errors="strict", newline="")
    monkeypatch.setattr(sys, "stdout", gbk)
    try:
        cli.validate_command(config=config_path, excel=None)  # 不应抛 UnicodeEncodeError
        gbk.flush()
        assert b"OK" in raw.getvalue()
    finally:
        monkeypatch.undo()


def test_cli_check_reports_stale_and_fresh(tmp_path: Path, config_text: str) -> None:
    """FINDINGS #8.3：check 能判定"过期"，且刷新后恢复为一致。"""
    config_path = write_config(tmp_path, config_text)
    excel_path = tmp_path / "template.xlsx"
    runner.invoke(app, ["init", "--config", str(config_path), "--output", str(excel_path)])
    runner.invoke(
        app,
        [
            "render",
            "--config",
            str(config_path),
            "--excel",
            str(excel_path),
            "--write-excel",
            "--no-show",
        ],
    )

    result = runner.invoke(app, ["check", "--config", str(config_path), "--excel", str(excel_path)])
    assert result.exit_code == 0, output_of(result)
    assert "一致" in output_of(result)

    # 改参数不重跑 -> check 必须失败，并指出第几行不同
    workbook = load_workbook(excel_path)
    try:
        workbook["Global Parameter"]["B2"] = 9600
        workbook.save(excel_path)
    finally:
        workbook.close()

    result = runner.invoke(app, ["check", "--config", str(config_path), "--excel", str(excel_path)])
    assert result.exit_code == 1
    text = output_of(result)
    assert "已过期" in text
    assert "参数改过了" in text
    assert "第 2 行不同" in text  # UART_Init(…) 是渲染结果的第 2 行

    # 刷新之后又一致
    runner.invoke(
        app,
        [
            "render",
            "--config",
            str(config_path),
            "--excel",
            str(excel_path),
            "--write-excel",
            "--no-show",
        ],
    )
    result = runner.invoke(app, ["check", "--config", str(config_path), "--excel", str(excel_path)])
    assert result.exit_code == 0, output_of(result)


def test_cli_check_detects_missing_render(tmp_path: Path, config_text: str) -> None:
    """从没渲染过的工作簿：check 要报"没有记录"，并以 1 退出。"""
    config_path = write_config(tmp_path, config_text)
    excel_path = tmp_path / "template.xlsx"
    runner.invoke(app, ["init", "--config", str(config_path), "--output", str(excel_path)])

    result = runner.invoke(app, ["check", "--config", str(config_path), "--excel", str(excel_path)])
    assert result.exit_code == 1
    assert "没有渲染记录" in output_of(result)


def test_cli_init_accepts_case_names(tmp_path: Path, config_text: str) -> None:
    """FINDINGS #7：--cases 既接受数量，也接受逗号分隔的工况名。"""
    config_path = write_config(tmp_path, config_text)
    excel_path = tmp_path / "named.xlsx"

    result = runner.invoke(
        app,
        [
            "init",
            "--config",
            str(config_path),
            "--output",
            str(excel_path),
            "--cases",
            "EXT-T20.559,INT-T15",
        ],
    )
    assert result.exit_code == 0, output_of(result)

    workbook = load_workbook(excel_path)
    try:
        sheet = workbook["Local Parameter"]
        assert sheet.cell(row=1, column=5).value == "EXT-T20.559"
        assert sheet.cell(row=1, column=6).value == "INT-T15"
    finally:
        workbook.close()

    # 非法值仍然要报错
    result = runner.invoke(
        app,
        ["init", "--config", str(config_path), "--output", str(excel_path), "--cases", "0"],
    )
    assert result.exit_code == 1
    assert "至少为 1" in output_of(result)


def test_cli_init_no_howto(tmp_path: Path, config_text: str) -> None:
    config_path = write_config(tmp_path, config_text)
    excel_path = tmp_path / "no_howto.xlsx"
    result = runner.invoke(
        app,
        [
            "init",
            "--config",
            str(config_path),
            "--output",
            str(excel_path),
            "--no-howto",
            "--no-template-sheet",
        ],
    )
    assert result.exit_code == 0, output_of(result)
    workbook = load_workbook(excel_path)
    try:
        assert "HOWTO" not in workbook.sheetnames
        assert "Template" not in workbook.sheetnames
    finally:
        workbook.close()


UNUSED_YAML = """\
version: 1

excel:
  output: "unused.xlsx"
  template_sheet: null
  howto_sheet: null
  sheets:
    global: "Global Parameter"
    local: "Local Parameter"
    outputs: ["Output"]

variables:
  global:
    - name: used
      default: 1
    - name: never_used
      default: 2
  local:
    - name: also_unused
      default: 3

templates:
  - name: only_used
    code: |
      value = {{ used }}
"""


FORMULA_CLI_YAML = """\
version: 1

excel:
  output: "formula_cli.xlsx"
  template_sheet: "Template"
  howto_sheet: "HOWTO"
  sheets:
    global: "Global Parameter"
    local: "Local Parameter"
    outputs: ["Output"]

variables:
  global:
    - name: baud
      default: 115200
      type: int
  local:
    - name: port
      default: "A"
      prefix: "GPIO"

templates:
  - name: uart_init
    output_sheet: "Output"
    start_cell: "B2"
    direction: "horizontal"
    engine: "excel"
    code: |
      // {{ case_name }}
      UART_Init({{ baud }}, {{ port }});
"""


def test_cli_formula_mode_check_semantics(tmp_path: Path) -> None:
    """公式模式：改参数不用重跑（check 仍 0），改模板才需要重跑（check 1）。"""
    config_path = tmp_path / "formula_cli.yaml"
    config_path.write_text(FORMULA_CLI_YAML, encoding="utf-8")
    excel_path = tmp_path / "formula_cli.xlsx"

    runner.invoke(app, ["init", "--config", str(config_path), "--output", str(excel_path)])
    runner.invoke(
        app,
        ["render", "--config", str(config_path), "--excel", str(excel_path), "--write-excel", "--no-show"],
    )

    result = runner.invoke(app, ["check", "--config", str(config_path), "--excel", str(excel_path)])
    assert result.exit_code == 0, output_of(result)
    assert "公式: uart_init" in output_of(result)
    assert "改参数不需要重跑" in output_of(result)

    # ① 只改参数：公式不变，Excel 会自己重算 —— 不算过期
    workbook = load_workbook(excel_path)
    try:
        workbook["Global Parameter"]["B2"] = 9600
        workbook["Local Parameter"]["E2"] = "C"
        workbook.save(excel_path)
    finally:
        workbook.close()

    result = runner.invoke(app, ["check", "--config", str(config_path), "--excel", str(excel_path)])
    assert result.exit_code == 0, output_of(result)
    assert "公式模板会自动重算" in output_of(result)

    # ② 改模板：公式必须重写 —— 算过期，并说明比的是公式
    text = config_path.read_text(encoding="utf-8").replace(
        "      UART_Init({{ baud }}, {{ port }});",
        "      UART_Init2({{ baud }}, {{ port }});",
    )
    config_path.write_text(text, encoding="utf-8")
    result = runner.invoke(app, ["check", "--config", str(config_path), "--excel", str(excel_path)])
    assert result.exit_code == 1
    assert "公式模式：比的是公式" in output_of(result)

    # 重跑一次写回 → 恢复一致
    runner.invoke(
        app,
        ["render", "--config", str(config_path), "--excel", str(excel_path), "--write-excel", "--no-show"],
    )
    result = runner.invoke(app, ["check", "--config", str(config_path), "--excel", str(excel_path)])
    assert result.exit_code == 0, output_of(result)


def test_cli_formula_mode_validate_and_howto(tmp_path: Path) -> None:
    config_path = tmp_path / "formula_cli.yaml"
    config_path.write_text(FORMULA_CLI_YAML, encoding="utf-8")
    excel_path = tmp_path / "formula_cli.xlsx"

    result = runner.invoke(app, ["validate", "--config", str(config_path)])
    assert result.exit_code == 0, output_of(result)
    assert "excel·公式" in output_of(result)
    assert "公式模式（engine: excel）的模板" in output_of(result)

    runner.invoke(app, ["init", "--config", str(config_path), "--output", str(excel_path)])
    workbook = load_workbook(excel_path)
    try:
        howto = "\n".join(
            str(cell.value) for (cell,) in workbook["HOWTO"].iter_rows(min_col=1, max_col=1) if cell.value
        )
        assert "engine: excel" in howto
        assert "公式·自动重算" in howto
    finally:
        workbook.close()


def test_cli_formula_mode_rejects_for_loop(tmp_path: Path) -> None:
    """公式模式下 {% for %} 要在 validate 阶段就报出具体行（一行 = 一个单元格，循环没法表达）。"""
    config_path = tmp_path / "bad_formula.yaml"
    config_path.write_text(
        FORMULA_CLI_YAML.replace(
            "      UART_Init({{ baud }}, {{ port }});",
            "      {% for x in [1] %}UART_Init({{ baud }}, {{ port }});{% endfor %}",
        ),
        encoding="utf-8",
    )
    result = runner.invoke(app, ["validate", "--config", str(config_path)])
    assert result.exit_code == 1
    text = output_of(result)
    assert "不支持" in text
    assert "行内容" in text
    assert "engine: snapshot" in text


def test_cli_formula_mode_if_requires_attribute(tmp_path: Path) -> None:
    """行内 {% if %} 是支持的，但条件里的变量必须写 .value —— 裸变量在快照模式会直接报 TypeError。"""
    config_path = tmp_path / "bad_if.yaml"
    config_path.write_text(
        FORMULA_CLI_YAML.replace(
            "      UART_Init({{ baud }}, {{ port }});",
            "      {% if baud %}UART_Init({{ baud }}, {{ port }});{% endif %}",
        ),
        encoding="utf-8",
    )
    result = runner.invoke(app, ["validate", "--config", str(config_path)])
    assert result.exit_code == 1
    text = output_of(result)
    assert "裸写" in text
    assert "baud.value" in text


def flat(text: str) -> str:
    """去掉所有空白：rich 会按终端宽度折行，断言时不能依赖换行位置。"""
    return re.sub(r"\s+", "", text)


def test_cli_check_verifies_formula_values(tmp_path: Path) -> None:
    """公式模式：check 会把 Output 表里的公式算一遍与 Python 渲染比对。

    这里分两段验证：
    ① 结构破坏（插 Case 列）—— 公式文本/表头/行数会报出不一致，退出码 1；
    ② 值校验这条链路本身 —— 把求值结果替换成错的，必须被抓出来（否则默认开就没意义）。
    """
    from excel_codegen import cli as cli_module

    config_path = tmp_path / "formula_cli.yaml"
    config_path.write_text(FORMULA_CLI_YAML, encoding="utf-8")
    excel_path = tmp_path / "formula_cli.xlsx"

    runner.invoke(app, ["init", "--config", str(config_path), "--output", str(excel_path), "--cases", "2"])
    runner.invoke(
        app,
        ["render", "--config", str(config_path), "--excel", str(excel_path), "--write-excel", "--no-show"],
    )
    result = runner.invoke(app, ["check", "--config", str(config_path), "--excel", str(excel_path)])
    assert result.exit_code == 0, output_of(result)
    assert "值校验已开启" in output_of(result)

    # ① 在 Local 表 E 之后插一列：公式里的列标不会跟着变
    workbook = load_workbook(excel_path)
    try:
        sheet = workbook["Local Parameter"]
        sheet.insert_cols(6)
        sheet.cell(row=1, column=6, value="Case1b")
        workbook.save(excel_path)
    finally:
        workbook.close()

    result = runner.invoke(app, ["check", "--config", str(config_path), "--excel", str(excel_path)])
    assert result.exit_code == 1
    assert "--write-excel" in flat(output_of(result))

    # ② 值校验：把求值结果换成错的，只有它能把这种情况指出来（换一本干净的工作簿）
    clean = tmp_path / "formula_clean.xlsx"
    runner.invoke(app, ["init", "--config", str(config_path), "--output", str(clean), "--cases", "2"])
    runner.invoke(
        app,
        ["render", "--config", str(config_path), "--excel", str(clean), "--write-excel", "--no-show"],
    )
    result = runner.invoke(app, ["check", "--config", str(config_path), "--excel", str(clean)])
    assert result.exit_code == 0, output_of(result)

    monkeypatch = pytest.MonkeyPatch()
    try:
        real = cli_module.evaluate_template_values

        def broken(workbook, project, template, case_names):
            values = real(workbook, project, template, case_names)
            first = case_names[0]
            values[first] = [*values[first][:-1], "// 被改坏的一行"]
            return values

        monkeypatch.setattr(cli_module, "evaluate_template_values", broken)

        # 关掉值校验就只剩公式文本比对：这种情况它看不见
        result = runner.invoke(app, ["check", "--config", str(config_path), "--excel", str(clean), "--no-values"])
        assert result.exit_code == 0, output_of(result)

        # 打开值校验就被抓出来
        result = runner.invoke(app, ["check", "--config", str(config_path), "--excel", str(clean)])
        assert result.exit_code == 1
        text = flat(output_of(result))
        assert "公式算出来的文本" in text
        assert "插/删过Case列" in text
    finally:
        monkeypatch.undo()


def test_cli_warns_about_blank_case_column(tmp_path: Path) -> None:
    """全空的 Case 列会静默落进 case_filter 的规则集 —— 必须提醒。"""
    config_path = tmp_path / "filtered.yaml"
    config_path.write_text(
        FORMULA_CLI_YAML.replace(
            '    engine: "excel"',
            '    engine: "excel"\n    case_filter: "port != \'Z\'"',
        ),
        encoding="utf-8",
    )
    excel_path = tmp_path / "filtered.xlsx"
    runner.invoke(app, ["init", "--config", str(config_path), "--output", str(excel_path), "--cases", "2"])

    # 把第二列清空（模拟"新插了一个空列"）
    workbook = load_workbook(excel_path)
    try:
        sheet = workbook["Local Parameter"]
        for row in range(2, sheet.max_row + 1):
            sheet.cell(row=row, column=6).value = None
        workbook.save(excel_path)
    finally:
        workbook.close()

    result = runner.invoke(
        app,
        ["render", "--config", str(config_path), "--excel", str(excel_path), "--no-show"],
    )
    assert result.exit_code == 0, output_of(result)
    text = flat(output_of(result))
    assert "在Local表里是空的" in text
    assert "case_filter" in text
    assert "归属变量" in text
    assert "Case2" in text


def test_cli_warns_about_long_formula(tmp_path: Path) -> None:
    """一行塞太多占位符时给出公式长度警告（可编译，但人读不懂/接近上限）。"""
    config_path = tmp_path / "long.yaml"
    config_path.write_text(
        FORMULA_CLI_YAML.replace(
            "      UART_Init({{ baud }}, {{ port }});",
            "      " + "{{ baud }}" * 12 + ";",
        ),
        encoding="utf-8",
    )
    result = runner.invoke(app, ["validate", "--config", str(config_path)])
    assert result.exit_code == 0, output_of(result)
    text = output_of(result)
    assert "最长公式" in text
    assert "拆行" in text


def test_cli_validate_warns_about_unused_variable(tmp_path: Path) -> None:
    """报告里"非缺陷的观察 1"：定义了却没人用的变量要提示。"""
    config_path = tmp_path / "unused.yaml"
    config_path.write_text(UNUSED_YAML, encoding="utf-8")

    result = runner.invoke(app, ["validate", "--config", str(config_path)])
    assert result.exit_code == 0, output_of(result)
    text = output_of(result)
    assert "没有被任何模板引用" in text
    assert "never_used" in text and "also_unused" in text
    assert "used" in text  # 表格里的变量清单仍然列出它
