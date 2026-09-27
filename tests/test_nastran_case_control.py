"""生成测试：NASTRAN 工况控制语句（Case Control）。

工程背景
--------
Nastran 的工况控制段是"一行一个子工况、一列一条语句"的形态：

```
SUBCASE 1
SUBTITLE = HEAD SEA
LOAD = 101
SPC = 1
SUBCOM 4
SUBSEQ = 1, 1.2, 1.3
```

这类数据在实践里往往是从表格 / 文本 / 另一个软件里整块拿到的，所以本测试用
``excel.local_direction: vertical``（指南 §19）：一个子工况一行，可以直接粘贴。

要覆盖的三件事
--------------
1. **语句是可选的吗** —— 每一条语句留空时，输出里**不能有这一行**，也不能留下空行；
2. **SUBCASE 与 SUBCOM 两种块**都对，``SUBSEQ`` 的内容原样输出（工具不解释它）；
3. **整块拼起来**是一段合法的 case control（按 ``seq`` 排序连起来）。

另外还钉住两个曾经会出问题的地方：``{% if %}`` 的换行必须写在 if **里面**
（否则语句被省略时会留下空行），以及 ``filename`` 里用到只做排序的变量 ``seq``。
"""

from __future__ import annotations

from pathlib import Path

import pytest
from openpyxl import load_workbook

from excel_codegen.excel_io import create_template
from excel_codegen.models import ProjectConfig, load_config
from excel_codegen.renderer import export_files, render_all
from excel_codegen.utils import ExcelError

CASE_CONTROL_YAML = """\
version: 1

excel:
  output: "cc.xlsx"
  template_sheet: null
  howto_sheet: null
  local_direction: "vertical"
  sheets:
    global: "Global Parameter"
    local: "Local Parameter"
    outputs: ["Case Control"]

variables:
  global:
    - name: analysis
      default: "STATICS"
      type: string
  local:
    - name: seq
      default: "01"
      type: string
    - name: kind
      default: "SUBCASE"
      type: string
      choices: ["SUBCASE", "SUBCOM"]
    - name: subcase_id
      default: 1
      type: int
      min: 1
    - name: label
      default: ""
      type: string
    - name: subtitle
      default: ""
      type: string
    - name: method
      default: ""
      type: string
    - name: load
      default: ""
      type: string
    - name: spc
      default: ""
      type: string
    - name: subseq
      default: ""
      type: string

asserts:
  - "kind.value != 'SUBCOM' or subseq.value != ''"
  - "kind.value != 'SUBCASE' or subseq.value == ''"

templates:
  - name: case_control
    output_sheet: "Case Control"
    start_cell: "B2"
    direction: "horizontal"
    write_case_headers: true
    filename: "cc_{{ seq }}_{{ case_name }}.inc"
    code: |-
      {{ kind }} {{ subcase_id }}{% if label %}
      LABEL = {{ label }}{% endif %}{% if subtitle %}
      SUBTITLE = {{ subtitle }}{% endif %}{% if analysis %}
      ANALYSIS = {{ analysis }}{% endif %}{% if method %}
      METHOD = {{ method }}{% endif %}{% if load %}
      LOAD = {{ load }}{% endif %}{% if spc %}
      SPC = {{ spc }}{% endif %}{% if subseq %}
      SUBSEQ = {{ subseq }}{% endif %}
  - name: cc_summary
    output_sheet: "Case Control"
    start_cell: "H2"
    direction: "vertical"
    write_case_headers: true
    filename: "cc_summary_{{ seq }}_{{ case_name }}.md"
    code: |-
      {{ kind }} {{ subcase_id }}
"""

#: 样例子工况：第 3 行是"组合"块，最后一行**除编号外全空**
ROWS: dict[str, dict[str, object]] = {
    "LC1": {"seq": "01", "kind": "SUBCASE", "subcase_id": 1, "subtitle": "HEAD SEA", "load": "101", "spc": "1"},
    "LC2": {"seq": "02", "kind": "SUBCASE", "subcase_id": 2, "subtitle": "BEAM SEA", "load": "102", "spc": "1"},
    "LC3": {
        "seq": "03",
        "kind": "SUBCASE",
        "subcase_id": 3,
        "label": "QTR",
        "subtitle": "QUARTERING SEA",
        "method": "1",
        "load": "103",
        "spc": "2",
    },
    "COMB4": {
        "seq": "04",
        "kind": "SUBCOM",
        "subcase_id": 4,
        "subtitle": "COMBINED 1.0 / 1.2 / 1.3",
        "subseq": "1, 1.2, 1.3",
    },
    "LC5": {"seq": "05", "kind": "SUBCASE", "subcase_id": 5},
}

EXPECTED = {
    "LC1": ["SUBCASE 1", "SUBTITLE = HEAD SEA", "ANALYSIS = STATICS", "LOAD = 101", "SPC = 1"],
    "LC2": ["SUBCASE 2", "SUBTITLE = BEAM SEA", "ANALYSIS = STATICS", "LOAD = 102", "SPC = 1"],
    "LC3": [
        "SUBCASE 3",
        "LABEL = QTR",
        "SUBTITLE = QUARTERING SEA",
        "ANALYSIS = STATICS",
        "METHOD = 1",
        "LOAD = 103",
        "SPC = 2",
    ],
    "COMB4": ["SUBCOM 4", "SUBTITLE = COMBINED 1.0 / 1.2 / 1.3", "ANALYSIS = STATICS", "SUBSEQ = 1, 1.2, 1.3"],
    # 只填了编号：别的语句一条都不输出
    "LC5": ["SUBCASE 5", "ANALYSIS = STATICS"],
}

#: 连成一整段 case control（按 seq 顺序）应该长这样
EXPECTED_DECK = """\
SUBCASE 1
SUBTITLE = HEAD SEA
ANALYSIS = STATICS
LOAD = 101
SPC = 1
SUBCASE 2
SUBTITLE = BEAM SEA
ANALYSIS = STATICS
LOAD = 102
SPC = 1
SUBCASE 3
LABEL = QTR
SUBTITLE = QUARTERING SEA
ANALYSIS = STATICS
METHOD = 1
LOAD = 103
SPC = 2
SUBCOM 4
SUBTITLE = COMBINED 1.0 / 1.2 / 1.3
ANALYSIS = STATICS
SUBSEQ = 1, 1.2, 1.3
SUBCASE 5
ANALYSIS = STATICS"""


@pytest.fixture()
def project(tmp_path: Path) -> ProjectConfig:
    path = tmp_path / "case_control.yaml"
    path.write_text(CASE_CONTROL_YAML, encoding="utf-8")
    return load_config(path)


@pytest.fixture()
def workbook_path(tmp_path: Path, project: ProjectConfig) -> Path:
    """建竖排 Local 表并按 ROWS 填值 —— 这一步就是"整块粘贴"的自动化版本。"""
    path = create_template(project, tmp_path / "cc.xlsx", cases=list(ROWS))
    workbook = load_workbook(path)
    try:
        sheet = workbook["Local Parameter"]
        columns = {sheet.cell(row=1, column=c).value: c for c in range(2, sheet.max_column + 1)}
        rows = {sheet.cell(row=r, column=1).value: r for r in range(2, sheet.max_row + 1)}
        for case, values in ROWS.items():
            for name, value in values.items():
                sheet.cell(row=rows[case], column=columns[name], value=value)
        workbook.save(path)
    finally:
        workbook.close()
    return path


# --------------------------------------------------------------------------- #
# 逐个 Case：语句可选
# --------------------------------------------------------------------------- #
def test_each_subcase_block(workbook_path: Path, project: ProjectConfig) -> None:
    out = render_all(project, workbook_path)
    results = out.results["case_control"]
    assert [result.case_name for result in results] == list(ROWS)
    for result in results:
        assert result.lines == EXPECTED[result.case_name], result.case_name


def test_blank_statements_are_omitted_not_emptied(workbook_path: Path, project: ProjectConfig) -> None:
    """留空 = 这一行不存在。曾经的做法会在省略处留下空行。"""
    results = {r.case_name: r.lines for r in render_all(project, workbook_path).results["case_control"]}
    assert "LABEL = " not in results["LC1"]
    assert not any(line.strip() == "" for lines in results.values() for line in lines)
    # LC5 只有 SUBCASE 与全局的 ANALYSIS 两条
    assert len(results["LC5"]) == 2


def test_subcom_uses_subseq_and_keeps_it_verbatim(workbook_path: Path, project: ProjectConfig) -> None:
    """SUBSEQ 的内容原样输出（工具不解释"子工况号 / 比例因子"的写法）。"""
    lines = next(r for r in render_all(project, workbook_path).results["case_control"] if r.case_name == "COMB4").lines
    assert lines[0] == "SUBCOM 4"
    assert "SUBSEQ = 1, 1.2, 1.3" in lines
    assert not any(line.startswith("LOAD") for line in lines)
    assert not any(line.startswith("SPC") for line in lines)


def test_every_row_starts_with_kind_and_id(workbook_path: Path, project: ProjectConfig) -> None:
    for result in render_all(project, workbook_path).results["case_control"]:
        kind, _, cid = result.lines[0].partition(" ")
        assert kind in {"SUBCASE", "SUBCOM"}, result.case_name
        assert cid.isdigit(), result.case_name


# --------------------------------------------------------------------------- #
# 拼成整段 case control
# --------------------------------------------------------------------------- #
def test_exported_files_concatenate_into_one_deck(workbook_path: Path, project: ProjectConfig, tmp_path: Path) -> None:
    out = render_all(project, workbook_path)
    files = export_files(project, out.results, tmp_path / "deck")
    names = sorted(path.name for path in files if path.suffix == ".inc")
    assert names == ["cc_01_LC1.inc", "cc_02_LC2.inc", "cc_03_LC3.inc", "cc_04_COMB4.inc", "cc_05_LC5.inc"]

    deck = "".join((tmp_path / "deck" / name).read_text(encoding="utf-8") for name in names)
    assert deck.rstrip("\n") == EXPECTED_DECK


def test_filename_can_use_any_parameter(workbook_path: Path, project: ProjectConfig, tmp_path: Path) -> None:
    """``seq`` 只出现在 filename 里 —— 不算"定义了没人用"，也要真的能渲染出来。"""
    from excel_codegen.renderer import collect_variables

    template = project.templates[0]
    assert "seq" in collect_variables(template, base_dir=project.source_dir)

    out = render_all(project, workbook_path)
    files = export_files(project, out.results, tmp_path / "deck")
    assert (tmp_path / "deck" / "cc_04_COMB4.inc").exists()
    assert len(files) == 2 * len(ROWS)  # 两个模板 × 5 个 Case


# --------------------------------------------------------------------------- #
# Output 表写回
# --------------------------------------------------------------------------- #
def test_blocks_are_written_to_the_sheet(workbook_path: Path, project: ProjectConfig) -> None:
    from excel_codegen.excel_io import write_results

    write_results(workbook_path, project, render_all(project, workbook_path).results)
    workbook = load_workbook(workbook_path)
    try:
        sheet = workbook["Case Control"]
        # 横向：一个 Case 一列；表头在 B1 起
        assert [sheet.cell(row=1, column=c).value for c in range(2, 7)] == list(ROWS)
        assert [sheet.cell(row=r, column=2).value for r in range(2, 7)] == [
            "SUBCASE 1",
            "SUBTITLE = HEAD SEA",
            "ANALYSIS = STATICS",
            "LOAD = 101",
            "SPC = 1",
        ]
        # 第 5 列（LC5）只有两行，第 3 行开始必须是空的 —— 不能残留上一版的行
        assert sheet.cell(row=4, column=6).value is None
    finally:
        workbook.close()


# --------------------------------------------------------------------------- #
# 填错地方要被拦下来（asserts）
# --------------------------------------------------------------------------- #
def test_subcom_without_subseq_is_rejected(workbook_path: Path, project: ProjectConfig) -> None:
    workbook = load_workbook(workbook_path)
    try:
        sheet = workbook["Local Parameter"]
        columns = {sheet.cell(row=1, column=c).value: c for c in range(2, sheet.max_column + 1)}
        # openpyxl 的 cell(..., value=None) 不会真的清空，要直接赋值
        sheet.cell(row=5, column=columns["subseq"]).value = None  # COMB4 的 SUBSEQ 抠掉
        workbook.save(workbook_path)
    finally:
        workbook.close()
    with pytest.raises(ExcelError, match=r"(?s)COMB4.*subseq"):
        render_all(project, workbook_path)


def test_subseq_on_a_plain_subcase_is_rejected(workbook_path: Path, project: ProjectConfig) -> None:
    workbook = load_workbook(workbook_path)
    try:
        sheet = workbook["Local Parameter"]
        columns = {sheet.cell(row=1, column=c).value: c for c in range(2, sheet.max_column + 1)}
        sheet.cell(row=2, column=columns["subseq"], value="1, 2")  # LC1 是 SUBCASE，不该有 SUBSEQ
        workbook.save(workbook_path)
    finally:
        workbook.close()
    with pytest.raises(ExcelError, match="asserts"):
        render_all(project, workbook_path)


def test_unknown_kind_is_rejected(workbook_path: Path, project: ProjectConfig) -> None:
    """kind 声明了 choices：写成别的词要在读表时就被拦住。"""
    workbook = load_workbook(workbook_path)
    try:
        sheet = workbook["Local Parameter"]
        columns = {sheet.cell(row=1, column=c).value: c for c in range(2, sheet.max_column + 1)}
        sheet.cell(row=2, column=columns["kind"], value="SUBCASE 1")  # 把整条语句塞进一格里
        workbook.save(workbook_path)
    finally:
        workbook.close()
    with pytest.raises(ExcelError, match=r"choices|允许列表"):
        render_all(project, workbook_path)


# --------------------------------------------------------------------------- #
# 仓库里那个示例本身也要是好的
# --------------------------------------------------------------------------- #
def test_shipped_example_matches_the_test_fixture() -> None:
    """``examples/nastran_case_control.xlsx`` 与这里的样例是同一套规则、同一批数据。"""
    here = Path(__file__).resolve().parents[1]
    example = here / "examples" / "nastran_case_control.yaml"
    book = here / "examples" / "nastran_case_control.xlsx"
    if not book.exists():  # pragma: no cover - sdist 里可能只带了 YAML
        pytest.skip("示例工作簿不在（只带了 YAML）")

    config = load_config(example)
    assert config.excel.local_direction == "vertical"
    out = render_all(config, book)
    results = {r.case_name: r.lines for r in out.results["case_control"]}
    assert list(results) == list(ROWS)
    for case, expected in EXPECTED.items():
        assert results[case] == expected, case

    deck_path = here / "examples" / "generated_nastran" / "case_control.deck"
    if not deck_path.exists():  # pragma: no cover - sdist 里可能只带了 YAML + 工作簿
        pytest.skip("拼好的 case_control.deck 不在")
    assert deck_path.read_text(encoding="utf-8").rstrip("\n") == EXPECTED_DECK
