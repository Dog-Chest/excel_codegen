"""pytest 公共 fixture。

同时把项目根目录加入 ``sys.path``，这样即使没有 ``pip install -e .`` 也能直接运行测试。
"""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from excel_codegen.excel_io import create_template  # noqa: E402
from excel_codegen.models import ProjectConfig, load_config  # noqa: E402

CONFIG_YAML = """\
version: 1

excel:
  output: "template.xlsx"
  template_sheet: "Template"
  sheets:
    global: "Global Parameter"
    local: "Local Parameter"
    outputs: ["Output", "Output Vertical"]

variables:
  global:
    - name: baud
      description: "波特率"
      default: 115200
      prefix: ""
      suffix: ""
      type: int
    - name: mcu
      description: "MCU 型号"
      default: "STM32F103"
      prefix: ""
      suffix: ""
  local:
    - name: port
      description: "端口"
      default: "A"
      prefix: "GPIO"
      suffix: "_PORT"
    - name: mode
      description: "模式"
      default: "MODE_TX_RX"
      prefix: ""
      suffix: ""

templates:
  # 这个 fixture 演示的是"快照模式 + 过滤器"，所以显式声明 engine: snapshot
  # （默认已经是公式模式 excel，见 tests/test_engine_default.py）
  - name: uart_init
    output_sheet: "Output"
    start_cell: "B2"
    direction: "horizontal"
    engine: "snapshot"
    filename: "uart_init_{{ case_name }}.c"
    code: |
      // Case: {{ case_name }}
      UART_Init({{ baud }}, {{ port }}, {{ mode }});
      // 纯值: {{ port.value }}, 前缀: {{ port.prefix }}, 后缀: {{ port.suffix }}
      // 过滤器: {{ port.value | pvs("GPIO", "_PORT") }}
      // MCU: {{ mcu }}

  - name: uart_summary
    output_sheet: "Output Vertical"
    start_cell: "B2"
    direction: "vertical"
    engine: "snapshot"
    filename: "uart_summary_{{ case_name }}.md"
    code: |
      | {{ case_name }} | {{ baud }} | {{ port }} | {{ mode }} |
"""


@pytest.fixture()
def config_text() -> str:
    return CONFIG_YAML


@pytest.fixture()
def project(tmp_path: Path, config_text: str) -> ProjectConfig:
    """把 YAML 写到临时目录并加载为 ProjectConfig。"""
    config_path = tmp_path / "example.yaml"
    config_path.write_text(config_text, encoding="utf-8")
    return load_config(config_path)


@pytest.fixture()
def workbook_path(tmp_path: Path, project: ProjectConfig) -> Path:
    """生成一个包含 2 个 Case 的已填写 Excel 模板。"""
    return create_template(project, tmp_path / "template.xlsx", cases=2, overwrite=True)
