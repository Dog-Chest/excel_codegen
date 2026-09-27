"""一键刷新脚本：``init`` 在工作簿旁边生成 ``*_render.bat`` / ``*_render.sh``。

目标用户是工程师而不是终端爱好者 —— HOWTO 表里写了命令，但还得自己开终端敲。
这组测试除了检查生成物，还会**真的执行一遍** ``.sh``（用一个桩程序冒充
``excel-codegen``），确认 cd、参数、退出码都对。
"""

from __future__ import annotations

import os
import stat
import subprocess
import sys
from pathlib import Path

import pytest

from excel_codegen import create_template, load_config

#: 这两个用例**真的执行**生成的 ``.sh``（用桩程序冒充 excel-codegen）。
#: Windows 上跑不了：PATH 是 Unix 格式、桩程序是无扩展名的 shell 脚本。
#: ``.bat`` 那一路改为断言内容（见 test_bat_uses_windows_separators_and_crlf）。
posix_only = pytest.mark.skipif(sys.platform == "win32", reason="执行 .sh 需要 POSIX shell")

CONFIG = """\
version: 1

excel:
  output: "template.xlsx"
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
  local:
    - name: port
      default: "A"

templates:
  - name: demo
    output_sheet: "Output"
    code: |
      // {{ port }} @ {{ baud }}
"""


@pytest.fixture()
def config_path(tmp_path: Path) -> Path:
    path = tmp_path / "demo.yaml"
    path.write_text(CONFIG, encoding="utf-8")
    return path


def test_generates_both_scripts_next_to_workbook(config_path: Path, tmp_path: Path) -> None:
    project = load_config(config_path)
    target = tmp_path / "template.xlsx"
    create_template(project, target, cases=1, overwrite=True)

    bat = tmp_path / "template_render.bat"
    sh = tmp_path / "template_render.sh"
    assert bat.exists() and sh.exists()
    assert "render -c" in bat.read_text(encoding="utf-8")
    assert "render -c" in sh.read_text(encoding="utf-8")


@pytest.mark.skipif(sys.platform == "win32", reason="Windows 没有可执行位")
def test_sh_is_executable(config_path: Path, tmp_path: Path) -> None:
    create_template(load_config(config_path), tmp_path / "template.xlsx", cases=1, overwrite=True)
    mode = (tmp_path / "template_render.sh").stat().st_mode
    assert mode & stat.S_IXUSR, "Linux / macOS 上要能直接 ./ 跑"


def test_bat_uses_windows_separators_and_crlf(config_path: Path, tmp_path: Path) -> None:
    """两个平台都生成；``.bat`` 用反斜杠与 CRLF，``.sh`` 用正斜杠。"""
    create_template(load_config(config_path), tmp_path / "template.xlsx", cases=1, overwrite=True)
    raw = (tmp_path / "template_render.bat").read_bytes()
    assert b"\r\n" in raw
    assert b'cd /d "%~dp0"' in raw
    # 配置路径相对脚本目录，且是 Windows 分隔符
    assert b'"demo.yaml"' in raw

    sh_text = (tmp_path / "template_render.sh").read_text(encoding="utf-8")
    assert '"demo.yaml"' in sh_text
    assert "\\" not in sh_text.split("render -c")[1].split("\n")[0]


def test_scripts_can_be_switched_off(config_path: Path, tmp_path: Path) -> None:
    create_template(
        load_config(config_path),
        tmp_path / "template.xlsx",
        cases=1,
        overwrite=True,
        include_scripts=False,
    )
    assert not list(tmp_path.glob("*_render.*"))


def test_config_path_is_relative_to_script_dir(config_path: Path, tmp_path: Path) -> None:
    """脚本放在工作簿旁边；配置在同目录时写文件名，在子目录时写相对路径。"""
    sub = tmp_path / "workbooks"
    create_template(load_config(config_path), sub / "template.xlsx", cases=1, overwrite=True)
    text = (sub / "template_render.sh").read_text(encoding="utf-8")
    assert '"../demo.yaml"' in text


def test_bat_prefers_uv_with_fallback(config_path: Path, tmp_path: Path) -> None:
    """装了 uv 用 uv run；没装则退回 PATH 里的 excel-codegen。"""
    create_template(load_config(config_path), tmp_path / "template.xlsx", cases=1, overwrite=True)
    bat = (tmp_path / "template_render.bat").read_text(encoding="utf-8")
    assert "where uv" in bat
    assert "uv run excel-codegen render" in bat
    assert "\n  excel-codegen render" in bat.replace("\r\n", "\n")


@posix_only
def test_generated_sh_actually_runs(config_path: Path, tmp_path: Path) -> None:
    """真跑一遍：桩程序把收到的参数写下来，脚本应当转发对、并把退出码带出去。"""
    create_template(load_config(config_path), tmp_path / "template.xlsx", cases=1, overwrite=True)

    stub_dir = tmp_path / "bin"
    stub_dir.mkdir()
    recorded = tmp_path / "args.txt"
    stub = stub_dir / "excel-codegen"
    stub.write_text(f'#!/bin/sh\nprintf "%s\\n" "$@" > "{recorded}"\n', encoding="utf-8")
    stub.chmod(0o755)

    env = dict(os.environ)
    env["PATH"] = f"{stub_dir}:/usr/bin:/bin"  # 屏蔽掉 uv，强制走 fallback 分支
    proc = subprocess.run(
        ["bash", str(tmp_path / "template_render.sh")],
        cwd=tmp_path,
        env=env,
        capture_output=True,
        text=True,
        check=False,
    )
    assert proc.returncode == 0, proc.stderr
    assert recorded.read_text(encoding="utf-8").splitlines() == [
        "render",
        "-c",
        "demo.yaml",
        "-x",
        "template.xlsx",
        "--write-excel",
    ]
    assert "[完成]" in proc.stdout


@posix_only
def test_generated_sh_propagates_failure(tmp_path: Path) -> None:
    """工具失败时脚本要以非 0 退出（CI 或批处理里能看出来）。"""
    path = tmp_path / "demo.yaml"
    path.write_text(CONFIG, encoding="utf-8")
    create_template(load_config(path), tmp_path / "template.xlsx", cases=1, overwrite=True)

    stub_dir = tmp_path / "bin"
    stub_dir.mkdir()
    stub = stub_dir / "excel-codegen"
    stub.write_text("#!/bin/sh\nexit 3\n", encoding="utf-8")
    stub.chmod(0o755)

    env = dict(os.environ)
    env["PATH"] = f"{stub_dir}:/usr/bin:/bin"
    proc = subprocess.run(
        ["bash", str(tmp_path / "template_render.sh")],
        env=env,
        capture_output=True,
        text=True,
        check=False,
    )
    assert proc.returncode == 3
    assert "[失败]" in proc.stdout
