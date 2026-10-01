"""一键刷新脚本：``init`` 在工作簿旁边生成 ``*_render.bat`` / ``*_render.sh``。

目标用户是工程师而不是终端爱好者 —— HOWTO 表里写了命令，但还得自己开终端敲。
这组测试除了检查生成物，还会**真的执行一遍** ``.sh``（用一个桩程序冒充
``spreadsheet-codegen``），确认 cd、参数、退出码都对。
"""

from __future__ import annotations

import os
import stat
import subprocess
import sys
from pathlib import Path

import pytest

from spreadsheet_codegen import create_template, load_config

#: 这四条用例**真的执行**生成的 ``.sh``（用桩程序冒充 spreadsheet-codegen）。
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


def test_bat_declares_utf8_code_page(config_path: Path, tmp_path: Path) -> None:
    """``.bat`` 里有中文，必须先切到 65001 —— 否则 GBK 控制台下是乱码，
    甚至因为字节被当成命令而报 ``'xxx' is not recognized``（真踩过）。"""
    create_template(load_config(config_path), tmp_path / "template.xlsx", cases=1, overwrite=True)
    bat = (tmp_path / "template_render.bat").read_text(encoding="utf-8")
    lines = [line.strip() for line in bat.replace("\r\n", "\n").splitlines()]
    assert "chcp 65001 >nul" in lines, "缺 chcp 65001，中文提示在中文 Windows 上会乱码"
    # chcp 必须在使用任何中文之前
    assert lines.index("chcp 65001 >nul") < lines.index('cd /d "%~dp0"')


def test_bat_fallback_chain_covers_module_invocation(config_path: Path, tmp_path: Path) -> None:
    """回退链要覆盖"用源码 / ``python -m`` 跑"的机器 —— 只有 ``uv`` + PATH 上那条命令
    的脚本在那种机器上**两条路都不通**，而旧提示还怪"依赖没装"（指错了原因）。

    包名是 ``spreadsheet_codegen``（下划线），命令才是 ``spreadsheet-codegen``（连字符）：
    ``python -m`` 那条路最稳，必须有。
    """
    create_template(load_config(config_path), tmp_path / "template.xlsx", cases=1, overwrite=True)
    bat = (tmp_path / "template_render.bat").read_text(encoding="utf-8").replace("\r\n", "\n")
    sh = (tmp_path / "template_render.sh").read_text(encoding="utf-8")

    for text in (bat, sh):
        assert "uv run spreadsheet-codegen render" in text
        assert "uv run python -m spreadsheet_codegen render" in text
        assert "spreadsheet-codegen render" in text
    assert "python -m spreadsheet_codegen render" in bat, ".bat 要能走本机 python 的模块调用"
    assert "python3 -m spreadsheet_codegen render" in sh, ".sh 要能走本机 python3 的模块调用"

    # 失败提示必须说实话：列出试过的方式 + 两条真正的出路，而不是只说"依赖没装"
    assert "SPREADSHEET_CODEGEN_PY" in bat and "SPREADSHEET_CODEGEN_PY" in sh
    assert "pip install spreadsheet-codegen" in sh
    assert "--recreate" not in bat and "--recreate" not in sh
    assert "uv sync" not in bat and "uv sync" not in sh


def test_bat_tries_every_route_before_giving_up(config_path: Path, tmp_path: Path) -> None:
    """每条路都要真的被"试过"（``&&`` 短路 + 成功标记），而不是第一条失败就退出。"""
    create_template(load_config(config_path), tmp_path / "template.xlsx", cases=1, overwrite=True)
    bat = (tmp_path / "template_render.bat").read_text(encoding="utf-8").replace("\r\n", "\n")
    assert 'set "SPREADSHEET_CODEGEN_OK=1"' in bat
    assert bat.count("if defined SPREADSHEET_CODEGEN_OK goto :done") >= 5
    assert bat.count("SPREADSHEET_CODEGEN_OK=1") == bat.count('&& set "SPREADSHEET_CODEGEN_OK=1"')
    # pause 留着（双击时看得见结论），但它**必须在退出码之前** —— 见下一条测试
    assert "pause" in bat


def test_bat_exits_non_zero_when_every_route_failed(config_path: Path, tmp_path: Path) -> None:
    """全部失败时 ``.bat`` 要以非 0 退出，否则 CI / 批处理看不出失败。

    回归：脚本结尾是 ``echo`` + ``pause``，两条内建命令都会把 ERRORLEVEL 归零 ——
    于是"五条路全失败"照样以 0 结束。``pause`` 之后必须显式 ``exit /b``。
    """
    create_template(load_config(config_path), tmp_path / "template.xlsx", cases=1, overwrite=True)
    bat = (tmp_path / "template_render.bat").read_text(encoding="utf-8").replace("\r\n", "\n")
    assert "if defined SPREADSHEET_CODEGEN_OK (exit /b 0) else (exit /b 1)" in bat
    # 退出码是最后一条命令：pause 在它前面（窗口留得住，退出码又不会被归零）
    body = bat.rstrip().splitlines()
    assert body[-1].strip() == "if defined SPREADSHEET_CODEGEN_OK (exit /b 0) else (exit /b 1)"
    assert "pause" in body[-2], "pause 要紧挨在退出码前：先让人看见结论，再带出退出码"


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
    """装了 uv 先试 uv；uv 的两条路都不通再退回本机 python / PATH 上的命令。"""
    create_template(load_config(config_path), tmp_path / "template.xlsx", cases=1, overwrite=True)
    bat = (tmp_path / "template_render.bat").read_text(encoding="utf-8").replace("\r\n", "\n")
    assert "where uv" in bat
    # uv 的两条路都在 where uv 后面的 if 块里，且命令那条排在模块那条前面
    after_uv = bat.split("where uv", 1)[1]
    uv_block = after_uv.split("if defined SPREADSHEET_CODEGEN_PY", 1)[0]
    assert "uv run spreadsheet-codegen render" in uv_block
    assert "uv run python -m spreadsheet_codegen render" in uv_block
    assert uv_block.index("uv run spreadsheet-codegen render") < uv_block.index(
        "uv run python -m spreadsheet_codegen render"
    )


@posix_only
def test_generated_sh_actually_runs(config_path: Path, tmp_path: Path) -> None:
    """真跑一遍：桩程序把收到的参数写下来，脚本应当转发对、并把退出码带出去。"""
    create_template(load_config(config_path), tmp_path / "template.xlsx", cases=1, overwrite=True)

    stub_dir = tmp_path / "bin"
    stub_dir.mkdir()
    recorded = tmp_path / "args.txt"
    stub = stub_dir / "spreadsheet-codegen"
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
def test_generated_sh_falls_back_to_module_invocation(tmp_path: Path) -> None:
    """真跑一遍：装着的命令用不了时，脚本要退到 ``uv run python -m spreadsheet_codegen``。

    这正是反馈里的场景：一台"用源码 / ``python -m`` 跑"的机器上，旧脚本的两条路
    （``uv run spreadsheet-codegen`` 与 PATH 上那条命令）**都不通**，
    而失败提示还怪"依赖没装"。桩 ``uv`` 只认模块调用那条路。
    """
    config_path = tmp_path / "demo.yaml"
    config_path.write_text(CONFIG, encoding="utf-8")
    create_template(load_config(config_path), tmp_path / "template.xlsx", cases=1, overwrite=True)

    stub_dir = tmp_path / "bin"
    stub_dir.mkdir()
    recorded = tmp_path / "args.txt"
    # 桩 uv：只认模块调用那条路（$1=run、$2=python、$3=-m）；`uv run spreadsheet-codegen …` 失败。
    # 这样就能证明脚本真的**退到了第二条路**，而不是靠 PATH 上的命令蒙对。
    stub = stub_dir / "uv"
    stub.write_text(
        "#!/bin/sh\n"
        'if [ "$2" = "python" ] && [ "$3" = "-m" ]; then\n'
        f'  printf "%s\\n" "$@" > "{recorded}"\n'
        "  exit 0\n"
        "fi\n"
        "exit 1\n",
        encoding="utf-8",
    )
    stub.chmod(0o755)

    env = dict(os.environ)
    env["PATH"] = f"{stub_dir}:/usr/bin:/bin"
    proc = subprocess.run(
        ["bash", str(tmp_path / "template_render.sh")],
        cwd=tmp_path,
        env=env,
        capture_output=True,
        text=True,
        check=False,
    )
    assert proc.returncode == 0, proc.stderr + proc.stdout
    assert recorded.read_text(encoding="utf-8").splitlines() == [
        "run",
        "python",
        "-m",
        "spreadsheet_codegen",
        "render",
        "-c",
        "demo.yaml",
        "-x",
        "template.xlsx",
        "--write-excel",
    ]
    assert "[完成]" in proc.stdout


@posix_only
def test_generated_sh_reports_every_route_it_tried(tmp_path: Path) -> None:
    """全都失败时提示要**说实话**：列出试过的方式与两条真正的出路，而不是怪"依赖没装"。"""
    config_path = tmp_path / "demo.yaml"
    config_path.write_text(CONFIG, encoding="utf-8")
    create_template(load_config(config_path), tmp_path / "template.xlsx", cases=1, overwrite=True)

    stub_dir = tmp_path / "bin"
    stub_dir.mkdir()
    for name in ("uv", "spreadsheet-codegen", "python3"):
        failing = stub_dir / name
        failing.write_text("#!/bin/sh\nexit 1\n", encoding="utf-8")
        failing.chmod(0o755)

    env = dict(os.environ)
    env["PATH"] = f"{stub_dir}:/usr/bin:/bin"
    proc = subprocess.run(
        ["bash", str(tmp_path / "template_render.sh")],
        cwd=tmp_path,
        env=env,
        capture_output=True,
        text=True,
        check=False,
    )
    assert proc.returncode != 0
    assert "pip install spreadsheet-codegen" in proc.stdout
    assert "SPREADSHEET_CODEGEN_PY" in proc.stdout
    # 不允许再把"依赖没装 / uv sync"当唯一解释（那会指向一个不存在的原因）
    assert "uv sync" not in proc.stdout
    assert "setup.sh" not in proc.stdout


@posix_only
def test_generated_sh_propagates_failure(tmp_path: Path) -> None:
    """工具失败时脚本要以非 0 退出（CI 或批处理里能看出来）。"""
    path = tmp_path / "demo.yaml"
    path.write_text(CONFIG, encoding="utf-8")
    create_template(load_config(path), tmp_path / "template.xlsx", cases=1, overwrite=True)

    stub_dir = tmp_path / "bin"
    stub_dir.mkdir()
    stub = stub_dir / "spreadsheet-codegen"
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
