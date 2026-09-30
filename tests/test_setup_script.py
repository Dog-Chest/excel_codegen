"""``setup.sh`` 复用旧环境之前的体检：坏环境必须在**装依赖之前**被拦住。

背景（真踩过，见 CHANGELOG 0.8.1）：开发机的 venv 是按 Python 3.12 建的，后来系统的
``/usr/bin/python3.12`` 被删、``python3`` 指向 3.14。venv 里的 ``bin/python3`` 是**通用**
链接（``-> /usr/bin/python3``）而不是钉死版本，于是解释器悄悄漂到 3.14，而依赖还躺在
``lib/python3.12/site-packages`` 里。结果很反直觉：

* ``bin/python -c ""`` 照样返回 0 —— 老版体检只测这一句，于是判定"可以复用"；
* 可 ``import jinja2`` 与 ``-m pip`` 全部失败，脚本最终停在一句 ``No module named pip``。

这组用例把那三种坏法各钉一条，并确认正常环境**不会**被误判。

另外钉一条**可移植性**：shell 里 ``$VAR`` 紧挨着中文必须写成 ``${VAR}``。macOS 自带
bash 3.2，解析器会把多字节字符的首字节吞进变量名 —— ``$REPO_FS（`` 变成
``REPO_FS<0xEF>``，在 ``set -u`` 下第一屏就 ``unbound variable`` 退出。Linux 的
bash ≥ 4 解析正常，所以这个坑只在 macOS 上炸（见 CHANGELOG 0.9.0）。
"""

from __future__ import annotations

import os
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

import pytest

from spreadsheet_codegen import load_config
from spreadsheet_codegen.example_pack import examples_root
from spreadsheet_codegen.excel_io import write_run_scripts

ROOT = Path(__file__).resolve().parents[1]
SETUP = ROOT / "setup.sh"

#: ``$VAR`` 紧挨着一个非 ASCII 字节 —— 老 bash 会把那个字节当成变量名的一部分
BARE_EXPANSION_BEFORE_MULTIBYTE = re.compile(rb"\$([A-Za-z_][A-Za-z0-9_]*)(?=[\x80-\xff])")

posix_only = pytest.mark.skipif(sys.platform == "win32", reason="setup.sh 需要 POSIX shell")


@pytest.fixture(scope="module")
def base_venv(tmp_path_factory: pytest.TempPathFactory) -> Path:
    """一个真的、可复制的空 venv（``--without-pip`` 建得最快，也正好是"没有 pip"的样本）。"""
    path = tmp_path_factory.mktemp("base") / "env"
    subprocess.run(
        [sys.executable, "-m", "venv", "--without-pip", str(path)],
        check=True,
        capture_output=True,
    )
    return path


@pytest.fixture()
def env_dir(base_venv: Path, tmp_path: Path) -> Path:
    """每个用例拿一份副本，改坏了不影响别人。"""
    dest = tmp_path / "env"
    shutil.copytree(base_venv, dest, symlinks=True)
    return dest


def _site_packages(env_dir: Path) -> Path:
    return next((env_dir / "lib").glob("python*/site-packages"))


def _run_setup(env_dir: Path, tmp_path: Path) -> subprocess.CompletedProcess[str]:
    """按 ``SPREADSHEET_CODEGEN_VENV`` 指定环境跑一遍 ``setup.sh --no-test``。

    PATH 里放一个指向当前解释器的 ``python3`` 垫片，并**屏蔽 uv** —— 否则装了 uv 的机器
    会走 ``uv sync`` 那条路，根本到不了这里要测的体检分支。
    """
    shim = tmp_path / "pathbin"
    shim.mkdir()
    (shim / "python3").symlink_to(sys.executable)

    env = dict(os.environ)
    env["PATH"] = f"{shim}:/usr/bin:/bin"
    env["SPREADSHEET_CODEGEN_VENV"] = str(env_dir)
    return subprocess.run(
        ["bash", str(SETUP), "--no-test"],
        cwd=ROOT,
        env=env,
        capture_output=True,
        text=True,
        # 不赌子进程的输出编码：脚本真出问题时 bash 会把变量名连同坏字节一起打出来
        # （见文件头的 bash 3.2 可移植性一节），严格解码会先抛 UnicodeDecodeError，
        # 断言反而没机会跑。与 build.py / 探针里读子进程输出的写法一致。
        encoding="utf-8",
        errors="replace",
        check=False,
    )


def _stub_pip(env_dir: Path) -> None:
    """放一个只会退出的 pip 模块：让体检过关，后面的安装也不必联网。"""
    pkg = _site_packages(env_dir) / "pip"
    pkg.mkdir()
    (pkg / "__init__.py").write_text("", encoding="utf-8")
    (pkg / "__main__.py").write_text("import sys\nsys.exit(0)\n", encoding="utf-8")


@posix_only
def test_stale_interpreter_is_caught_before_installing(env_dir: Path, tmp_path: Path) -> None:
    """venv 按旧 Python 建、解释器却漂到新版 —— 老版体检放行的正是这一种。"""
    actual = f"{sys.version_info[0]}.{sys.version_info[1]}.{sys.version_info[2]}"
    recorded = "3.12.3" if actual != "3.12.3" else "3.12.4"
    cfg = env_dir / "pyvenv.cfg"
    cfg.write_text(
        re.sub(r"^version = .*$", f"version = {recorded}", cfg.read_text(encoding="utf-8"), flags=re.M),
        encoding="utf-8",
    )

    # 当初的 bug：只测"能不能执行"的话，这种环境是**通过**的
    assert subprocess.run([str(env_dir / "bin" / "python"), "-c", ""], check=False).returncode == 0

    proc = _run_setup(env_dir, tmp_path)
    assert proc.returncode != 0
    # 断言只锚在 ASCII 证据上：子进程输出在某些平台上可能有一个坏字节被
    # errors="replace" 换成 U+FFFD，用中文短语当断言会变成"看运气"。
    assert "--recreate" in proc.stderr, proc.stderr
    assert recorded in proc.stderr and actual in proc.stderr, "报错要把两个版本都写出来才算可操作"
    # 关键：在碰 pip 之前就停住，不再以裸的 "No module named pip" 收场
    assert "pip" not in proc.stdout.lower(), proc.stdout


@posix_only
def test_site_packages_mismatch_is_caught(env_dir: Path, tmp_path: Path) -> None:
    """版本号对得上，但包目录不在当前解释器的搜索路径上（换 Python 后的遗留形态）。"""
    _site_packages(env_dir).parent.rename(env_dir / "lib" / "python9.9")

    proc = _run_setup(env_dir, tmp_path)
    assert proc.returncode != 0
    assert "site-packages" in proc.stderr, proc.stderr
    assert "pip" not in proc.stdout.lower(), proc.stdout


@posix_only
def test_missing_pip_is_caught(env_dir: Path, tmp_path: Path) -> None:
    """没有 pip 的环境复用不了：下一句就是 ``-m pip install``。"""
    proc = _run_setup(env_dir, tmp_path)
    assert proc.returncode != 0
    assert "pip" in proc.stderr.lower(), proc.stderr
    # 必须是**体检**给出的出路，而不是把 pip 自己的报错原样漏出来
    assert "--recreate" in proc.stderr, proc.stderr


@posix_only
def test_usable_environment_is_reused(env_dir: Path, tmp_path: Path) -> None:
    """体检过关就照旧复用 —— 别把好环境也判死。"""
    _stub_pip(env_dir)

    proc = _run_setup(env_dir, tmp_path)
    assert proc.returncode == 0, proc.stderr
    # 走到安装那一步（stdout 出现 pip）才算"没被误拦"
    assert "pip" in proc.stdout.lower(), proc.stdout
    assert "ERROR" not in proc.stderr, proc.stderr


# --------------------------------------------------------------------------- #
# 可移植性：$VAR 紧挨中文会被老 bash（macOS 自带 3.2）吞进变量名
# --------------------------------------------------------------------------- #
def _shell_scripts_to_scan() -> list[Path]:
    return [SETUP, *sorted(examples_root().rglob("*.sh"))]


@pytest.mark.parametrize("path", _shell_scripts_to_scan(), ids=lambda p: p.name)
def test_bundled_shell_scripts_brace_expansions_before_multibyte(path: Path) -> None:
    """``$VAR（`` 在 bash 3.2 上会被解析成 ``VAR<首字节>: unbound variable``。"""
    hits = BARE_EXPANSION_BEFORE_MULTIBYTE.findall(path.read_bytes())
    assert not hits, f"{path.name} 里有紧邻多字节字符的裸变量展开，改成 ${{VAR}}：{[h.decode() for h in hits]}"


def test_generated_render_scripts_are_free_of_that_hazard() -> None:
    """生成的 ``*_render.sh`` 也要守同一条 —— 它同样是给人双击/直接跑的。"""
    config = load_config(examples_root() / "basic" / "example.yaml")
    with tempfile.TemporaryDirectory() as tmp:
        written = write_run_scripts(config, Path(tmp) / "template.xlsx")
        shells = [p for p in written if p.suffix == ".sh"]
        assert shells, "应当生成 .sh"
        for path in shells:
            hits = BARE_EXPANSION_BEFORE_MULTIBYTE.findall(path.read_bytes())
            assert not hits, f"{path.name}: {[h.decode() for h in hits]}"
