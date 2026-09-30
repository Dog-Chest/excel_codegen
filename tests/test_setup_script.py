"""``setup.sh`` 复用旧环境之前的体检：坏环境必须在**装依赖之前**被拦住。

背景（真踩过，见 CHANGELOG 0.8.1）：开发机的 venv 是按 Python 3.12 建的，后来系统的
``/usr/bin/python3.12`` 被删、``python3`` 指向 3.14。venv 里的 ``bin/python3`` 是**通用**
链接（``-> /usr/bin/python3``）而不是钉死版本，于是解释器悄悄漂到 3.14，而依赖还躺在
``lib/python3.12/site-packages`` 里。结果很反直觉：

* ``bin/python -c ""`` 照样返回 0 —— 老版体检只测这一句，于是判定"可以复用"；
* 可 ``import jinja2`` 与 ``-m pip`` 全部失败，脚本最终停在一句 ``No module named pip``。

这组用例把那三种坏法各钉一条，并确认正常环境**不会**被误判。
"""

from __future__ import annotations

import os
import re
import shutil
import subprocess
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
SETUP = ROOT / "setup.sh"

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
    """按 ``EXCEL_CODEGEN_VENV`` 指定环境跑一遍 ``setup.sh --no-test``。

    PATH 里放一个指向当前解释器的 ``python3`` 垫片，并**屏蔽 uv** —— 否则装了 uv 的机器
    会走 ``uv sync`` 那条路，根本到不了这里要测的体检分支。
    """
    shim = tmp_path / "pathbin"
    shim.mkdir()
    (shim / "python3").symlink_to(sys.executable)

    env = dict(os.environ)
    env["PATH"] = f"{shim}:/usr/bin:/bin"
    env["EXCEL_CODEGEN_VENV"] = str(env_dir)
    return subprocess.run(
        ["bash", str(SETUP), "--no-test"],
        cwd=ROOT,
        env=env,
        capture_output=True,
        text=True,
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
    assert "解释器被换掉了" in proc.stderr, proc.stderr
    assert recorded in proc.stderr and actual in proc.stderr, "报错要把两个版本都写出来才算可操作"
    assert "--recreate" in proc.stderr
    # 关键：在碰 pip 之前就停住，不再以裸的 "No module named pip" 收场
    assert "pip 升级" not in proc.stdout
    assert "安装 excel_codegen" not in proc.stdout


@posix_only
def test_site_packages_mismatch_is_caught(env_dir: Path, tmp_path: Path) -> None:
    """版本号对得上，但包目录不在当前解释器的搜索路径上（换 Python 后的遗留形态）。"""
    _site_packages(env_dir).parent.rename(env_dir / "lib" / "python9.9")

    proc = _run_setup(env_dir, tmp_path)
    assert proc.returncode != 0
    assert "site-packages" in proc.stderr, proc.stderr
    assert "pip 升级" not in proc.stdout


@posix_only
def test_missing_pip_is_caught(env_dir: Path, tmp_path: Path) -> None:
    """没有 pip 的环境复用不了：下一句就是 ``-m pip install``。"""
    proc = _run_setup(env_dir, tmp_path)
    assert proc.returncode != 0
    assert "pip 用不了" in proc.stderr, proc.stderr


@posix_only
def test_usable_environment_is_reused(env_dir: Path, tmp_path: Path) -> None:
    """体检过关就照旧复用 —— 别把好环境也判死。"""
    _stub_pip(env_dir)

    proc = _run_setup(env_dir, tmp_path)
    assert proc.returncode == 0, proc.stderr
    assert "复用已有环境" in proc.stdout
    assert "跑不起来" not in proc.stderr
