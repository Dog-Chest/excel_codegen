#!/usr/bin/env bash
# ============================================================================
# spreadsheet_codegen —— Linux / macOS 开发环境搭建
#
#   ./setup.sh                              # 有 uv 用 uv（推荐），没有就退回 venv+pip
#   ./setup.sh --no-test                    # 只装环境，不跑 pytest
#   ./setup.sh --recreate                   # 删掉旧环境重建
#   SPREADSHEET_CODEGEN_VENV=/data/v ./setup.sh   # 指定环境位置（uv 与 pip 两条路都认）
#
# 两条路：
#   A) 机器上有 uv  → `uv sync`：自动挑/下载 Python、按 uv.lock 装依赖，**不需要**
#      python3-venv、不需要 pip、不需要 apt。跨平台同一条命令（Windows 上就是
#      PowerShell 里的 `uv run pytest`，连这个脚本都不需要）。
#   B) 机器上没有 uv → 退回 `python3 -m venv` + `pip install -e ".[dev]"`（需要
#      Ubuntu 的 python3-venv 包）。
#
# 环境放哪：venv 与 pip 会写上万个碎文件，
#   * 仓库在 NTFS / exFAT 外置盘上（从 Windows 拷过来的常见情形）时容易踩坑：
#     本项目实测在 ntfs3 挂载上 `rm -rf .venv` 会令进程停在 D（不可中断睡眠）状态；
#   * 在 ext4 / xfs / btrfs 等原生文件系统上，直接放仓库里的 `.venv/` 没问题。
# 所以本脚本先探测仓库所在文件系统，再决定环境放哪里。
#
# 复用旧环境之前会做一次体检（解释器版本 / site-packages / pip）。换系统、或把旧的
# Python 删掉之后，venv 里的 bin/python3 是通用链接，会**悄悄**漂到新解释器上，而依赖
# 还留在旧的 lib/pythonX.Y 里 —— 这种环境"能执行"却什么也 import 不到（连 pip 都没了）。
# 体检过不了就报出具体原因并提示 --recreate，绝不让它拖到 `pip install` 才炸。
# ============================================================================
set -euo pipefail

RUN_TESTS=1
RECREATE=0
for arg in "$@"; do
  case "$arg" in
    --no-test)  RUN_TESTS=0 ;;
    --recreate) RECREATE=1 ;;
    -h|--help)
      sed -n '2,21p' "$0" | sed 's/^# \{0,1\}//'
      exit 0 ;;
    *) echo "未知参数：${arg}（可用：--no-test / --recreate / --help）" >&2; exit 2 ;;
  esac
done

HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$HERE"

say()  { printf '%s\n' "$*"; }
warn() { printf '!  %s\n' "$*" >&2; }
die()  { printf 'ERROR  %s\n' "$*" >&2; exit 1; }

# --- 体检：已有的 venv 还能不能用 -------------------------------------------
# 只测 `bin/python -c ""` 是不够的，它漏掉最阴的一种坏法：
#   venv 建在 3.12 上，之后系统的 python3.12 被删、python3 指向 3.14；venv 里的
#   bin/python3 是**通用**链接（-> /usr/bin/python3）而不是钉死版本，于是解释器
#   悄悄漂到 3.14，可依赖还躺在 lib/python3.12/site-packages 里。
#   结果 `python -c ""` 照样返回 0，而所有 import（连 pip 本身）都没了。
# 所以这里逐项验：解释器能跑 → 版本没被换 → site-packages 在当前解释器上 → pip 可用。
# 返回 0 表示**有问题**，并把一句人读的原因打到 stdout（供调用方拼进报错）。
venv_problem() {
  local env_dir="$1"
  local vpy="$env_dir/bin/python"
  local want got
  if ! "$vpy" -c "" >/dev/null 2>&1; then
    printf '%s' '解释器执行不起来（多半是它依赖的 Python 被删掉了）'
    return 0
  fi
  # pyvenv.cfg 记着"当初按哪个版本建的"，与现在真跑的对不上就是被换掉了
  want="$(sed -n 's/^version *= *//p' "$env_dir/pyvenv.cfg" 2>/dev/null | head -1)"
  got="$("$vpy" -c 'import sys; print("%d.%d.%d" % sys.version_info[:3])' 2>/dev/null)"
  if [ -n "$want" ] && [ "$want" != "$got" ]; then
    printf '解释器被换掉了：环境是按 Python %s 建的，现在跑的是 %s' "$want" "$got"
    return 0
  fi
  if ! "$vpy" -c 'import os, sys; sys.exit(0 if os.path.isdir(os.path.join(sys.prefix, "lib", "python%d.%d" % sys.version_info[:2], "site-packages")) else 1)' >/dev/null 2>&1; then
    printf 'site-packages 不在当前解释器（%s）的搜索路径上' "$got"
    return 0
  fi
  if ! "$vpy" -m pip --version >/dev/null 2>&1; then
    printf '%s' 'pip 用不了（import 不到）'
    return 0
  fi
  return 1
}

# --- 1. 选路：uv 优先 -------------------------------------------------------
USE_UV=0
if command -v uv >/dev/null 2>&1; then
  USE_UV=1
  say "[1/5] 找到 uv $("uv" --version 2>/dev/null | awk '{print $2}')  → 走零配置路径（不需要 python3-venv / pip）"
else
  PYTHON="${PYTHON:-python3}"
  # 下限必须与 pyproject.toml 的 requires-python 一致，否则这里放行、pip 那边才报错
  MIN_PY="3.11"
  command -v "$PYTHON" >/dev/null 2>&1 || die "既没有 uv，也没有 ${PYTHON}。装一个 Python $MIN_PY+，或装 uv（https://docs.astral.sh/uv/）。"
  "$PYTHON" - "$MIN_PY" <<'PY' || die "Python 版本太低：本项目要求 $MIN_PY+（与 pyproject.toml 的 requires-python 一致）。"
import sys
need = tuple(int(part) for part in sys.argv[1].split("."))
raise SystemExit(0 if sys.version_info[: len(need)] >= need else 1)
PY
  say "[1/5] 没找到 uv，走经典路径：$("$PYTHON" -V 2>&1)（建议装 uv 以便跨平台零配置）"
fi

# --- 2. 探测文件系统，决定环境放哪 -----------------------------------------
fs_type_of() {
  local target="$1"
  [ -r /proc/mounts ] || { printf 'unknown\n'; return; }
  local best="" bestlen=0 type="" dev mnt fstype rest
  while read -r dev mnt fstype rest; do
    mnt="${mnt//\\040/ }"
    case "$target/" in
      "$mnt"/*)
        if [ "${#mnt}" -gt "$bestlen" ]; then best="$mnt"; bestlen="${#mnt}"; type="$fstype"; fi
        ;;
    esac
  done < /proc/mounts
  printf '%s\n' "${type:-unknown}"
}

REPO_FS="$(fs_type_of "$HERE")"
DEFAULT_VENV="$HERE/.venv"
case "$REPO_FS" in
  ntfs*|exfat|vfat|msdos|fuseblk|hfsplus|unknown)
    DEFAULT_VENV="$HOME/.venvs/spreadsheet_codegen"
    warn "仓库所在文件系统是 ${REPO_FS}（非原生 Linux 文件系统）。"
    say  "       环境默认改到 $DEFAULT_VENV —— 上万个碎文件写在 $REPO_FS 上既慢又容易卡。"
    say  "       （仓库本身放在这里没问题，只有 .venv 要挪走。）"
    ;;
  *)
    say "[2/5] 仓库文件系统 ${REPO_FS}，环境可直接放在仓库内的 .venv/  ✓"
    ;;
esac

# 兼容改名前的变量名（EXCEL_CODEGEN_VENV）：新名优先，老写法仍认
ENV_DIR="${SPREADSHEET_CODEGEN_VENV:-${EXCEL_CODEGEN_VENV:-$DEFAULT_VENV}}"
say "       环境位置：$ENV_DIR"

if [ "$RECREATE" = "1" ] && [ -d "$ENV_DIR" ]; then
  say "       --recreate：删除旧环境 $ENV_DIR"
  rm -rf "$ENV_DIR"
fi

# --- 3. 装依赖 --------------------------------------------------------------
if [ "$USE_UV" = "1" ]; then
  export UV_PROJECT_ENVIRONMENT="$ENV_DIR"
  say "[3/5] uv sync（按 uv.lock 锁定版本；缺 Python 3.12 时 uv 会自己下载）..."
  uv sync --quiet
  say "       ✓ 已同步 spreadsheet_codegen $(uv version --short 2>/dev/null || echo '')"
else
  if ! "$PYTHON" -c "import ensurepip" >/dev/null 2>&1; then
    warn "当前 Python 缺少 ensurepip，无法用 \`python3 -m venv\` 建虚拟环境。"
    say  "       Ubuntu / Debian 上装一下即可（顺带装 pip 与 git）："
    say  "           sudo apt install python3-venv python3-pip git"
    say  "       或者干脆装 uv，从此不需要 python3-venv："
    say  "           curl -LsSf https://astral.sh/uv/install.sh | sh"
    die  "缺少 python3-venv，已停止。"
  fi
  if [ -x "$ENV_DIR/bin/python" ]; then
    # 复用之前先体检（原因见函数上方的注释）：坏掉的旧环境必须在装依赖**之前**拦住，
    # 否则会停在 `pip install` 那里，报一句与真正病因无关的 "No module named pip"。
    if reason="$(venv_problem "$ENV_DIR")"; then
      warn "已有环境 $ENV_DIR 跑不起来：${reason}。"
      die  "用 ./setup.sh --recreate 删掉重建（工作簿与代码都在仓库里，删环境不会丢东西）。"
    fi
    say "[3/5] 复用已有环境（要重建加 --recreate）"
  else
    say "[3/5] 创建 venv ..."
    mkdir -p "$(dirname "$ENV_DIR")"
    "$PYTHON" -m venv "$ENV_DIR"
  fi
  VPY="$ENV_DIR/bin/python"
  say "       pip 升级 ..."
  "$VPY" -m pip install --quiet --upgrade pip
  say "       安装 spreadsheet_codegen[dev]（可编辑安装）..."
  if ! "$VPY" -m pip install --quiet -e ".[dev]"; then
    warn "安装失败。最常见的两种原因："
    say  "       1) 这个环境是上次装到一半留下的 —— 用 ./setup.sh --recreate 删掉重建；"
    say  "       2) $ENV_DIR 的权限不对（例如以前用 sudo 建过）—— 换一个位置："
    say  "          SPREADSHEET_CODEGEN_VENV=~/venvs/spreadsheet_codegen ./setup.sh"
    die  'pip install -e ".[dev]" 失败。'
  fi
  say "       ✓ 已安装：$("$VPY" -m pip show spreadsheet_codegen 2>/dev/null | awk '/^Version:/{print $2}')"
fi

# --- 4. 测试 ----------------------------------------------------------------
if [ "$RUN_TESTS" = "1" ]; then
  say "[4/5] 跑测试 ..."
  if [ "$USE_UV" = "1" ]; then
    uv run --quiet pytest
  else
    "$ENV_DIR/bin/python" -m pytest
  fi
else
  say "[4/5] 跳过测试（--no-test）"
fi

# --- 5. 提示：外部工具（可选的独立验证链要用）-------------------------------
say "[5/5] 检查可选工具 ..."
command -v node >/dev/null 2>&1 || warn "没找到 node：abs_fpi/compare_with_rules.js（第 4 层验证）跑不了。"
[ -d "$HERE/../GeniE/Rules" ] || warn "没找到同级目录 ../GeniE/Rules：第 4 层交叉校验会以 SKIP + 退出码 2 结束。"
command -v git  >/dev/null 2>&1 || warn "没找到 git：本仓库还没有版本控制，建议 git init。"

say ""
say "环境就绪。以后每条命令都这样跑（Windows / macOS / Linux 完全一样）："
if [ "$USE_UV" = "1" ]; then
  say "    uv run pytest"
  say "    uv run spreadsheet-codegen --version"
  say "    uv run python abs_fpi/build.py --check"
else
  say "    source \"$ENV_DIR/bin/activate\""
  say "    spreadsheet-codegen --version        # 或 python -m spreadsheet_codegen"
  say "    pytest                         # 全部用例"
fi
