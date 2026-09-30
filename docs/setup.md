# 安装与跨平台环境

需要 **Python 3.11+**，Windows / Linux / macOS 均可。实测通过 **3.11 / 3.12 / 3.13 / 3.14**。

---

## 推荐：uv（三个系统零配置）

仓库里的 `pyproject.toml` + `uv.lock` + `.python-version` 已经配好了。装了 uv 之后，
**不需要 venv、不需要 pip、不需要 apt、也不用管本机是哪个 Python 版本**：

```bash
# 每台机器装一次 uv（三选一）
curl -LsSf https://astral.sh/uv/install.sh | sh                 # Linux / macOS
powershell -c "irm https://astral.sh/uv/install.ps1 | iex"      # Windows PowerShell
pipx install uv                                                 # 任何系统

# 之后在任何系统、任何目录，命令完全一样
uv run pytest                                   # 自动建环境 + 按 uv.lock 装依赖 + 跑测试
uv run spreadsheet-codegen --version                  # 跑 CLI，不用 activate
uv run spreadsheet-codegen examples --copy ./examples # 内置示例（带已填好样例参数的工作簿）
uv run spreadsheet-codegen render -c examples/basic/example.yaml -x examples/basic/template.xlsx --write-excel
uv run python abs_fpi/build.py --check
```

它替你做的三件事，正好就是"额外配置"的来源：

| 麻烦 | uv 怎么办 |
| --- | --- |
| 本机 Python 版本不对 / 没有 | `.python-version` 写着 `3.12`，**没有就自动下载**（不需要管理员权限） |
| 各机器装到的依赖版本不一致 | `uv.lock` 是**跨平台锁文件**，同一份锁同时含 `win_amd64` / `macosx_*` / `manylinux` / `musllinux` 的 wheel，三个系统装出来完全一致 |
| 测试依赖要额外加参数 | `[dependency-groups] dev` 是 uv 的**默认**依赖组，`uv run` 会一起装上 —— 所以 `uv run pytest` 不需要 `--extra dev` |

这个仓库只需要记住两条命令：**`uv run pytest`**（验证）和 **`uv run spreadsheet-codegen ...`**（干活）。

> **环境放哪**：uv 默认在仓库里建 `.venv/`（在原生盘上这是最优的）。仓库在外置 NTFS / exFAT
> 盘上、想把环境挪走时用环境变量（这是 uv 唯一支持的方式，`pyproject.toml` 里没有对应开关）：
>
> ```bash
> export UV_PROJECT_ENVIRONMENT=~/.venvs/spreadsheet_codegen                # Linux / macOS
> $env:UV_PROJECT_ENVIRONMENT = "$HOME\.venvs\spreadsheet_codegen"          # Windows PowerShell
> ```
>
> 或者直接跑 `./setup.sh`，它会自动探测文件系统并替你设好。
> ⚠ `setup.sh` 是 **POSIX shell 脚本**（Linux / macOS）—— Windows 上没有 bash，
> 用上面那两条 uv 命令就够了，不需要它（见下面「Windows 上怎么用」）。

---

## 经典方式：venv + pip

```bash
cd spreadsheet_codegen

python3 -m venv .venv
# Windows
.venv\Scripts\activate
# Linux / macOS
source .venv/bin/activate

pip install -e ".[dev]"        # 含测试依赖
```

安装后会得到命令 `spreadsheet-codegen`（也可用 `python -m spreadsheet_codegen`）。

`[project.optional-dependencies] dev` 与 `[dependency-groups] dev` 内容一致：pip 侧用 extra，
uv 侧用依赖组。**改一处记得改另一处。**

仓库还自带一个 `./setup.sh`，两条路都认（有 uv 走 uv，没有就退回 venv + pip）：

```bash
./setup.sh                                       # 自动选位置 + 装依赖 + 跑测试
./setup.sh --no-test                             # 只装环境
./setup.sh --recreate                            # 删掉旧环境重建（环境半坏时用这个）
SPREADSHEET_CODEGEN_VENV=/data/venvs/ecg ./setup.sh    # 也可以自己指定环境位置
```

它会做几件事：挑 uv / venv 路径、探测仓库所在文件系统决定环境放哪、装依赖、跑一遍测试，
最后提示缺哪些可选工具（`node`、同级的 `GeniE/Rules`）。几个刻意的行为：

* **Python 下限与 `pyproject.toml` 的 `requires-python` 保持一致**（3.11）——
  脚本自己先拦，不然会一路装到 `pip` 才报版本不符；
* **复用已有环境前先体检**：不只测"解释器能不能执行"，还要 `pyvenv.cfg` 记的版本与现在
  真跑的对得上、`site-packages` 在当前解释器的搜索路径上、`pip` 可用。只测前者会漏掉最阴
  的一种坏法 —— 换过 Python（或删掉旧版本）之后，venv 里的 `bin/python3` 是**通用**链接，
  会悄悄漂到新解释器上，而依赖还留在旧的 `lib/pythonX.Y` 里：`python -c ""` 照样通过，
  可连 `pip` 都 import 不到。体检不过会报出具体原因（哪两个版本对不上）并让你 `--recreate`，
  而不是把它拖到 `pip install` 才以一句 `No module named pip` 收场；
* `pip install` 失败时会指出最常见的两种原因（环境半坏 / 权限不对）与对应的下一步。

---

## Windows 上怎么用

工具本身在 Windows 上是一等公民：CLI 全部命令、生成的 `.xlsx`、以及一并生成的
`*_render.bat` 都是跨平台的，CI 也在 `windows-latest` × Python 3.11 / 3.13 上跑全套测试。
只有 `setup.sh` 是 POSIX 脚本 —— Windows 上不需要它。

**用（推荐）**：发布之后装成全局命令，不碰仓库：

```powershell
powershell -c "irm https://astral.sh/uv/install.ps1 | iex"   # 装 uv（一次性）
uv tool install spreadsheet-codegen                                # 装工具
spreadsheet-codegen examples --copy .\examples                     # 拿示例（含已填好的工作簿）
start .\examples\abs_fpi\abs_fpi_internal.xlsx               # 打开改参数
```

**开发**：在仓库目录里用 uv，命令与 Linux / macOS 完全一致（不需要 `activate`）：

```powershell
uv run pytest
uv run spreadsheet-codegen --version
```

几个 Windows 特有的点：

* **一键刷新脚本**：`init` / `render --write-excel` 会在工作簿旁边生成
  `<工作簿名>_render.bat`，**双击即可**（`.sh` 是给 Linux / macOS 的）；
* **控制台编码**：早期版本在 GBK 控制台上打印 `✓` 会抛 `UnicodeEncodeError`，已修
  （FINDINGS #1）；如果仍见到乱码，先 `chcp 65001` 切到 UTF-8；
* **Excel 占用**：工作簿正被 Excel / WPS 打开时写回会失败 —— 先关掉再跑
  （脚本的失败提示里也写了这一条）；
* **别在装依赖时 Ctrl-C**：NTFS 上写到一半被打断会留下读不动的目录（见下一节）。

---

## Ubuntu / 从 Windows 拷过来的仓库：三个坑

本仓库原本在 Windows 上开发，整目录拷到 Ubuntu 后有三件事必须处理，否则"装不上"或"跑得极慢"：

1. **不要复用 `.venv/`**。虚拟环境里记着绝对路径（`pyvenv.cfg` 的 `home=`、`Scripts/` 目录），
   从 Windows 拷过来在 Linux 上完全不可用 —— 直接删掉重建（它本来就在 `.gitignore` 里）。
2. **`python3 -m venv` 报 `ensurepip is not available`**。Ubuntu 把虚拟环境拆成了独立包：

   ```bash
   sudo apt install python3-venv python3-pip git
   ```

3. **外置 NTFS / exFAT 盘上，环境放到原生文件系统，并且别中断安装。**
   实测结论（比"NTFS 不能用"更准确的说法）：
   * 干净地写入、干净地删除一个 36MB / 1080 个文件的 venv，在 ntfs3 上是**正常且很快**的；
   * 真正会把盘卡住的是**安装写到一半被强杀**：一次被打断的 `pip install --target` 留下了一个
     读不动的 `__pycache__`，之后连 `rm -rf` / `rmdir` 都停在 `D`（不可中断睡眠）状态，
     必须重启才能清掉。慢只是次要问题（碎文件多）。

   所以：环境放原生盘更省心，且**不要在装依赖时 Ctrl-C**。

   ```bash
   python3 -m venv ~/.venvs/spreadsheet_codegen
   source ~/.venvs/spreadsheet_codegen/bin/activate
   pip install -e ".[dev]" && pytest
   ```

---

## Python 版本策略（以后加功能时怎么判断）

| 角色 | 版本 | 定义在哪 |
| --- | --- | --- |
| **支持的底线** | **3.11** | `pyproject.toml` 的 `requires-python` |
| **CI 真正测的** | 3.11（底线）+ 3.13 | `.github/workflows/ci.yml` 的矩阵 |

> `[tool.ruff]` 的 `target-version` 也写的是 `py311`，与 `requires-python` **必须一致**：
> 它负责拦住"3.12 才支持、但下限写 3.11"的语法。这条护栏是真起过作用的 ——
> 有一次 f-string 表达式里写了引号（3.12 才允许），本地 3.12 跑得好好的，
> 一 push 到 3.11 就是语法错误。
| **开发默认** | **3.12** | `.python-version`（uv 据此挑解释器，本机没有会自动下载） |

三条规矩：

1. **`requires-python` 只写"CI 真的跑过"的最老版本。** 现在写 `>=3.11`，因为 **3.10 已在
   2026-10 结束支持**（[Python 版本支持周期](https://devguide.python.org/versions/)），
   继续兼容等于替一个没人维护的解释器背锅。
2. **想用更新的语法，就同时改三处**：`requires-python`、CI 矩阵、本文这一节；
   然后跑一次 `uv lock`（依赖解析会跟着变）与 `uv run pytest`。CI 里有一条
   `uv lock --check`，改了 `pyproject.toml` 忘了重新锁会直接失败。
3. **平时不用惦记版本**：为 3.11 写的代码在 3.12 / 3.13 / 3.14 上必然也能跑，
   反方向才会出事 —— 而拦住反方向的正是"CI 测底线"。

> 新增功能时只有两类东西需要先想一下版本：
> ① 3.12+ 才有的标准库 API（`itertools.batched`、`typing.override`；`tomllib` 是 3.11 就有）；
> ② 3.12+ 的新语法（PEP 695 的 `type X = ...`、`def f[T]()`）。
> 想用就把底线抬到 3.12（一行 + CI 矩阵），不想用就继续待在 3.11 —— **没有任何东西强制你选**。

---

## 验证链的外部依赖

`abs_fpi/compare_with_rules.js`（第 4 层交叉校验）需要**本仓库之外**的同级目录
`../../GeniE/Rules`。该目录不存在时脚本会打印 `SKIP` 并以退出码 **2** 结束
（`--allow-missing` 可视为跳过、退出码 0）—— **不会**把"没验证"伪装成"通过"。
其余三层（`check` 值校验、`verify_excel_engine.py`、`build.py --check`）不依赖它。
