# 发布与隐私（维护者向）

本仓库是**公开**的。往里面加东西之前过一遍这三条。

---

## 1. 不要提交密钥

`.gitignore` 已经挡住常见落点（`.env*`、`*.pem` / `*.key`、`secrets.*`、`.aws/`、
各家大模型 CLI 的配置目录……），但**它只在文件还没被跟踪时有效** —— 一旦 `git add` 过，
内容就永久留在历史里。提交前扫一遍：

```bash
# 只打印命中位置，不打印内容（避免把密钥二次暴露到终端 / 日志）
git grep -lIE 'sk-[A-Za-z0-9_-]{20,}|AIza[0-9A-Za-z_-]{35}|(AKIA|ASIA)[0-9A-Z]{16}|gh[pousr]_[A-Za-z0-9]{36,}|-----BEGIN [A-Z ]*PRIVATE KEY-----' -- .
```

**万一真提交了密钥**：① 立刻去服务商后台**吊销并轮换**（改历史救不了已泄露的凭据）；
② 改 `.gitignore`；③ 用 `git filter-repo` / BFG 清历史后强推；④ 只要推过公开仓库，
就按"已经泄露"处理 —— 旧提交即使被 force-push 覆盖，仍可能按 SHA 取到，直到 GitHub GC。

---

## 2. 打开 GitHub 的两个免费开关

仓库页 → **Settings → Code security and analysis**：

- **Secret scanning** —— 自动扫已知格式的密钥并告警（公开仓库默认开）。
- **Push protection** —— **确认它是开的**。它在 `git push` 那一刻就拦下含密钥的提交，
  比事后补救有用得多（公开仓库默认开；私有仓库要手动开且需要 Advanced Security）。

---

## 3. 提交邮箱

commit 里的邮箱会被 GitHub **永久公开**。建议用 noreply 地址，并在
**Settings → Emails** 勾上「Keep my email addresses private」与
「Block command line pushes that expose my email」。

本仓库当前作者是 `Dog-Chest <129145708+Dog-Chest@users.noreply.github.com>`。
换机器开发时记得同步（否则新提交又会带出真实邮箱）：

```bash
git config --global user.name  "Dog-Chest"
git config --global user.email "129145708+Dog-Chest@users.noreply.github.com"
```

---

## 首次推送到 GitHub

```bash
# 1) 网页上 New repository —— 不要勾 README / .gitignore / license（本地都有）
# 2) 关联并推送
git remote add origin git@github.com:<用户名>/spreadsheet_codegen.git
git push -u origin main
```

认证方式二选一：

* **SSH（推荐，配一次长期可用）**
  ```bash
  ssh-keygen -t ed25519 -C "<你的 noreply 邮箱>"
  cat ~/.ssh/id_ed25519.pub      # 整段复制 → GitHub → Settings → SSH and GPG keys
  ssh -T git@github.com          # 出现 "Hi <用户名>!" 即成功
  ```
* **HTTPS + Personal Access Token**：GitHub 早已不支持密码，用
  Settings → Developer settings → Personal access tokens（`Contents: Read and write`）。
  注意 token 会明文落在 `~/.git-credentials`，共享机器上不要用。

推送后 `.github/workflows/ci.yml` 会自动在 **3 个系统 × 2 个 Python 版本**上跑
测试、lint、类型检查与覆盖率门槛。

---

## 发布到 PyPI

发布流程已经在 [`.github/workflows/release.yml`](../.github/workflows/release.yml) 里了，
用 **Trusted Publishing（OIDC）** —— 仓库里不存任何 token，GitHub 每次发布时向 PyPI
换一张短期凭据。

> ⚠️ **改名说明**：本项目原名 `excel-codegen`（因为名字里带 Microsoft 商标 "Excel"，
> 有商标风险而改名）。**PyPI 项目名不能改**，所以：
>
> * `excel-codegen` 留在 PyPI 上，从 `0.9.2` 起是**弃用跳板版**（README 引导到新名），
>   并把旧版本**yank** 掉 —— 新装的人装不到，已锁定的依赖不受影响；
> * 本项目用**新名字 `spreadsheet-codegen` 重新起步**，tag 从 `v0.10.0` 开始。
>
> 也就是说：**新名字需要重新登记一次 pending publisher**（下面第 2 步），
> 而 GitHub 那边的 `pypi` 环境可以直接复用。

### 一次性配置（新名字要重做第 2 步）

1. **注册 PyPI 账号**并开启两步验证（发版必需）。
2. **登记 pending publisher**：PyPI → Account → Publishing → *Add a pending publisher*：

   | 字段 | 填什么 |
   | --- | --- |
   | PyPI Project Name | `spreadsheet-codegen` |
   | Owner | `Dog-Chest` |
   | Repository name | `spreadsheet_codegen` |
   | Workflow name | `release.yml` |
   | Environment name | `pypi` |

   > 项目名写成 `spreadsheet-codegen`（连字符）是因为**分发的名字**与**导入的名字**可以不同：
   > `pip install spreadsheet-codegen` 装进来的是 `import spreadsheet_codegen`。
3. GitHub 仓库 → Settings → Environments → **`pypi`**（已在旧名下建好，直接复用）。
4. 确认 `pyproject.toml` 的 `[project.urls]` 都是真地址。
   **当前四项都指向 `github.com/Dog-Chest/spreadsheet_codegen`** —— 前提是 GitHub 仓库
   已经改成这个名字（见下）。

### 改名要一起做的几件事

| 动作 | 在哪做 | 备注 |
| --- | --- | --- |
| GitHub 仓库改名 `excel_codegen` → `spreadsheet_codegen` | 仓库 Settings → Repository name | 旧 URL 会自动重定向，外链不丢；**改名前文档里的新链接会 404** |
| 新名登记 pending publisher | PyPI → Account → Publishing | 见上面第 2 步 |
| 旧名发弃用版 `0.9.2` + yank | 旧 tag/分支 + PyPI 页面 | yank 在 PyPI 的 *Manage → Releases* 里逐版本操作 |

### 每次发版

```bash
# 1) 改版本号（pyproject.toml 与 spreadsheet_codegen/__init__.py 两处）+ 写 CHANGELOG
# 2) 重新锁定（版本号变化会写进 uv.lock）—— CI 有 uv lock --check，忘了会红
uv lock
# 3) 本地过一遍质量闸
uv run ruff check . && uv run ruff format --check . && uv run mypy && uv run pytest --cov
# 4) 提交、打 tag、推（tag 必须与 pyproject 版本一致，release.yml 会校验）
git commit -am "0.9.0：……"
git tag v0.9.0
git push origin main --tags
```

> 示例：本次发布准备的是 `0.9.0`。**推 tag 那一刻才会真正发到 PyPI** ——
> 想先演练就走 Actions 页面手动触发 `release`、`target` 选 `testpypi`。

tag 一推，`release.yml` 就会：校验 tag 与 `pyproject.toml` 版本一致 → 再跑一遍质量闸 →
`uv build` → `twine check` → 打包内置示例（`examples.zip`）→ 发到 PyPI →
**建 GitHub Release**（挂上 `dist/*`：wheel / sdist / `examples.zip`，说明用
`--generate-notes` 生成，已存在则跳过）。

> `examples.zip` 是给"不想装 Python 的人"的：里面的工作簿是公式模式，用 Excel / WPS
> 打开就能填参数、自动重算。

> **先演练**：Actions 页面手动触发 `release`，`target` 选 `testpypi`，会发到 TestPyPI
> （用 `--index-url https://test.pypi.org/simple/` 装来验证）。TestPyPI 与 PyPI 是两个
> 独立的账号体系，需要在 TestPyPI 上单独登记一次 pending publisher。

### 发布之后，别人怎么装

```bash
uv tool install spreadsheet-codegen      # 装成全局命令
# 或
pipx install spreadsheet-codegen
# 或
pip install spreadsheet-codegen
```

> 装完 `spreadsheet-codegen examples --copy ./examples` 就能拿到带工作簿的内置示例。

---

### 排错：`invalid-publisher`

```
Trusted publishing exchange failure:
* `invalid-publisher`: valid token, but no corresponding publisher
```

**含义**：OIDC token 本身没问题（签名有效），只是 PyPI 上**没有能对上的 (pending) publisher**。
**一个字节都没上传**，所以版本号没有被烧掉 —— 改完直接重跑即可。

对着日志里渲染出来的 claims 逐字核对 PyPI 表单（下面是 `v0.9.0` 的实际值）：

| 日志里的 claim | PyPI 表单字段 | 实际值 |
| --- | --- | --- |
| `repository_owner` | Owner | `Dog-Chest` |
| `repository` | Repository name | `spreadsheet_codegen`（**下划线**） |
| `workflow_ref` 的文件名 | Workflow name | `release.yml`（**只写文件名**，别写路径） |
| `environment` | Environment name | `pypi`（**留空就对不上**） |
| 发行包 METADATA 的 `Name` | PyPI Project Name | `spreadsheet_codegen` → 归一化成 `spreadsheet-codegen` |

最常见的两个错：**Workflow name 写成 `.github/workflows/release.yml`**、
**Environment name 留空**（workflow 里用的是 `pypi`）。
另外确认加的是 **pypi.org 生产站**的 pending publisher，不是 TestPyPI 的。

**改完怎么重发**（不用重新走一遍 build）：

1. 首选：Actions 页面打开上次失败的 run → **Re-run failed jobs**
   （用 API 则是 `POST /repos/{owner}/{repo}/actions/runs/{id}/rerun-failed-jobs`，
   需要 `actions: write` 权限的 token）；
2. 没有那个权限时：**删掉 tag 再重推**，同样安全（版本没烧）：
   ```bash
   git push origin --delete v0.9.0 && git push origin v0.9.0
   ```

这样同事就不需要克隆仓库了 —— 「多平台零配置」这条线到此才算闭环。
