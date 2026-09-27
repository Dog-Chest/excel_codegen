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
git remote add origin git@github.com:<用户名>/excel_codegen.git
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

推送后 `.github/workflows/ci.yml` 会自动在 **3 个系统 × 2 个 Python 版本**上跑测试。

---

## 发布到 PyPI（可选，尚未做）

`uv build` 可以产出 wheel 与 sdist。发布前记得把 `pyproject.toml` 里注释掉的
`Documentation` URL 换成真地址 —— 占位符在 PyPI 页面上会变成一个打不开的链接。
