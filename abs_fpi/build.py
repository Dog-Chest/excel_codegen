"""ABS FPI 模板库：一条命令把三本工作簿刷新 + 验证完。

    python build.py            # compose → 建骨架（缺哪个建哪个）→ 渲染写回 → 导出 → 验证
    python build.py --check    # 只验证，不动文件
    python build.py --init     # 骨架不存在时也重建（**会清掉你填的参数**）

三本工作簿
----------
| YAML                    | 工作簿                    | 说明 |
| ---                     | ---                       | --- |
| `abs_fpi_external.yaml` | `abs_fpi_external.xlsx`   | 规则集单用：外压 5A-3-2/5.5，3 个样例工况 |
| `abs_fpi_internal.yaml` | `abs_fpi_internal.xlsx`   | 规则集单用：内压 5A-3-2/5.7，4 个样例工况 |
| `abs_fpi.yaml`          | `ABS_FPI_load_cases.xlsx` | **项目工作簿**（compose 生成）：两套规则、共用一张 Global 表 |

三本都是 `engine: excel`（公式模式）：改参数后 Excel / WPS 打开即重算，不用跑脚本。
本脚本负责的是"改模板 / 加规则集"这一侧：重新写公式、重新导出代码，并把两边验一遍。

验证有两层（外加一条独立的交叉比对，见第 3 条）
----------------------------------------------
1. `verify_excel_engine.py` —— 项目侧自己写的公式求值器：把 `Output` 表里的 **公式在 Python 里
   算一遍**，与同一份模板的 Python 渲染结果逐行比对。它的解析器是独立实现的
   （拼接 / `IF` / `ISBLANK` / `TEXT` / `INDEX`+`MATCH`，不含算术），数值格式化复用工具的
   `utils.to_text` 以保持语义一致 —— 所以它是"工具自带求值器有没有算错"的第二重证据。
   0.4.0 起工具自带的 `check` 在公式模式下**也会**求值比对（见第 2 条），这一层因此从
   "补缺口"变成了"独立复核"；一旦模板用上 `derived:`（含算术），本脚本会报"不能识别的公式
   片段"，那时以工具自带的 `excel-codegen check --values` 为准（见 FINDINGS 0.5.0）。
2. `excel-codegen check` —— 工具自带：公式 / 快照是否与当前 YAML + 参数一致；
   公式模式下默认把公式算一遍再比对（`--no-values` 可关）。CI 可用。
3. `compare_with_rules.js` —— 与另一套已经过量纲校验的生成器（同级目录的 `../../GeniE/Rules`）
   逐行比对。该仓库在**本仓库之外**，缺失时脚本会打印 `SKIP` 并以退出码 2 结束
   （`--allow-missing` 可视为跳过）；它不属于本脚本的流程，需要时手动跑 `node compare_with_rules.js`。
"""

from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

import compose as composer  # noqa: E402
import fill_cases  # noqa: E402
import verify_excel_engine as verifier  # noqa: E402

from excel_codegen import (  # noqa: E402
    create_template,
    export_files,
    load_config,
    render_all,
    write_results,
)

#: (YAML, 工作簿, 样例工况名) —— 项目工作簿放最后：它的 YAML 由 compose 生成
PROJECTS = [
    ("abs_fpi_external.yaml", "abs_fpi_external.xlsx", list(fill_cases.EXTERNAL)),
    ("abs_fpi_internal.yaml", "abs_fpi_internal.xlsx", list(fill_cases.INTERNAL)),
    ("abs_fpi.yaml", fill_cases.PROJECT, list(fill_cases.EXTERNAL) + list(fill_cases.INTERNAL)),
]

GENERATED = HERE / "generated"


def verify(yaml_name: str, book_name: str) -> tuple[int, int]:
    """把 Output 表里的公式算一遍，与 Python 渲染结果逐行比对。"""
    from openpyxl import load_workbook

    cfg = load_config(HERE / yaml_name)
    book = HERE / book_name
    out = render_all(cfg, book)
    checks = failures = 0
    wb = load_workbook(book)
    try:
        for template in cfg.templates:
            if template.engine != "excel":
                continue
            results = out.results[template.name]
            if not results:
                continue
            got = verifier.evaluate_block(wb, cfg, template, len(results))
            for res, lines in zip(results, got, strict=False):
                checks += 1
                if lines == res.lines:
                    continue
                failures += 1
                print(f"    FAIL  {template.name} / {res.case_name}")
                for k in range(max(len(lines), len(res.lines))):
                    a = lines[k] if k < len(lines) else "<缺>"
                    b = res.lines[k] if k < len(res.lines) else "<多>"
                    if a != b:
                        print(f"            第 {k + 1} 行：公式算出 {a!r} / Python {b!r}")
                        break
    finally:
        wb.close()
    return checks, failures


def cli_check(yaml_name: str, book_name: str) -> bool:
    proc = subprocess.run(
        [sys.executable, "-m", "excel_codegen", "check", "-c", str(HERE / yaml_name), "-x", str(HERE / book_name)],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        check=False,  # 退出码自己判断：1 = 已过期，是正常结果而不是异常
    )
    text = (proc.stdout or "") + (proc.stderr or "")
    line = [
        row.strip()
        for row in text.splitlines()
        if row.strip().startswith("OK ") or row.strip().startswith("ERROR ") or "参数指纹" in row
    ]
    print(f"  [check] {book_name:26s} {'OK' if proc.returncode == 0 else 'STALE':5s} {line[0] if line else ''}")
    return proc.returncode == 0


def refresh(yaml_name: str, book_name: str, cases: list[str], *, allow_init: bool) -> None:
    yaml_path = HERE / yaml_name
    book_path = HERE / book_name
    cfg = load_config(yaml_path)

    if not book_path.exists():
        if not allow_init:
            raise SystemExit(f"{book_name} 不存在：加 --init 首次生成骨架")
        print(f"[init]  {book_name}: 生成骨架 + 灌入 {len(cases)} 个样例工况")
        create_template(cfg, book_path, cases=cases, overwrite=True)
        fill_cases.main(["fill_cases.py", str(book_path)])

    out = render_all(cfg, book_path)
    write_results(book_path, cfg, out.results, command=f"python build.py  ({yaml_name})")
    files = export_files(cfg, out.results, GENERATED)

    skipped = sum(len(v) for v in getattr(out, "skipped", {}).values())
    print(
        f"[ok]    {book_name}: {len(out.cases)} 工况 × {len(cfg.templates)} 模板 "
        f"→ {len(files)} 个导出文件" + (f"，case_filter 跳过 {skipped}" if skipped else "")
    )


def main(argv: list[str]) -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--check", action="store_true", help="只验证，不写文件")
    ap.add_argument("--init", action="store_true", help="骨架不存在时也重建")
    args = ap.parse_args(argv)

    if not args.check:
        print("== compose  规则集 → 项目工作簿 YAML")
        composer.compose(composer.MANIFEST)

        print("\n== 写入     渲染 → 写回 Excel → 导出代码")
        for yaml_name, book_name, cases in PROJECTS:
            refresh(yaml_name, book_name, cases, allow_init=args.init)

    print("\n== 验证     公式求值 ↔ Python 渲染")
    total = failed = 0
    for yaml_name, book_name, _ in PROJECTS:
        checks, failures = verify(yaml_name, book_name)
        total += checks
        failed += failures
        print(
            f"  [公式] {book_name:26s} {checks:3d} 项, {failures} 项失败"
            + ("   <= 公式算出来的代码与 Python 渲染不一致" if failures else "")
        )

    print("\n== 验证     excel-codegen check")
    ok = all(cli_check(y, b) for y, b, _ in PROJECTS)

    print(f"\n公式求值 {total} 项 / {failed} 项失败；check {'全过' if ok else '有过期'}")
    return 1 if (failed or not ok) else 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
