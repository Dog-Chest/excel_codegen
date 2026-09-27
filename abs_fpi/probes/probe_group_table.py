"""探针 G：第三层作用域（Tank Data 成员表）在 ABS FPI 里到底省了什么。

0.8.0 起 `abs_fpi_internal.yaml` 不再把舱数据按工况摊平，而是放进
`variables.group`（成员表 `Tank Data`，一行一个舱，指南 §18）。这个探针按顺序证明四件事：

    python probe_group_table.py

1. **舱只写一遍**：`Tank Data` 表 3 行，Local 表里已经没有 l_tank / h_tank 这些行；
2. **改一次舱数据 → 所有用到它的工况跟着变**（WBT6 被 2 个工况用到）；
3. **改某个 Case 的 tank_ref 指针 → 它的舱数据整组换掉**；
4. **指针指向不存在的成员 → 明确报错并列出可选成员**（不是悄悄回落默认值）。

外加一条：`case_name` / 派生参数 / 公式模式都照常工作 —— 第 4 步用的是
`excel-codegen check`（工具自带，含"把公式在 Python 里算一遍"）。

注意：探针只在 `probes/_out/` 里的**副本**上动手，不会碰仓库里的工作簿。
"""

from __future__ import annotations

import shutil
import subprocess
import sys
from pathlib import Path

from openpyxl import load_workbook

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent  # abs_fpi/
YAML = ROOT / "abs_fpi_internal.yaml"
BOOK = ROOT / "abs_fpi_internal.xlsx"
OUT = HERE / "_out"
TEMPLATE = "genie_int"


def banner(text: str) -> None:
    print("\n" + "=" * 78)
    print(text)
    print("=" * 78)


def load():
    sys.path.insert(0, str(ROOT))
    from excel_codegen import load_config

    return load_config(YAML)


def render(cfg, book: Path) -> dict[str, list[str]]:
    from excel_codegen import render_all

    out = render_all(cfg, book)
    return {result.case_name: result.lines for result in out.results[TEMPLATE]}


def line_of(lines: list[str], prefix: str) -> str:
    return next(item for item in lines if item.startswith(prefix))


def fresh_copy() -> Path:
    OUT.mkdir(parents=True, exist_ok=True)
    target = OUT / "probe_group_table.xlsx"
    shutil.copy(BOOK, target)
    return target


def set_cell(book: Path, sheet: str, *, row: str, column: str, value: object) -> None:
    """按**名字**定位成员行 / 变量列（这里刻意不用坐标，免得探针本身被布局绑死）。"""
    workbook = load_workbook(book)
    worksheet = workbook[sheet]
    rows = {worksheet.cell(row=r, column=1).value: r for r in range(2, worksheet.max_row + 1)}
    columns = {worksheet.cell(row=1, column=c).value: c for c in range(2, worksheet.max_column + 1)}
    worksheet.cell(row=rows[row], column=columns[column], value=value)
    workbook.save(book)
    workbook.close()


CASES = ("WBT6-d8-mu90", "WBT6-d15.059-mu0", "WBT7-d20.559-mu90", "COT1-d20.559-mu60")


def step1(cfg) -> None:
    banner("① 舱只写一遍：成员表 3 行 vs Local 表里没有舱数据行")
    workbook = load_workbook(BOOK)
    try:
        group = workbook["Tank Data"]
        local = workbook["Local Parameter"]
        print("Tank Data 表：")
        for row in group.iter_rows(min_row=1, max_row=group.max_row, max_col=4, values_only=True):
            print("   ", row)
        local_names = {
            str(local.cell(row=r, column=1).value).strip()
            for r in range(2, local.max_row + 1)
            if local.cell(row=r, column=1).value
        }
        leaked = sorted(local_names & {v.name for v in cfg.group_variables})
        print(f"\nLocal 表里的行: {len(local_names)} 个")
        print(f"成员表里的变量: {[v.name for v in cfg.group_variables]}")
        print(f"Local 表里残留的舱变量: {leaked or '（无）'}")
    finally:
        workbook.close()


def step2(cfg, book: Path) -> None:
    banner("② 改一次舱数据 → 用到它的工况全跟着变")
    before = render(cfg, book)
    print("改前：")
    for name in CASES:
        print(f"    {name:22s} {line_of(before[name], 'var l_tank')}")
    set_cell(book, "Tank Data", row="WBT6", column="l_tank", value=30)
    after = render(cfg, book)
    print("\n把 Tank Data 里 WBT6 的 l_tank 从 42 改成 30（只动这一格）：")
    for name in CASES:
        print(f"    {name:22s} {line_of(after[name], 'var l_tank')}")
    changed = [name for name in CASES if before[name] != after[name]]
    print(f"\n内容发生变化的工况: {changed}")
    assert changed == ["WBT6-d8-mu90", "WBT6-d15.059-mu0"], changed
    print("OK  一次编辑影响 2 个工况；WBT7 / COT1 不受影响")


def step3(cfg, book: Path) -> None:
    banner("③ 改 Case 的 tank_ref 指针 → 整组舱数据换掉")
    before = render(cfg, book)
    print("把 COT1-d20.559-mu60 的 tank_ref 从 COT1 改成 WBT7：")
    _set_case_cell(book, case="COT1-d20.559-mu60", variable="tank_ref", value="WBT7")
    after = render(cfg, book)
    for name in ("COT1-d20.559-mu60", "WBT7-d20.559-mu90"):
        print(f"    {name:22s} {line_of(after[name], 'var l_tank')}  {line_of(after[name], 'var rho_tank')}")
    assert line_of(after["COT1-d20.559-mu60"], "var l_tank") == "var l_tank = 14.5 m;"
    assert after["WBT7-d20.559-mu90"] == before["WBT7-d20.559-mu90"]
    print("OK  指向哪个舱就取哪个舱的数据（另一个工况不动）")


def _set_case_cell(book: Path, *, case: str, variable: str, value: object) -> None:
    workbook = load_workbook(book)
    sheet = workbook["Local Parameter"]
    columns = {sheet.cell(row=1, column=c).value: c for c in range(5, sheet.max_column + 1)}
    rows = {sheet.cell(row=r, column=1).value: r for r in range(2, sheet.max_row + 1)}
    sheet.cell(row=rows[variable], column=columns[case], value=value)
    workbook.save(book)
    workbook.close()


def step4(cfg, book: Path) -> None:
    banner("④ 指针指向不存在的成员 → 明确报错（而不是悄悄回落 default）")
    _set_case_cell(book, case="COT1-d20.559-mu60", variable="tank_ref", value="WBT9")
    try:
        render(cfg, book)
    except Exception as exc:
        print(f"    {type(exc).__name__}: {exc}")
        print("    ↑ tank_ref 声明了 choices，所以取值约束先拦住（最贴近用户看到的那条）")
    else:  # pragma: no cover - 探针失败路径
        raise AssertionError("指向不存在的成员居然没报错")

    # 去掉 choices 之后，拦住它的是成员表查找本身 —— 这条路径也要有明确的话
    import re

    import yaml as _yaml

    # 放在 abs_fpi/ 根下：模板文件的相对路径按**配置文件所在目录**解析（指南 §16.2）
    plain = ROOT / "_probe_group_nokeys.yaml"
    raw = _yaml.safe_load(YAML.read_text(encoding="utf-8"))
    for item in raw["variables"]["local"]:
        if item["name"] == "tank_ref":
            item.pop("choices", None)
    raw["excel"]["output"] = "probe_group_nokeys.xlsx"
    plain.write_text(_yaml.safe_dump(raw, sort_keys=False, allow_unicode=True), encoding="utf-8")

    from excel_codegen import load_config

    plain_cfg = load_config(plain)
    try:
        render(plain_cfg, book)
    except Exception as exc:
        message = str(exc)
        assert "WBT9" in message and "WBT6" in message, message
        print(f"\n    （去掉 choices 后）{type(exc).__name__}: {message}")
        print("OK  报错里点名了错的成员，并列出了可选成员")
    else:  # pragma: no cover - 探针失败路径
        raise AssertionError("去掉 choices 后，指向不存在的成员居然没报错")
    finally:
        plain.unlink(missing_ok=True)
    del re


def step5(cfg, book: Path) -> None:
    banner("⑤ 工具自带 check：公式模式下的成员表链条也能算出来")
    _set_case_cell(book, case="COT1-d20.559-mu60", variable="tank_ref", value="COT1")
    workbook = load_workbook(book)
    try:
        sheet = workbook["Code"]
        formulas = [
            str(sheet.cell(row=row, column=2).value)
            for row in range(2, sheet.max_row + 1)
            if isinstance(sheet.cell(row=row, column=2).value, str)
        ]
    finally:
        workbook.close()
    hit = next((item for item in formulas if "Tank Data" in item), None)
    assert hit is not None, formulas[:3]
    print("Code 表里引用成员表的公式（节选）：")
    print("   ", hit[:200], "…")
    proc = subprocess.run(
        [sys.executable, "-m", "excel_codegen", "check", "-c", str(YAML), "-x", str(book)],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        check=False,
    )
    text = ((proc.stdout or "") + (proc.stderr or "")).strip().splitlines()
    print("\ncheck 输出：")
    for row in text[-6:]:
        print("   ", row.strip())
    print(
        f"\n退出码 {proc.returncode}"
        "（0 = 工作簿里的公式算出来与当前参数一致；1 = 过期；2 = 用法/配置错）"
    )
    assert proc.returncode == 0, text
    print("OK  公式引用到了成员表，check 能把整条链（Case → 舱 → 变量）在 Python 里算一遍")


def main() -> int:
    cfg = load()
    book = fresh_copy()
    print(f"探针副本: {book.relative_to(ROOT.parent)}")
    step1(cfg)
    step2(cfg, book)
    step3(cfg, book)
    step4(cfg, book)
    step5(cfg, book)
    print("\n探针 G 全部通过")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
