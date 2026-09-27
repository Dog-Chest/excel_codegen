"""把 ``Local Parameter`` 表的 Case 列填成真实工况。

``excel-codegen init`` 只会把每个变量的 YAML ``default`` 复制到每一列 Case，
所以"逐工况不同"的取值必须有人填 —— 这个脚本就是那个"人"。
它只动 ``Local Parameter`` 的 E 列及右侧，不碰 Global 表、不碰模板表。

用法::

    python fill_cases.py external.xlsx
    python fill_cases.py internal.xlsx
"""

from __future__ import annotations

import sys
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

FIRST_CASE_COLUMN = 5  # E 列
LOCAL_SHEET = "Local Parameter"

# --------------------------------------------------------------------------- #
# 外压 ABS FPI 5A-3-2/5.5 -- 3 个工况（取自原手写脚本）
# --------------------------------------------------------------------------- #
EXTERNAL: dict[str, dict[str, object]] = {
    "EXT-T20.559-mu90-kf+1": {
        "draft": 20.559,
        "mu_deg": 90,
        "k_c": 1.0,
        "k_f0": 1.0,
        "k_u": 1.1,
        "beta_EPS": 0.733,
        "beta_EPP": 0.733,
        "k_esf": 1.1,
        "x_o": 51,
    },
    "EXT-T15-mu0-kf-1": {
        "draft": 15,
        "mu_deg": 0,
        "k_c": 0.5,
        "k_f0": -1.0,
        "k_u": 1.1,
        "beta_EPS": 0.733,
        "beta_EPP": 0.733,
        "k_esf": 1.1,
        "x_o": 51,
    },
    "EXT-T15-mu90-kf+1": {
        "draft": 15,
        "mu_deg": 90,
        "k_c": 1.0,
        "k_f0": 1.0,
        "k_u": 1.1,
        "beta_EPS": 0.733,
        "beta_EPP": 0.733,
        "k_esf": 1.1,
        "x_o": 51,
    },
}

# --------------------------------------------------------------------------- #
# 内压 ABS FPI 5A-3-2/5.7 -- 4 个工况
#
# 注意：本工具的局部变量只有"每个 Case 一列"一种作用域，没有"舱"这一层，
# 所以舱的数据只能按工况摊平（同名舱在两个工况里各写一遍）。
# WBT6 / WBT7 / COT1 三个舱的原始数据见 Rules/gen.js 的 INT_TANKS。
# --------------------------------------------------------------------------- #
_WBT6 = {
    "rho_tank": 1025,
    "l_tank": 42,
    "b_tank": 32,
    "h_tank": 32,
    "eta_deck": 0,
    "eta_overflow": 0,
    "C_dp": 1,
    "C_ru": 1,
    "p_vp": 0,
    "GM_full_in": 0,
    "k_r_in": 0,
    "tank_is_ballast": 1,
    "member_11_17": 0,
}
_WBT7 = {
    "rho_tank": 1025,
    "l_tank": 14.5,
    "b_tank": 10.15,
    "h_tank": 32,
    "eta_deck": 0,
    "eta_overflow": 0,
    "C_dp": 1,
    "C_ru": 1,
    "p_vp": 0,
    "GM_full_in": 0,
    "k_r_in": 0,
    "tank_is_ballast": 1,
    "member_11_17": 0,
}
_COT1 = {
    "rho_tank": 900,
    "l_tank": 42,
    "b_tank": 32,
    "h_tank": 32,
    "eta_deck": 0,
    "eta_overflow": 0,
    "C_dp": 1,
    "C_ru": 1,
    "p_vp": 0.21,
    "GM_full_in": 0,
    "k_r_in": 0,
    "tank_is_ballast": 0,
    "member_11_17": 1,
}
_COMMON = {
    "k_u": 1.1,
    "k_esf": 1.1,
    "w_v": 0.75,
    "w_l": 0.25,
    "w_t": 0.75,
    "beta_VAC": 0.660,
    "beta_LAC": 0.758,
    "beta_TAC": 0.557,
    "beta_PMO": 0.686,
    "beta_RMO": 0.492,
    "delta_b": 0,
    "delta_h": 0,
}


def _case(tank_ref: str, tank: dict, **over: object) -> dict[str, object]:
    row: dict[str, object] = {"tank_ref": tank_ref}
    row.update(_COMMON)
    row.update(tank)
    row.update(over)
    return row


INTERNAL: dict[str, dict[str, object]] = {
    "WBT6-d8-mu90": _case("WBT6", _WBT6, draft=8, mu_deg=90, k_c=1.0, xi=21),
    "WBT6-d15.059-mu0": _case("WBT6", _WBT6, draft=15.059, mu_deg=0, k_c=0.4, xi=21),
    "WBT7-d20.559-mu90": _case("WBT7", _WBT7, draft=20.559, mu_deg=90, k_c=1.0, xi=7.25),
    "COT1-d20.559-mu60": _case("COT1", _COT1, draft=20.559, mu_deg=60, k_c=0.5, xi=21),
}

BOOKS = {"abs_fpi_external.xlsx": EXTERNAL, "abs_fpi_internal.xlsx": INTERNAL}
#: 合成项目工作簿（compose.py 生成）：外压在前、内压在后 —— 顺序很重要，
#: 见 README「为什么内压在后面」。这里只补两个例子用的最小工况集。
PROJECT = "ABS_FPI_load_cases.xlsx"

_HEAD_FILL = PatternFill("solid", fgColor="EDEDED")
_HEAD_FONT = Font(bold=True)
_TOP = Alignment(vertical="top", wrap_text=False)


def _write(path: Path, cases: dict[str, dict[str, object]], *, header: bool = True) -> None:
    """把 cases 写进 Local 表的 E 列起；header=False 时沿用表里已有的 Case 名。"""
    workbook = load_workbook(path)
    sheet = workbook[LOCAL_SHEET]

    row_of: dict[str, int] = {}
    for row in range(2, sheet.max_row + 1):
        name = sheet.cell(row=row, column=1).value
        if name:
            row_of[str(name).strip()] = row

    unknown = sorted({k for values in cases.values() for k in values} - set(row_of))
    if unknown:
        raise SystemExit(f"{path.name}: 这些变量在 {LOCAL_SHEET} 里没有行: {', '.join(unknown)}")

    for offset, (case_name, values) in enumerate(cases.items()):
        column = FIRST_CASE_COLUMN + offset
        if header:
            head = sheet.cell(row=1, column=column, value=case_name)
            head.font = _HEAD_FONT
            head.fill = _HEAD_FILL
            head.alignment = Alignment(vertical="center", horizontal="center")
        else:
            existing = sheet.cell(row=1, column=column).value
            if str(existing).strip() != case_name:
                raise SystemExit(
                    f"{path.name}: 第 {column} 列表头是 {existing!r}，期望 {case_name!r}"
                    "（列顺序必须与 fill_cases.py 一致，因为 CASE 列是位置对齐的）"
                )
        for name, value in values.items():
            sheet.cell(row=row_of[name], column=column, value=value).alignment = _TOP
        sheet.column_dimensions[get_column_letter(column)].width = max(16, min(36, len(case_name) + 6))

    workbook.save(path)
    print(f"{path.name}: 写入 {len(cases)} 个工况 -> " + ", ".join(cases))


def fill(path: Path) -> None:
    """规则集工作簿：表头 + 取值都由这里写。"""
    _write(path, BOOKS[path.name])


def fill_project(path: Path) -> None:
    """合成项目工作簿：Case 名已由 --cases 建好，这里只按顺序填取值。"""
    cases: dict[str, dict[str, object]] = {}
    cases.update(EXTERNAL)
    for name, values in INTERNAL.items():
        row = {"kind": "INT"}
        row.update(values)
        cases[name] = row
    for name in list(cases):
        cases[name] = dict(cases[name], kind=cases[name].get("kind", "EXT"))
    _write(path, cases, header=False)


def main(argv: list[str]) -> int:
    if len(argv) != 2:
        print(__doc__)
        return 2
    path = Path(argv[1])
    if path.name == PROJECT:
        fill_project(path)
    else:
        fill(path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
