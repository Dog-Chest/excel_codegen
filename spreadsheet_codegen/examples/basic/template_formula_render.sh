#!/usr/bin/env bash
# ===========================================================================
#  由 spreadsheet_codegen 生成 —— 改完参数跑一次本文件即可把结果写回 Output 表。
#  重新生成工作簿（init）时会一并覆盖本文件。
# ===========================================================================
set -uo pipefail
cd "$(dirname "$0")"

# 依次尝试：uv + 安装好的命令 / uv + 模块 / 本机 python + 模块。
# SPREADSHEET_CODEGEN_PY 可以指定一个"能跑 -m spreadsheet_codegen"的解释器。
status=0
solved=0
# 跑一条路；成功（退出码 0）就定案，后面的路不再执行。
# 函数本身恒定返回 0：调用方只用全局的 status / solved，不靠返回值判断。
run_route() {
  if [ "$solved" -ne 0 ]; then return 0; fi
  "$@"
  status=$?
  if [ "$status" -eq 0 ]; then solved=1; fi
  return 0
}

if command -v uv >/dev/null 2>&1; then
  run_route uv run spreadsheet-codegen render -c "example_formula.yaml" -x "template_formula.xlsx" --write-excel
  run_route uv run python -m spreadsheet_codegen render -c "example_formula.yaml" -x "template_formula.xlsx" --write-excel
fi

if [ -n "${SPREADSHEET_CODEGEN_PY:-}" ]; then
  run_route "$SPREADSHEET_CODEGEN_PY" -m spreadsheet_codegen render -c "example_formula.yaml" -x "template_formula.xlsx" --write-excel
fi

# 最后两条：命令与模块。命令**存在**时只试它 —— 见函数文档：它若报错，那是工具
# 真的失败，退出码要原样带出去（CI 靠 1/2/3/4 分流），不该被"没装"的假象盖住。
if command -v spreadsheet-codegen >/dev/null 2>&1; then
  run_route spreadsheet-codegen render -c "example_formula.yaml" -x "template_formula.xlsx" --write-excel
else
  run_route python3 -m spreadsheet_codegen render -c "example_formula.yaml" -x "template_formula.xlsx" --write-excel
fi

echo
if [ "$solved" -ne 0 ]; then
  echo "[完成] 回到 Excel 打开「Output」表看结果。"
else
  echo "[失败] 这些调用方式都没成功：uv run spreadsheet-codegen / uv run python -m spreadsheet_codegen /"
  echo "       spreadsheet-codegen / python3 -m spreadsheet_codegen。"
  echo "       请确认已安装（pip install spreadsheet-codegen），"
  echo "       或设置 SPREADSHEET_CODEGEN_PY 指向能跑 -m spreadsheet_codegen 的 Python。"
  echo "       也请确认 Excel / WPS 没有占着这个文件。"
fi
exit "$status"
