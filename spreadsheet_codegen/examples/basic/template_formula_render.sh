#!/usr/bin/env bash
# ===========================================================================
#  由 spreadsheet_codegen 生成 —— 改完参数跑一次本文件即可把结果写回 Output 表。
#  重新生成工作簿（init）时会一并覆盖本文件。
# ===========================================================================
set -uo pipefail
cd "$(dirname "$0")"

if command -v uv >/dev/null 2>&1; then
  uv run spreadsheet-codegen render -c "example_formula.yaml" -x "template_formula.xlsx" --write-excel
else
  spreadsheet-codegen render -c "example_formula.yaml" -x "template_formula.xlsx" --write-excel
fi
status=$?

echo
if [ "$status" -ne 0 ]; then
  echo "[失败] 上面有报错信息。常见原因：依赖没装（跑一次 ./setup.sh 或 uv sync）、"
  echo "       或者 Excel / WPS 正开着这个文件（先关掉再试）。"
else
  echo "[完成] 回到 Excel 打开「Output」表看结果。"
fi
exit "$status"
