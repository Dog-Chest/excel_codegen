#!/usr/bin/env bash
# ===========================================================================
#  由 excel_codegen 生成 —— 改完参数跑一次本文件即可把结果写回 Output 表。
#  重新生成工作簿（init）时会一并覆盖本文件。
# ===========================================================================
set -uo pipefail
cd "$(dirname "$0")"

if command -v uv >/dev/null 2>&1; then
  uv run excel-codegen render -c "abs_fpi.yaml" -x "ABS_FPI_load_cases.xlsx" --write-excel
else
  excel-codegen render -c "abs_fpi.yaml" -x "ABS_FPI_load_cases.xlsx" --write-excel
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
