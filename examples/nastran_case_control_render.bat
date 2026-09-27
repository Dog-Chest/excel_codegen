@echo off
REM ===========================================================================
REM  由 excel_codegen 生成 —— 改完参数双击本文件即可把结果写回 Output 表。
REM  重新生成工作簿（init）时会一并覆盖本文件。
REM ===========================================================================
cd /d "%~dp0"

where uv >nul 2>nul
if %errorlevel%==0 (
  uv run excel-codegen render -c "nastran_case_control.yaml" -x "nastran_case_control.xlsx" --write-excel
) else (
  excel-codegen render -c "nastran_case_control.yaml" -x "nastran_case_control.xlsx" --write-excel
)

echo.
if errorlevel 1 (
  echo [失败] 上面有报错信息。常见原因：依赖没装（跑一次 setup.sh / uv sync）、
  echo        或者 Excel 正开着这个文件（先关掉再试）。
) else (
  echo [完成] 回到 Excel 打开「Output」表看结果。
)
pause
