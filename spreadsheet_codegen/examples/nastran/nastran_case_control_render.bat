@echo off
REM ===========================================================================
REM  由 spreadsheet_codegen 生成 —— 改完参数双击本文件即可把结果写回 Output 表。
REM  重新生成工作簿（init）时会一并覆盖本文件。
REM ===========================================================================
REM  本文件是 UTF-8 编码，先切到 65001 代码页，否则中文提示在 GBK 控制台是乱码。
chcp 65001 >nul
cd /d "%~dp0"

REM 依次尝试：uv + 安装好的命令 / uv + 模块 / 本机 python + 模块。
REM SPREADSHEET_CODEGEN_PY 可以指定一个"能跑 -m spreadsheet_codegen"的解释器。
set "SPREADSHEET_CODEGEN_OK="

where uv >nul 2>nul
if %errorlevel%==0 (
  uv run spreadsheet-codegen render -c "nastran_case_control.yaml" -x "nastran_case_control.xlsx" --write-excel && set "SPREADSHEET_CODEGEN_OK=1"
  if defined SPREADSHEET_CODEGEN_OK goto :done
  uv run python -m spreadsheet_codegen render -c "nastran_case_control.yaml" -x "nastran_case_control.xlsx" --write-excel && set "SPREADSHEET_CODEGEN_OK=1"
  if defined SPREADSHEET_CODEGEN_OK goto :done
)

if defined SPREADSHEET_CODEGEN_PY (
  "%SPREADSHEET_CODEGEN_PY%" -m spreadsheet_codegen render -c "nastran_case_control.yaml" -x "nastran_case_control.xlsx" --write-excel && set "SPREADSHEET_CODEGEN_OK=1"
  if defined SPREADSHEET_CODEGEN_OK goto :done
)

spreadsheet-codegen render -c "nastran_case_control.yaml" -x "nastran_case_control.xlsx" --write-excel && set "SPREADSHEET_CODEGEN_OK=1"
if defined SPREADSHEET_CODEGEN_OK goto :done
python -m spreadsheet_codegen render -c "nastran_case_control.yaml" -x "nastran_case_control.xlsx" --write-excel && set "SPREADSHEET_CODEGEN_OK=1"
if defined SPREADSHEET_CODEGEN_OK goto :done

:done
echo.
if defined SPREADSHEET_CODEGEN_OK (
  echo [OK] Done. Open the Output sheet in Excel to see the result.
) else (
  echo [FAILED] None of the ways to run spreadsheet_codegen worked.
  echo           Tried: uv run spreadsheet-codegen / uv run python -m spreadsheet_codegen
  echo                  %%SPREADSHEET_CODEGEN_PY%% -m spreadsheet_codegen / spreadsheet-codegen / python -m spreadsheet_codegen
  echo           Please make sure the tool is installed: pip install spreadsheet-codegen
  echo           or set SPREADSHEET_CODEGEN_PY to a Python that can run -m spreadsheet_codegen.
  echo           Also check that Excel is not holding this file open.
)
REM pause 放在结论之后：双击运行时窗口留得住、看得见结论。
REM 退出码必须**在 pause 之后显式给** —— echo / pause 这些内建命令会把 ERRORLEVEL
REM 归零，不写这两行的话"五条路全失败"照样以 0 结束，批处理与 CI 都看不出失败。
pause
if defined SPREADSHEET_CODEGEN_OK (exit /b 0) else (exit /b 1)
