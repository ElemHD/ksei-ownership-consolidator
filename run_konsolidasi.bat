@echo off
chcp 65001 > nul
echo =======================================================================
echo   KONSOLIDASI DATA KEPEMILIKAN EFEK KSEI (TXT KE EXCEL STATA-READY)
echo =======================================================================
echo.

IF EXIST ".venv\Scripts\python.exe" (
    SET "PY_EXEC=.venv\Scripts\python.exe"
) ELSE (
    SET "PY_EXEC=python"
)

"%PY_EXEC%" ksei_txt_to_excel.py %*

echo.
pause
