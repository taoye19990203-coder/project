@echo off
chcp 65001 >nul
cd /d "%~dp0"
where py >nul 2>nul
if errorlevel 1 goto use_python
py -3 bootstrap.py
set "audit_exit=%errorlevel%"
goto finished
:use_python
where python >nul 2>nul
if errorlevel 1 goto missing_python
python bootstrap.py
set "audit_exit=%errorlevel%"
goto finished
:missing_python
echo 请先安装 Python 3.10 或以上版本。
set "audit_exit=1"
:finished
if not "%audit_exit%"=="0" echo 启动未完成，请查看上方具体错误信息。
pause
exit /b %audit_exit%
