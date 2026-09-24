@echo off
chcp 936 >nul
setlocal
cd /d "%~dp0"

echo.
echo === latex-omml skill 安装 ===
echo.
echo 本脚本做三件事：
echo   1. 在当前目录建独立虚拟环境 .venv （不污染系统 Python）
echo   2. 安装 requirements.txt 里的三个依赖
echo   3. 跑一遍内置自检，确认能正常工作
echo.

REM ---------- 1. 找 Python ----------
set "PY="
where py >nul 2>nul
if not errorlevel 1 set "PY=py -3"
if defined PY goto foundpy

where python >nul 2>nul
if not errorlevel 1 set "PY=python"
if defined PY goto foundpy

echo [X] 没找到 Python。
echo.
echo     请到 https://www.python.org/downloads/ 安装 3.9 以上版本，
echo     安装时务必勾选 "Add Python to PATH"，然后重新运行本脚本。
echo.
pause
exit /b 1

:foundpy
echo [1/3] 找到 Python:
%PY% -c "import sys; print('      版本', sys.version.split()[0])"
echo.

REM ---------- 2. 建虚拟环境 ----------
if exist ".venv\Scripts\python.exe" goto hasvenv
echo [2/3] 正在创建虚拟环境 .venv ...
%PY% -m venv .venv
if errorlevel 1 goto venvfail
echo      完成
goto step3

:hasvenv
echo [2/3] .venv 已存在，跳过
goto step3

:venvfail
echo [X] 虚拟环境创建失败。
pause
exit /b 1

REM ---------- 3. 装依赖 ----------
:step3
echo [3/3] 正在安装依赖（首次约 1-2 分钟）...
".venv\Scripts\python.exe" -m pip install --upgrade pip -q
".venv\Scripts\python.exe" -m pip install -r requirements.txt -q -i https://pypi.tuna.tsinghua.edu.cn/simple
if not errorlevel 1 goto selftest

echo      清华源失败，改用官方源重试...
".venv\Scripts\python.exe" -m pip install -r requirements.txt -q
if errorlevel 1 goto pipfail

REM ---------- 4. 自检 ----------
:selftest
echo.
echo === 自检 ===
".venv\Scripts\python.exe" scripts\latex2omml.py
if errorlevel 1 goto testfail
echo.
".venv\Scripts\python.exe" scripts\mtef.py --selftest
if errorlevel 1 goto testfail

echo.
echo === 安装完成 ===
echo.
echo 以后这样用（必须用 .venv 里的 python）：
echo.
echo   .venv\Scripts\python.exe scripts\md2docx.py input.md -o output.docx
echo   .venv\Scripts\python.exe scripts\audit_math.py "C:\讲义目录"
echo   .venv\Scripts\python.exe scripts\wmf_symbols.py "讲义.docx" --json sym.json
echo   .venv\Scripts\python.exe scripts\mtef.py "讲义.docx" --json eqs.json
echo   .venv\Scripts\python.exe scripts\mtef2docx.py "讲义.docx" --apply
echo.
echo 要装到别的机器：把整个 latex-omml 文件夹拷过去，再跑一次本脚本。
echo.
pause
exit /b 0

:pipfail
echo [X] 依赖安装失败，请检查网络后重试。
pause
exit /b 1

:testfail
echo.
echo [X] 自检未通过，上面的 PASS/FAIL 列表里能看到是哪一项。
pause
exit /b 1
