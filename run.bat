@echo off
rem 双击运行「定时关机助手」，无控制台窗口
setlocal
cd /d "%~dp0"

set "PYW="
where pyw >nul 2>nul && set "PYW=pyw -3"
if not defined PYW where pythonw >nul 2>nul && set "PYW=pythonw"

if defined PYW (
    start "" %PYW% "%~dp0shutdown_timer.py"
) else (
    echo 未找到 Python，请先安装 Python 3.8+ 并勾选 "Add to PATH"。
    pause
)
endlocal
