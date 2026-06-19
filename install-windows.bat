@echo off
setlocal EnableExtensions
chcp 65001 >nul
cd /d "%~dp0"
set "LOG_FILE=%~dp0windows-start.log"

echo ============================================== > "%LOG_FILE%"
echo Battle Report Tool Windows Install Log >> "%LOG_FILE%"
echo Date: %date% %time% >> "%LOG_FILE%"
echo Directory: %cd% >> "%LOG_FILE%"
echo ============================================== >> "%LOG_FILE%"

echo ==============================================
echo 智启新程战报生成器 - Windows 一键安装并启动
echo ==============================================
echo.
echo 本工具本身不用安装，只需要电脑有 Python 3.9+。
echo 如果缺少 Python，本脚本会尝试自动安装 Python 3.12。
echo 自动导出 PNG 还需要桌面版 Microsoft PowerPoint。
echo.

call :find_python_39
if not errorlevel 1 goto :start_tool

echo [INFO] 未检测到 Python 3.9 或更高版本。
echo [INFO] 正在检查 Windows 是否支持 winget 自动安装...
echo 未检测到 Python，准备尝试 winget 安装。>> "%LOG_FILE%"
where winget >nul 2>> "%LOG_FILE%"
if errorlevel 1 goto :manual_python

echo.
echo [INFO] 正在通过 winget 安装 Python 3.12，请根据系统弹窗确认。
echo [INFO] 如果公司电脑限制安装，请把窗口里的错误发给维护人员。
winget install -e --id Python.Python.3.12 --scope user --accept-package-agreements --accept-source-agreements >> "%LOG_FILE%" 2>&1
if errorlevel 1 goto :winget_failed

echo.
echo [INFO] Python 安装流程已结束，正在重新检测...
call :find_python_39
if not errorlevel 1 goto :start_tool

for %%V in (313 312 311 310 39) do (
    if exist "%LocalAppData%\Programs\Python\Python%%V\python.exe" (
        set "PYTHON_CMD="%LocalAppData%\Programs\Python\Python%%V\python.exe""
        goto :start_tool
    )
)

goto :path_pending

:start_tool
echo.
echo [OK] Python 已就绪：%PYTHON_CMD%
echo [INFO] 正在启动工具...
echo.
call start.bat %*
exit /b %errorlevel%

:manual_python
echo.
echo [ERROR] 当前 Windows 没有 winget，无法自动安装 Python。
echo 请手动安装 Python 3.12，并在安装界面勾选 Add Python to PATH。
echo 下载页将自动打开：
echo https://www.python.org/downloads/windows/
start "" "https://www.python.org/downloads/windows/"
echo.
echo 安装完成后，请重新双击 start.bat 或 install-windows.bat。
echo 日志文件：%LOG_FILE%
pause
exit /b 1

:winget_failed
echo.
echo [ERROR] winget 自动安装 Python 失败。
echo 常见原因：公司电脑禁用安装、网络限制、Microsoft Store/winget 不可用。
echo 请手动安装 Python 3.12，并勾选 Add Python to PATH。
echo 下载页将自动打开：
echo https://www.python.org/downloads/windows/
start "" "https://www.python.org/downloads/windows/"
echo.
echo 详细日志：%LOG_FILE%
pause
exit /b 1

:path_pending
echo.
echo [WARN] Python 可能已安装完成，但当前窗口还没有刷新 PATH。
echo 请关闭这个窗口后，重新双击 start.bat。
echo 如果仍失败，请重启电脑后再试。
echo 日志文件：%LOG_FILE%
pause
exit /b 1

:find_python_39
set "PYTHON_CMD="
py -3 -c "import sys; raise SystemExit(0 if sys.version_info >= (3, 9) else 1)" >nul 2>> "%LOG_FILE%"
if not errorlevel 1 (
    set "PYTHON_CMD=py -3"
    exit /b 0
)
python -c "import sys; raise SystemExit(0 if sys.version_info >= (3, 9) else 1)" >nul 2>> "%LOG_FILE%"
if not errorlevel 1 (
    set "PYTHON_CMD=python"
    exit /b 0
)
python3 -c "import sys; raise SystemExit(0 if sys.version_info >= (3, 9) else 1)" >nul 2>> "%LOG_FILE%"
if not errorlevel 1 (
    set "PYTHON_CMD=python3"
    exit /b 0
)
for %%V in (313 312 311 310 39) do (
    if exist "%LocalAppData%\Programs\Python\Python%%V\python.exe" (
        "%LocalAppData%\Programs\Python\Python%%V\python.exe" -c "import sys; raise SystemExit(0 if sys.version_info >= (3, 9) else 1)" >nul 2>> "%LOG_FILE%"
        if not errorlevel 1 (
            set "PYTHON_CMD="%LocalAppData%\Programs\Python\Python%%V\python.exe""
            exit /b 0
        )
    )
)
exit /b 1
