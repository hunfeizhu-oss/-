@echo off
setlocal EnableExtensions
chcp 65001 >nul
cd /d "%~dp0"
set "LOG_FILE=%~dp0windows-start.log"

echo ============================================== > "%LOG_FILE%"
echo Battle Report Tool Windows Start Log >> "%LOG_FILE%"
echo Date: %date% %time% >> "%LOG_FILE%"
echo Directory: %cd% >> "%LOG_FILE%"
echo ============================================== >> "%LOG_FILE%"

call :find_python
if errorlevel 1 goto :python_missing

echo [OK] Python command: %PYTHON_CMD%
echo Python command: %PYTHON_CMD% >> "%LOG_FILE%"
%PYTHON_CMD% -c "import sys; print('Executable:', sys.executable); print('Version:', sys.version)" >> "%LOG_FILE%" 2>&1
if errorlevel 1 goto :python_broken

%PYTHON_CMD% -c "import sys; raise SystemExit(0 if sys.version_info >= (3, 9) else 1)" >> "%LOG_FILE%" 2>&1
if errorlevel 1 goto :python_old

echo [OK] 正在启动战报生成工具...
echo 如果浏览器没有自动打开，请查看窗口中显示的 http://127.0.0.1 地址。
echo.
%PYTHON_CMD% "%~dp0app.py" %*
set "APP_EXIT=%errorlevel%"
if not "%APP_EXIT%"=="0" goto :app_failed
exit /b 0

:find_python
set "PYTHON_CMD="
py -3 -c "import sys; raise SystemExit(0)" >nul 2>> "%LOG_FILE%"
if not errorlevel 1 (
    set "PYTHON_CMD=py -3"
    exit /b 0
)
python -c "import sys; raise SystemExit(0 if sys.version_info.major == 3 else 1)" >nul 2>> "%LOG_FILE%"
if not errorlevel 1 (
    set "PYTHON_CMD=python"
    exit /b 0
)
python3 -c "import sys; raise SystemExit(0 if sys.version_info.major == 3 else 1)" >nul 2>> "%LOG_FILE%"
if not errorlevel 1 (
    set "PYTHON_CMD=python3"
    exit /b 0
)
for %%V in (313 312 311 310 39) do (
    if exist "%LocalAppData%\Programs\Python\Python%%V\python.exe" (
        set "PYTHON_CMD="%LocalAppData%\Programs\Python\Python%%V\python.exe""
        exit /b 0
    )
)
exit /b 1

:python_missing
echo.
echo [ERROR] 未找到可用的 Python 3。
echo 请安装 Python 3.9 或更高版本，并在安装时勾选 Add Python to PATH。
echo 也可以双击 install-windows.bat，尝试自动安装 Python 后启动。
echo 日志文件：%LOG_FILE%
pause
exit /b 1

:python_broken
echo.
echo [ERROR] Python 命令存在，但无法正常运行。
echo 请打开日志文件查看详情：%LOG_FILE%
pause
exit /b 1

:python_old
echo.
echo [ERROR] Python 版本过低。请安装 Python 3.9 或更高版本。
echo 日志文件：%LOG_FILE%
pause
exit /b 1

:app_failed
echo.
echo [ERROR] 工具启动失败，错误代码：%APP_EXIT%
echo 请将此窗口中的报错和日志文件发给维护人员：
echo %LOG_FILE%
pause
exit /b %APP_EXIT%
