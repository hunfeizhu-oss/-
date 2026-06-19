@echo off
setlocal EnableExtensions
chcp 65001 >nul
cd /d "%~dp0"

echo ==============================================
echo 智启新程战报生成器 - Windows 环境诊断
echo ==============================================
echo.
call start.bat --check
echo.
echo 诊断结束。日志文件位于：
echo %~dp0windows-start.log
pause
