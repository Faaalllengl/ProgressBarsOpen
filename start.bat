@echo off
setlocal EnableExtensions
cd /d "%~dp0"

set "SERVER=%~dp0server.py"

where py >nul 2>&1
if %errorlevel% equ 0 (
  py "%SERVER%"
  goto :finish
)

where python >nul 2>&1
if %errorlevel% equ 0 (
  python "%SERVER%"
  goto :finish
)

echo.
echo Python ne nayden.
echo Skachayte i ustanovite: https://www.python.org/downloads/
echo Pri ustanovke vklyuchite galochku: Add python.exe to PATH
echo.
pause
exit /b 1

:finish
echo.
pause
endlocal
