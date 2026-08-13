@echo off
chcp 65001 >nul
title ProgressBarsOpen - Git Update

echo.
echo ==========================================
echo       ProgressBarsOpen - UPDATE
echo ==========================================
echo.

echo [1/4] Проверка Git...
git --version >nul 2>&1

if errorlevel 1 (
    echo [ERROR] Git НЕ установлен!
    echo.
    echo Устанавливаю Git...
    echo.

    winget install --id Git.Git -e --source winget

    echo.
    echo ==========================================
    echo Git установлен.
    echo Перезапустите этот файл.
    echo ==========================================
    pause
    exit /b
)

echo [OK] Git установлен:
git --version

echo.
echo [2/4] Определение папки проекта...
cd /d "%~dp0"

echo [OK] Папка:
echo %CD%

echo.
echo [3/4] Проверка обновлений...
echo.

git fetch

if errorlevel 1 (
    echo.
    echo [ERROR] Не удалось получить данные с GitHub.
    echo Проверьте подключение к интернету.
    pause
    exit /b
)

echo.
echo Текущий статус:
echo ------------------------------------------
git status -uno
echo ------------------------------------------

echo.
echo [4/4] Обновление проекта...
echo.

git pull

if errorlevel 1 (
    echo.
    echo ==========================================
    echo [ERROR] Обновление НЕ выполнено!
    echo ==========================================
    echo.
    pause
    exit /b
)

echo.
echo ==========================================
echo          [OK] ОБНОВЛЕНИЕ ЗАВЕРШЕНО
echo ==========================================
echo.
echo Дата и время:
echo %date% %time%
echo.

pause