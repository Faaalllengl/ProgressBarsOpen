@echo off
chcp 65001 >nul
setlocal EnableExtensions
title ProgressBarsOpen - запуск и обновление
cd /d "%~dp0"

set "OK=[ OK ]"
set "INFO=[ .. ]"
set "WARN=[ !! ]"
set "ERR=[ERR ]"

echo.
echo ============================================================
echo              ProgressBarsOpen - запуск проекта
echo ============================================================
echo.

echo %INFO% Проверка Git...
where git >nul 2>&1
if errorlevel 1 goto :git_error
echo %OK% Git найден

echo.
echo %INFO% Проверка локальных изменений кода...
for /f "delims=" %%S in ('git status --porcelain 2^>nul') do (
  echo %WARN% Найдены локальные изменения:
  git status --short
  echo.
  echo Сначала сохраните их в отдельный коммит или разберите вручную.
  echo Автоматическое обновление остановлено, чтобы не затереть работу.
  goto :error
)
echo %OK% Локальных изменений нет

echo.
echo %INFO% Проверка обновлений на GitHub...
git fetch origin
if errorlevel 1 goto :network_error

set "BEHIND=0"
for /f "delims=" %%N in ('git rev-list --count HEAD..origin/main 2^>nul') do set "BEHIND=%%N"

if "%BEHIND%"=="0" (
  echo %OK% Установлена последняя версия
) else (
  echo %INFO% Найдено новых коммитов: %BEHIND%
  echo %INFO% Обновление проекта...
  git pull --ff-only
  if errorlevel 1 goto :update_error
  echo %OK% Проект обновлён
)

echo.
echo %INFO% Запуск программы...
echo ------------------------------------------------------------
call start.bat
set "EXIT_CODE=%errorlevel%"
echo ------------------------------------------------------------
if not "%EXIT_CODE%"=="0" goto :error
echo %OK% Работа завершена
goto :finish

:git_error
echo %ERR% Git не найден в системе.
echo Установите Git: https://git-scm.com/download/win
goto :error

:network_error
echo %ERR% Не удалось проверить GitHub.
echo Проверьте подключение к интернету.
goto :error

:update_error
echo %ERR% Не удалось выполнить безопасное обновление.
echo Возможно, локальная ветка разошлась с GitHub.
goto :error

:error
echo.
echo Обновление или запуск остановлены.
pause
exit /b 1

:finish
endlocal
exit /b 0
