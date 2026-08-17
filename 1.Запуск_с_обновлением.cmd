@echo off
chcp 65001 >nul
setlocal EnableExtensions
title ProgressBarsOpen - запуск и обновление
cd /d "%~dp0"

set "OK=[ OK ]"
set "INFO=[ .. ]"
set "WARN=[ !! ]"
set "ERR=[ERR ]"
set "SERVER=%~dp0server.py"
set "IS_RUNTIME_UPDATE=0"
if /i "%~1"=="--update" set "IS_RUNTIME_UPDATE=1"

if "%IS_RUNTIME_UPDATE%"=="1" (
  echo %INFO% Ожидание остановки текущей версии...
  timeout /t 3 /nobreak >nul
)

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
for /f "delims=" %%S in ('git status --porcelain 2^>nul') do goto :dirty_worktree
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

:launch
echo.
echo %INFO% Запуск программы...
echo ------------------------------------------------------------

where python >nul 2>&1
if errorlevel 1 goto :try_py
python "%SERVER%"
set "EXIT_CODE=%errorlevel%"
goto :server_finished

:try_py
where py >nul 2>&1
if errorlevel 1 goto :python_error
py "%SERVER%"
set "EXIT_CODE=%errorlevel%"
goto :server_finished

:dirty_worktree
echo %WARN% Найдены локальные изменения. Проверка обновлений пропущена.
echo %INFO% Программа будет запущена без обновления.
goto :launch

:server_finished
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

:python_error
echo %ERR% Python не найден в системе.
echo Установите Python и включите Add python.exe to PATH.
goto :error

:error
echo.
echo Запуск или обновление остановлены.
pause
exit /b 1

:finish
endlocal
exit /b 0
