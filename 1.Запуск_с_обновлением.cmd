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

set "BRANCH="
for /f "delims=" %%B in ('git branch --show-current 2^>nul') do set "BRANCH=%%B"
if not defined BRANCH goto :branch_error
set "REMOTE_REF=origin/%BRANCH%"

set "BEHIND="
for /f "delims=" %%N in ('git rev-list --count HEAD..%REMOTE_REF% 2^>nul') do set "BEHIND=%%N"
if not defined BEHIND goto :compare_error

if "%BEHIND%"=="0" (
  echo %OK% Установлена последняя версия
) else (
  echo %INFO% Найдено новых коммитов: %BEHIND%
  if "%IS_RUNTIME_UPDATE%"=="1" (
    echo %INFO% Обновление проекта...
    git pull --ff-only
    if errorlevel 1 goto :update_error
    echo %OK% Проект обновлён
  ) else (
    echo %INFO% Обновление будет предложено в интерфейсе.
  )
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

:branch_error
echo %ERR% Не удалось определить текущую ветку Git.
goto :error

:compare_error
echo %ERR% Не удалось сравнить локальную ветку с GitHub.
echo Проверьте, что на GitHub существует ветка %BRANCH%.
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
