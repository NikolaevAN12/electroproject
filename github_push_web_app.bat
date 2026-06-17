@echo off
chcp 65001 >nul
setlocal EnableDelayedExpansion
cd /d "%~dp0"

rem Ищем git: системный PATH или portable в .tools
set "GIT="
where git >nul 2>&1 && set "GIT=git"
if not defined GIT if exist ".tools\MinGit\cmd\git.exe" set "GIT=.tools\MinGit\cmd\git.exe"
if not defined GIT if exist ".tools\PortableGit\cmd\git.exe" set "GIT=.tools\PortableGit\cmd\git.exe"

if not defined GIT (
    echo.
    echo Git не найден.
    echo.
    echo 1. Откройте: https://github.com/git-for-windows/git/releases
    echo 2. Скачайте MinGit-*-64-bit.zip
    echo 3. Распакуйте содержимое в папку:
    echo    %CD%\.tools\MinGit
    echo    ^(должен появиться файл .tools\MinGit\cmd\git.exe^)
    echo 4. Запустите этот скрипт снова.
    echo.
    pause
    exit /b 1
)

set "REPO_URL=%~1"
if "%REPO_URL%"=="" (
    set /p REPO_URL=URL репозитория GitHub ^(https://github.com/user/repo.git^): 
)
if "%REPO_URL%"=="" (
    echo URL не указан.
    pause
    exit /b 1
)

echo Используется: %GIT%
echo Репозиторий: %REPO_URL%
echo Ветка: web-app
echo.

if not exist .git (
    "%GIT%" init
)

"%GIT%" checkout web-app 2>nul
if errorlevel 1 "%GIT%" checkout -b web-app

"%GIT%" add -A
"%GIT%" status --short
echo.
"%GIT%" diff --cached --quiet
if errorlevel 1 (
    "%GIT%" commit -m "Web version: FastAPI UI, LAN access, fire check MV cable"
) else (
    echo Нет новых изменений для коммита.
)

"%GIT%" remote get-url origin >nul 2>&1
if errorlevel 1 (
    "%GIT%" remote add origin "%REPO_URL%"
) else (
    "%GIT%" remote set-url origin "%REPO_URL%"
)

echo.
echo Отправка ветки web-app на GitHub...
echo При запросе логина: имя — ваш GitHub-логин, пароль — Personal Access Token.
echo.
"%GIT%" push -u origin web-app
set "PUSH_ERR=!errorlevel!"

echo.
if "!PUSH_ERR!"=="0" (
    echo Готово. Ветка web-app отправлена. Ветка main на GitHub не изменялась.
) else (
    echo Ошибка push. Проверьте URL, токен и доступ к репозиторию.
)
echo.
pause
exit /b %PUSH_ERR%
