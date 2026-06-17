@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo Выгрузка на GitHub без установки Git (через API)
echo.
echo Нужен Personal Access Token: GitHub - Settings - Developer settings - Tokens
echo Право: repo (или Fine-grained: Contents Read and Write)
echo.

set "REPO_URL=%~1"
if "%REPO_URL%"=="" (
    set /p REPO_URL=URL репозитория ^(https://github.com/user/repo^): 
)

if "%GITHUB_TOKEN%"=="" (
    if exist ".github_token" (
        echo Токен взят из файла .github_token
    ) else (
        set /p GITHUB_TOKEN=Вставьте токен GitHub: 
    )
)

echo.
python github_upload_api.py "%REPO_URL%"
set "ERR=%errorlevel%"
echo.
if not "%ERR%"=="0" (
    echo Если ошибка SSL, попробуйте:
    echo   set GITHUB_SSL_NO_VERIFY=1
    echo   github_upload_web_app_api.bat "%REPO_URL%"
)
pause
exit /b %ERR%
