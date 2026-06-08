@echo off
setlocal EnableExtensions
chcp 65001 >nul
cd /d "%~dp0"

set "PYCMD=python"
%PYCMD% --version >nul 2>&1 || set "PYCMD=py"
%PYCMD% --version >nul 2>&1 || goto :err_python

if not exist "app\main.py" goto :err_main
if not exist "pack.py" goto :err_pack

echo Installing project dependencies from requirements.txt...
set "PIP_TRIES=0"
:deps_retry
set /a PIP_TRIES+=1
echo Pip attempt %PIP_TRIES%/4...
%PYCMD% -m pip install -r requirements.txt --retries 10 --timeout 120
if not errorlevel 1 goto :deps_ok
if %PIP_TRIES% geq 4 goto :err_pip
echo Network timeout. Retrying in 5 seconds...
timeout /t 5 /nobreak >nul
goto :deps_retry

:deps_ok

echo Building EXE via pack.py...
%PYCMD% pack.py
if errorlevel 1 goto :err_build

echo.
echo Done. Run: dist\RUN_THIS_EXE.bat
echo      or: dist\Electroproject.exe
pause
exit /b 0

:err_python
echo Error: Python was not found (python/py).
pause
exit /b 1

:err_main
echo Error: app\main.py is missing in project folder.
pause
exit /b 1

:err_pack
echo Error: pack.py is missing in project folder.
pause
exit /b 1

:err_pip
echo Error: failed to install dependencies from requirements.txt.
pause
exit /b 1

:err_build
echo Error: EXE build failed.
pause
exit /b 1
