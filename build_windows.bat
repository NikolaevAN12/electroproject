@echo off
setlocal EnableExtensions
chcp 65001 >nul
cd /d "%~dp0"

set "PYCMD=python"
%PYCMD% --version >nul 2>&1 || set "PYCMD=py"
%PYCMD% --version >nul 2>&1 || goto :err_python

if not exist "main.py" goto :err_main

echo Checking PyInstaller...
%PYCMD% -m PyInstaller --version >nul 2>&1
if not errorlevel 1 (
  echo PyInstaller is already installed. Skipping pip install.
  goto :pip_ok
)

echo Installing PyInstaller...
set "PIP_TRIES=0"
:pip_retry
set /a PIP_TRIES+=1
echo Pip attempt %PIP_TRIES%/4...
%PYCMD% -m pip install pyinstaller --retries 10 --timeout 60
if not errorlevel 1 goto :pip_ok
if %PIP_TRIES% geq 4 goto :err_pip
echo Network timeout. Retrying in 5 seconds...
timeout /t 5 /nobreak >nul
goto :pip_retry

:pip_ok

echo Cleaning previous build...
if exist "build" rmdir /s /q "build"
if exist "dist\Electroproject.exe" del /f /q "dist\Electroproject.exe" 2>nul
if exist "dist\Electroproject.exe" (
  echo ERROR: dist\Electroproject.exe is in use. Close Electroproject and run build again.
  pause
  exit /b 1
)
if exist "dist" (
  for %%F in ("dist\*") do del /f /q "%%F" 2>nul
)
if exist "Electroproject.spec" del /f /q "Electroproject.spec"

echo Building EXE from main.py...
%PYCMD% -m PyInstaller --onefile --name "Electroproject" "main.py"
if errorlevel 1 goto :err_build

echo Done: dist\Electroproject.exe
pause
exit /b 0

:err_python
echo Error: Python was not found (python/py).
pause
exit /b 1

:err_main
echo Error: main.py is missing in project folder.
pause
exit /b 1

:err_pip
echo Error: failed to install/update PyInstaller.
pause
exit /b 1

:err_build
echo Error: EXE build failed.
pause
exit /b 1
