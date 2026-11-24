@echo off
rem Starts the Photo Stitcher, which joins overlapping photos into one big picture.
cd /d "%~dp0"
call :need_python "Photo Stitcher"
if errorlevel 1 exit /b 1
start "" pyw photo_stitcher.py
exit /b

rem ---- The same Python check is in every tool's start file ----
:need_python
py -c "import sys; sys.exit(sys.version_info < (3, 10))" >nul 2>nul
if not errorlevel 1 exit /b 0
title %~1 needs Python
echo.
where py >nul 2>nul
if not errorlevel 1 goto add_python
echo  %~1 needs Python, which is free. This PC doesn't have it yet.
where winget >nul 2>nul
if errorlevel 1 goto python_by_hand
choice /c YN /m "  Install Python now (from python.org, through Windows' own installer)"
if errorlevel 2 goto python_declined
echo  If Windows asks you to agree to its terms, type Y and press Enter.
winget install --id Python.PythonInstallManager -e --source winget
if errorlevel 1 goto python_by_hand
where py >nul 2>nul
if errorlevel 1 goto python_reopen
:add_python
echo  Adding Python 3.14, the version %~1 is tested with...
py install 3.14
py -c "import sys; sys.exit(sys.version_info < (3, 10))" >nul 2>nul
if not errorlevel 1 exit /b 0
:python_by_hand
echo.
echo  Please install Python 3.10 or newer from https://www.python.org/downloads/
echo  then double-click this file again.
pause
exit /b 1
:python_declined
echo  Nothing was installed. %~1 can't start without Python.
pause
exit /b 1
:python_reopen
echo  Python is installed. Double-click this file again to start %~1.
pause
exit /b 1
