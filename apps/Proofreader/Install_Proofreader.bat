@echo off
rem One-time setup for the Proofreader: installs everything it needs except
rem Microsoft Word (which is optional). Safe to run again.
setlocal
title Install the Proofreader
cd /d "%~dp0"

echo.
echo  PROOFREADER SETUP
echo  =================
echo  This installs what the Proofreader needs:
echo    - Python (only if this computer does not have it)
echo    - the Python add-ons PySide6, comtypes and python-docx
echo    - the free checker LanguageTool with its own Java (about 300 MB)
echo    - a shortcut on the Desktop
echo  Microsoft Word is NOT installed. The Proofreader works without it.
echo.

where py >nul 2>nul
if not errorlevel 1 goto have_py

echo  Python was not found on this computer.
where winget >nul 2>nul
if errorlevel 1 goto no_winget
choice /c YN /m "  Install Python now (free, from python.org through Windows' own installer)"
if errorlevel 2 goto declined
winget install --id Python.PythonInstallManager -e --source winget
if errorlevel 1 goto python_failed
where py >nul 2>nul
if errorlevel 1 goto reopen

:have_py
py -c "import sys; sys.exit(0 if sys.version_info >= (3, 10) else 1)" >nul 2>nul
if not errorlevel 1 goto run_setup
echo  Adding a current version of Python...
py install 3.14
py -c "import sys; sys.exit(0 if sys.version_info >= (3, 10) else 1)" >nul 2>nul
if errorlevel 1 goto old_python

:run_setup
py proofreader_setup.py
echo.
pause
exit /b

:no_winget
echo  Please install Python 3.10 or newer from https://www.python.org/downloads/
echo  then run this file again.
goto stop
:declined
echo  Nothing was installed. Python is needed to run the Proofreader.
goto stop
:python_failed
echo  Python could not be installed automatically. Install it from
echo  https://www.python.org/downloads/ then run this file again.
goto stop
:reopen
echo  Python is installed. Close this window and run this file once more
echo  so Windows can find it.
goto stop
:old_python
echo  This computer's Python is older than 3.10. Install a current one from
echo  https://www.python.org/downloads/ then run this file again.
:stop
echo.
pause
