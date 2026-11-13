@echo off
rem One-time setup for the Song Cleaner: installs everything it needs.
rem Safe to run again.
setlocal
title Install the Song Cleaner
cd /d "%~dp0"

echo.
echo  SONG CLEANER SETUP
echo  ==================
echo  This installs what the Song Cleaner needs:
echo    - Python (only if this computer does not have it)
echo    - the Python add-ons PySide6, numpy, mutagen and imageio-ffmpeg
echo      (imageio-ffmpeg brings FFmpeg, which reads and cuts the MP3s)
echo    - if you want it, the small song recogniser used to look names up online
echo    - a shortcut on the Desktop
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
py song_cleaner_setup.py
echo.
pause
exit /b

:no_winget
echo  Please install Python 3.10 or newer from https://www.python.org/downloads/
echo  then run this file again.
goto stop
:declined
echo  Nothing was installed. Python is needed to run the Song Cleaner.
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
