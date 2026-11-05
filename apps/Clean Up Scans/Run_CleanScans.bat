@echo off
cd /d "%~dp0"
py -m pip install --quiet numpy opencv-python-headless pillow pillow-heif
py clean_scans.py
if errorlevel 1 pause
