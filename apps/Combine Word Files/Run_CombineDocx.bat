@echo off
cd /d "%~dp0"
py -m pip install --quiet python-docx docxcompose
py combine_docx.py
if errorlevel 1 pause
