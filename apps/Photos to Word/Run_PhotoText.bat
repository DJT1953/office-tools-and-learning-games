@echo off
cd /d "%~dp0"
py -m pip install --quiet pillow pytesseract python-docx pillow-heif
py photos_to_docx.py
if errorlevel 1 pause
