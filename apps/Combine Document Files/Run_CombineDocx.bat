@echo off
rem Starts Combine Document Files. A folder or documents dropped onto this file open in it.
rem It offers to add anything it still needs the first time.
cd /d "%~dp0"
start "" pyw combine_docx.py %*
