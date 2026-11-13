@echo off
rem Starts the Song Cleaner.
cd /d "%~dp0"
start "" pyw song_cleaner.py %*
