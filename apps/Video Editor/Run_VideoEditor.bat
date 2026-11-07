@echo off
rem Starts the Video Editor. Videos or sound files dragged onto this file open in it.
cd /d "%~dp0"
start "" pyw video_editor.py %*
