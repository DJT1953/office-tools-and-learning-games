@echo off
rem Starts the Video Player. Videos or folders dragged onto this file open in it.
cd /d "%~dp0"
start "" pyw video_player.py %*
