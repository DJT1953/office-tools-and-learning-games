@echo off
rem Starts the Music Player. Songs, folders or playlists dragged onto this file play in it.
cd /d "%~dp0"
start "" pyw music_player.py %*
