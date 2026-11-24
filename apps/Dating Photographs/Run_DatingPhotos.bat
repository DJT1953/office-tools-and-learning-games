@echo off
rem Starts Dating Photographs, which reads the dates written on old photos and renames them by date.
cd /d "%~dp0"
start "" pyw dating_photos.py
