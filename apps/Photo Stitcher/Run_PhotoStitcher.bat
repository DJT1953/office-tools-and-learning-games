@echo off
rem Starts the Photo Stitcher, which joins overlapping photos into one big picture.
cd /d "%~dp0"
start "" pyw photo_stitcher.py
