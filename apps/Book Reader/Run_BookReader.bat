@echo off
rem Starts the Book Reader. Books dragged onto this file open in it.
cd /d "%~dp0"
start "" pyw book_reader.py %*
