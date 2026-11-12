@echo off
rem Starts the Proofreader. Documents or folders dragged onto this file go into its list.
cd /d "%~dp0"
start "" pyw proofreader.py %*
