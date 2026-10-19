@echo off
REM Runs every scraper (run_all.py) once and appends output to weekly_run_output.txt.
REM Each site remembers which articles it already saved (scraper\state\*.json),
REM so every run only adds NEW articles to the .xlsx files - nothing is duplicated.
REM This file is launched automatically by the "Enmaeya Weekly News Scraper" task
REM created by setup_weekly_task.bat - you don't need to run it by hand.

cd /d "%~dp0"

where python >nul 2>nul
if %errorlevel%==0 (
    set PYCMD=python
) else (
    set PYCMD=py
)

echo ============================================== >> weekly_run_output.txt
echo Run started:  %date% %time% >> weekly_run_output.txt
%PYCMD% run_all.py --incremental --delay 2 >> weekly_run_output.txt 2>&1
echo Run finished: %date% %time% >> weekly_run_output.txt
