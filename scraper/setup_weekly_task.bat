@echo off
REM Run this file ONCE (double-click it) to schedule the news scraper to run
REM automatically every Sunday at 8:00 AM via Windows Task Scheduler.
REM Safe to double-click again later if you ever want to re-create/reset the task.

schtasks /create /tn "Enmaeya Weekly News Scraper" /tr "\"%~dp0run_weekly.bat\"" /sc weekly /d SUN /st 08:00 /f

echo.
echo Done. "Enmaeya Weekly News Scraper" is now scheduled to run every Sunday at 8:00 AM.
echo It will scrape all ~45 sources and add only newly published articles to the
echo .xlsx files in the News folder (nothing gets duplicated).
echo.
echo Notes:
echo  - Your computer needs to be turned on (and logged in) at that time for it to run.
echo  - To check/change it later: open "Task Scheduler" and look under Task Scheduler Library
echo    for "Enmaeya Weekly News Scraper".
echo  - Output/errors from each run are appended to weekly_run_output.txt in this folder,
echo    and the detailed per-site log is in scrape_log.txt as always.
pause
