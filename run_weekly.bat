@echo off
rem Run the weekly deck script from this batch file's folder.
cd /d "%~dp0"
rem Open Edge visibly so Looker sign-in and chart captures can be monitored.
py "%~dp0\smt_weekly_deck.py" --build --headed
