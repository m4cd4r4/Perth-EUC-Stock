@echo off

REM Change the command line path to C:\scratch\assets-app-current-version
cd /d C:\scratch\assets-app-current-version

REM Run the pip install command
py -m pip install -r requirements.txt

REM Pause to view any error messages or confirmation
pause