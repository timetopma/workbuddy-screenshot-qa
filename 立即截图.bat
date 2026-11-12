@echo off
rem Run one capture pass, then exit. Useful for testing the pipeline.
call "%~dp0_find_python.bat"
if not defined PYWEXE exit /b 1
start "" "%PYWEXE%" %PYWARG% "%~dp0screenshot_qa.py" --once
