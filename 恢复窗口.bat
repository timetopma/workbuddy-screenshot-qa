@echo off
chcp 65001 >nul
rem Emergency recovery: bring the WorkBuddy window back if it stayed hidden.
call "%~dp0_find_python.bat"
if not defined PYEXE goto :nopython
"%PYEXE%" %PYARG% "%~dp0screenshot_qa.py" --restore
echo.
pause
exit /b 0

:nopython
echo.
pause
exit /b 1
