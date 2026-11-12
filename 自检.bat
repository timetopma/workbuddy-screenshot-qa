@echo off
chcp 65001 >nul
rem Environment self-check: Python / Pillow / tkinter / screen / WorkBuddy window.
call "%~dp0_find_python.bat"
if not defined PYEXE goto :nopython
"%PYEXE%" %PYARG% "%~dp0screenshot_qa.py" --selftest
echo.
pause
exit /b 0

:nopython
echo.
pause
exit /b 1
