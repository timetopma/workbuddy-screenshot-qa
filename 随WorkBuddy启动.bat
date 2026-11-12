@echo off
chcp 65001 >nul
rem One-time setup: register a per-user autostart entry + the self-healing task.
rem After this the helper stays resident and the button appears whenever WorkBuddy opens.
call "%~dp0_find_python.bat"
if not defined PYEXE goto :nopython
"%PYEXE%" %PYARG% "%~dp0screenshot_qa.py" --install-autostart
echo.
pause
exit /b 0

:nopython
echo.
pause
exit /b 1
