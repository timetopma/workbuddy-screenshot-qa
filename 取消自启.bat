@echo off
chcp 65001 >nul
rem Remove the per-user autostart entry and the self-healing scheduled task.
call "%~dp0_find_python.bat"
if not defined PYEXE goto :nopython
"%PYEXE%" %PYARG% "%~dp0screenshot_qa.py" --remove-autostart
echo.
pause
exit /b 0

:nopython
echo.
pause
exit /b 1
