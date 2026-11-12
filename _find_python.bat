@echo off
rem ---------------------------------------------------------------------------
rem  Locate Python 3 and export both builds. Nothing is hardcoded, so the
rem  folder keeps working after cloning to another machine.
rem
rem  Usage:  call "%~dp0_find_python.bat"
rem  Exports: PYEXE / PYARG    console build  (python)
rem           PYWEXE / PYWARG  windowed build (pythonw, no console window)
rem ---------------------------------------------------------------------------
set "PYEXE="
set "PYARG="
set "PYWEXE="
set "PYWARG="

rem 1) Python launcher - follows whatever "py -3" resolves to
where py >nul 2>nul
if not errorlevel 1 (
    set "PYEXE=py"
    set "PYARG=-3"
)

rem 2) python.exe on PATH
if not defined PYEXE (
    where python >nul 2>nul
    if not errorlevel 1 set "PYEXE=python"
)

rem 3) per-user installs, newest version first
if not defined PYEXE (
    for %%v in (314 313 312 311 310 39) do (
        if not defined PYEXE (
            if exist "%LOCALAPPDATA%\Programs\Python\Python%%v\python.exe" (
                set "PYEXE=%LOCALAPPDATA%\Programs\Python\Python%%v\python.exe"
            )
        )
    )
)

rem 4) machine-wide installs
if not defined PYEXE (
    for %%v in (313 312 311 310 39) do (
        if not defined PYEXE (
            if exist "C:\Program Files\Python%%v\python.exe" (
                set "PYEXE=C:\Program Files\Python%%v\python.exe"
            )
        )
    )
)

if not defined PYEXE (
    echo.
    echo [!] Python 3 was not found.
    echo     Install it from https://www.python.org/downloads/
    echo     and tick "Add python.exe to PATH" during setup.
    exit /b 0
)

rem ---- derive the windowed build from the console build ----
if /i "%PYEXE%"=="py" (
    set "PYWEXE=pyw"
    set "PYWARG=-3"
    exit /b 0
)
if /i "%PYEXE%"=="python" (
    set "PYWEXE=pythonw"
    exit /b 0
)
for %%p in ("%PYEXE%") do set "PYWEXE=%%~dppythonw.exe"
if not exist "%PYWEXE%" set "PYWEXE=%PYEXE%"
exit /b 0
