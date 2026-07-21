@echo off
if exist "%~dp0..\python\python.exe" set "PATH=%~dp0..\python;%PATH%"
python "%~dp0build_reverse_charge.py"
if errorlevel 1 (
    echo.
    echo *** Fehler aufgetreten – bitte Meldung oben lesen ***
)
pause
