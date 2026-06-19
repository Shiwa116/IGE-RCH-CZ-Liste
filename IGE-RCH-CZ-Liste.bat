@echo off
python "%~dp0build_reverse_charge.py"
if errorlevel 1 (
    echo.
    echo *** Fehler aufgetreten – bitte Meldung oben lesen ***
)
pause
