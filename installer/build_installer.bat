@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
set "ISCC=C:\Program Files (x86)\Inno Setup 6\ISCC.exe"

"%ISCC%" "%SCRIPT_DIR%regkit.iss"

endlocal
