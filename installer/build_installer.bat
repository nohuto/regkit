@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
set "ISCC=%userprofile%\AppData\Local\Programs\Inno Setup 6\ISCC.exe"

"%ISCC%" "%SCRIPT_DIR%regkit.iss"

endlocal
