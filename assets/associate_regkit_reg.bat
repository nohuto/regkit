@echo off
setlocal

set "regkit_exe=%~1"
if "%regkit_exe%"=="" set "regkit_exe=%~dp0..\\regkit.exe"

for %%I in ("%regkit_exe%") do set "regkit_exe=%%~fI"

set "base=HKCU\Software\Classes\SystemFileAssociations\.reg\shell\EditWithRegKit"
reg add "%base%" /ve /d "Edit with RegKit" /f >nul
reg add "%base%" /v "Icon" /d "\"%regkit_exe%\",0" /f >nul
reg add "%base%\command" /ve /d "\"%regkit_exe%\" \"%%1\"" /f >nul

endlocal