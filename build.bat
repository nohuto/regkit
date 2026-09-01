@echo off
cd /d "%~dp0"
if /I "%~1"=="x86" goto :x86
cmake --preset x64-release || exit /b 1
cmake --build --preset x64-release || exit /b 1
if /I "%~1"=="x64" exit /b 0
:x86
cmake --preset x86-release || exit /b 1
cmake --build --preset x86-release || exit /b 1
