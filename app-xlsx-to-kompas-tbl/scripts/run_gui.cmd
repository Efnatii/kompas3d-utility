@echo off
setlocal EnableExtensions

set "ROOT=%~dp0.."
for %%I in ("%ROOT%") do set "ROOT=%%~fI"

set "APP_EXE=%ROOT%\bin\app-xlsx-to-kompas-tbl.exe"
if exist "%APP_EXE%" (
  "%APP_EXE%"
  set "RC=%ERRORLEVEL%"
  exit /b %RC%
)

set "LAUNCHER=%ROOT%\scripts\run_gui.vbs"
if not exist "%LAUNCHER%" (
  echo ERROR: run_gui.vbs not found: "%LAUNCHER%"
  exit /b 2
)

wscript.exe "%LAUNCHER%"
set "RC=%ERRORLEVEL%"

exit /b %RC%
