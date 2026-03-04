@echo off
setlocal EnableExtensions

set "ROOT=%~dp0.."
for %%I in ("%ROOT%") do set "ROOT=%%~fI"

set "XLSX=%ROOT%\fixtures\table_M2.xlsx"
set "OUT_TBL=%ROOT%\out\table_M2.tbl"
set "LAYOUT_CFG=%ROOT%\config\table_layout.ini"

if not "%~1"=="" set "XLSX=%~1"
if not "%~2"=="" set "OUT_TBL=%~2"
if not "%~3"=="" set "LAYOUT_CFG=%~3"

for %%I in ("%OUT_TBL%") do set "OUT_DIR=%%~dpI"
if not exist "%OUT_DIR%" (
  mkdir "%OUT_DIR%"
)

echo INFO: Input XLSX: "%XLSX%"
echo INFO: Output TBL: "%OUT_TBL%"
echo INFO: Layout config: "%LAYOUT_CFG%"
echo INFO: Running cscript...

cscript //nologo "%ROOT%\src\create_tbl.vbs" "%XLSX%" "%OUT_TBL%" "%LAYOUT_CFG%"
set "RC=%ERRORLEVEL%"

echo INFO: Return code: %RC%
exit /b %RC%
