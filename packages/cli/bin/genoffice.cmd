@echo off
rem genoffice launcher for the packaged Windows app: <install>\resources\cli\genoffice.cmd
setlocal
set ELECTRON_RUN_AS_NODE=1
"%~dp0..\..\GenOffice.exe" "%~dp0genoffice.cjs" %*
endlocal
