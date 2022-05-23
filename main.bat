@echo off
for /f "usebackq tokens=*" %%a in (`CSCRIPT "utils.vbs" "GetDate"`) do set num=%%a
echo %num%

REM call ws
REM utils.vbs PortiaWebService 1 1

pause