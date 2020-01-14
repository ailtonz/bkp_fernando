@echo off
set SESSION=ftps://80397306:41L70N*()@10.238.3.56/
set REMOTE_PATH=/Ciclo de receita/_AILTON
 
rem echo open %SESSION% >> script.tmp
 
rem Generate "put" command for each line in list file
for /F %%i in (list.txt) do echo put "%%i" "%REMOTE_PATH%" >> script.tmp
 
echo ls >> script.tmp
 
winscp.com /script=script.tmp
set RESULT=%ERRORLEVEL%
 
rem del script.tmp
 
rem Propagating WinSCP exit code
exit /b %RESULT%