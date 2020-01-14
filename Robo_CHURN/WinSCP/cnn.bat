@echo off

option batch abort
option confirm off

set SESSION=ftps://GCASH_FTP:2Uku*?gu48th#Vut@ftp.telefonicabpo.com:990/
set REMOTE_PATH=/GCASH
 
echo open %SESSION% >> script.tmp

echo cd %REMOTE_PATH% 
 
rem Generate "put" command for each line in list file
for /F %%i in (list.txt) do echo put "%%i" "%REMOTE_PATH%" >> script.tmp
 
echo ls >> script.tmp
 
winscp.com /script=script.tmp
set RESULT=%ERRORLEVEL%

pause 
rem del script.tmp
 
rem Propagating WinSCP exit code
exit /b %RESULT%