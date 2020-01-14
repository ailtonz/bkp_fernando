SET Today=%Date: =0%
SET Year=%Today:~-4%
SET Month=%Today:~-7,2%
SET Day=%Today:~-10,2%

set hr=%TIME: =0%
set hr=%hr:~0,2%
set min=%TIME:~3,2%

for /d %%X in (OlxAutomation) do "C:\app\7-ZipPortable\App\7-Zip64\7z.exe" a %Year%%Month%%Day%-%hr%%min%_"%%X.7z" "%%X\"

pause