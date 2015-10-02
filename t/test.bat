@echo off

echo --REMEMBER-- Edit 'orig.log' 'FILE = ' statements with current directory path

del out.log 2>nul

 > commands.log echo ..\pests
>> commands.log echo ..\pests -d
>> commands.log echo ..\pests -c 1:1;1:2;1:3;2:2;2:3;2:4
>> commands.log echo ..\pests -d -c 1:1;1:2;1:3;2:2;2:3;2:4

>> commands.log echo ..\pests -t txt,csv
>> commands.log echo ..\pests -t txt,csv -d -W out.xls

>> commands.log echo ..\pests -d -W out.xls
>> commands.log echo ..\pests -d -R -W out.xls
>> commands.log echo ..\pests -d -S -W out.xls
>> commands.log echo ..\pests -d -c 1:1;1:2;1:3;2:2;2:3;2:4 -W out.xls
>> commands.log echo ..\pests -d -c 1:1;1:2;1:3;2:2;2:3;2:4 -R -W out.xls
>> commands.log echo ..\pests -d -c 1:1;1:2;1:3;2:2;2:3;2:4 -S -W out.xls
>> commands.log echo ..\pests -d -c 1:1;1:2;1:3;2:2;2:3;2:4 -C 10:10 -W out.xls
>> commands.log echo ..\pests -d -c 1:1;1:2;1:3;2:2;2:3;2:4 -C 10:10 -R -W out.xls
>> commands.log echo ..\pests -d -c 1:1;1:2;1:3;2:2;2:3;2:4 -C 10:10 -S -W out.xls
>> commands.log echo ..\pests -S
>> commands.log echo ..\pests -S -t txt,xls,csv
>> commands.log echo ..\pests -r Book1.xls
>> commands.log echo ..\pests -F Book1.xls

REM ERRORS
>> commands.log echo ..\pests -c 1;1
>> commands.log echo ..\pests -C 1:1
>> commands.log echo ..\pests -C 1!1:1 -W out.xls
>> commands.log echo ..\pests -c 1:1;1:2;1:3 -C 1:1;1:2 -W out.xls
>> commands.log echo ..\pests -t txt,csv -s 2 -d -W out.xls

for /f "tokens=*" %%i in (commands.log) do @echo C: %%i >> out.log && %%i >> out.log && echo. >> out.log

fc /n /t /l orig.log out.log > nul
if %ERRORLEVEL%==0 (echo OK) else fc /n /t /l orig.log out.log

del out.log 2>nul
del commands.log 2>nul
