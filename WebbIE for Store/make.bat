@echo off
rem This batch file converts a Win32 application into an AppX for the Windows 10 Store.
set appname="WebbIE"
rem "C:\Program Files (x86)\Windows Kits\10\bin\x64\makeappx.exe" pack /f mappings.ini /p %appname%.appx
"C:\Program Files (x86)\Windows Kits\10\bin\x64\makeappx.exe" pack /d App /p %appname%.appx /o
rem We can now sign it with a developer licence key for testing or shipping to a site licencee, or (better) submit it to the Windows Store.
pause