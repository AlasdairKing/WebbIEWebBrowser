@echo off
title Windows Installer using WiX

set SIGNTOOL=C:\Program Files\Microsoft SDKs\Windows\v7.1\Bin\signtool.exe

echo Signing EXE
if exist "%SIGNTOOL%" (
	rem OK
	) else ( 
	echo Signtool.exe is not where the batch file thinks it is. You'll have to edit the batch file where it says SIGNTOOL=. SIGNTOOL.EXE is in the Windows SDK.
	goto:END
)
:sign
"%SIGNTOOL%" sign /sha1 890e6445bfa2a4c9566975680b3a7e66ff48fed0 "..\WebbIE4.NET\bin\Release\WebbIE4.exe"
"%SIGNTOOL%" sign /sha1 890e6445bfa2a4c9566975680b3a7e66ff48fed0 "..\WebbIE4.NET\bin\Release\WebbIEUpdater.dll"
if ERRORLEVEL 0 goto:SIGNEDOK
echo Failed to sign the executable. Probably the WebbIE signing certificate is not installed on your machine, which
echo means you're not Alasdair. Delete the signtool line from the build.bat batch script.
goto:END
:SIGNEDOK
if exist webbie4.wixobj del WebbIE4.wixobj
echo Candle
"%WIX%\bin\candle" WebbIE4.wxs -nologo -ext WixNetfxExtension -ext WixUtilExtension -ext WixUiExtension 
echo Light
"%WIX%\bin\light" WebbIE4.wixobj -spdb -sice:ICE91 -nologo -ext WixNetfxExtension -ext WixUtilExtension -ext WixUiExtension
echo Sign with my licence key. Better than nothing!
"%SIGNTOOL%" sign /sha1 890e6445bfa2a4c9566975680b3a7e66ff48fed0 WebbIE4.msi
echo All done!
:END
@pause

