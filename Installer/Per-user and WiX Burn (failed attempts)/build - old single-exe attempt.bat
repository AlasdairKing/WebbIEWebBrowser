@title Windows Installer using WiX
@echo Opening WiX file for version.
@call notepad Webbie4.wxs
@echo Copying files
@copy ..\WebbIE4.NET\bin\Release\WebbIE4.exe SourceDir\WebbIE4.exe
@copy ..\WebbIE4.NET\bin\Release\Common.Language.xml SourceDir\Common.Language.xml
@copy ..\WebbIE4.NET\bin\Release\WebbIE4.exe.config SourceDir\WebbIE4.exe.config
@copy ..\WebbIE4.NET\bin\Release\WebbIE4.Help-en.rtf SourceDir\WebbIE4.Help-en.rtf
@copy ..\WebbIE4.NET\bin\Release\WebbIE4.Language.xml SourceDir\WebbIE4.Language.xml
@copy ..\WebbIE4.NET\bin\Release\WebbIE4.ico SourceDir\WebbIE4.ico
@copy ..\WebbIE4.NET\bin\Release\WebbIE4htmldoc.ico SourceDir\WebbIE4htmldoc.ico
@echo Signing executable
@signtool.exe sign /sha1 605A4D6DDF8DBB97FE42475C8600A5E23B6C6230  SourceDir\WebbIE4.exe
@path=%path%;C:\Program Files (x86)\WiX Toolset v3.7\bin
@if exist webbie4.wixobj del WebbIE4.wixobj
@candle WebbIE4.wxs
@if exist webbie4.msi del WebbIE4.msi 
@light WebbIE4.wixobj -ext WixUIExtension -spdb -sice:ICE91 -ext WixUtilExtension -ext WixNetfxExtension
@signtool.exe sign /sha1 605A4D6DDF8DBB97FE42475C8600A5E23B6C6230  WebbIE4.msi
@echo Copying installer to the setup.exe project.
@if exist ..\setup\WebbIE4.msi del ..\setup\WebbIE4.msi
@copy WebbIE4.msi ..\setup\WebbIE4.msi
@if exist webbie4.msi del WebbIE4.msi
@echo Copying files to the local FTP folder.
@if exist C:\users\alasdair\documents\personal\aljo\webbie\webbrowser\setup.exe del C:\users\alasdair\documents\personal\aljo\webbie\webbrowser\setup.exe
@if exist C:\users\alasdair\documents\personal\aljo\webbie\webbrowser\WebbIE4.msi del C:\users\alasdair\documents\personal\aljo\webbie\webbrowser\WebbIE4.msi
@echo Copying setup.exe
@copy setup.exe C:\users\alasdair\documents\personal\aljo\webbie\webbrowser\setup.exe /Y
@echo Copying WebbIE4.msi
@copy WebbIE4.msi C:\users\alasdair\documents\personal\aljo\webbie\webbrowser\WebbIE4.msi /Y
@if exist WebbIE4.msi del WebbIE4.msi
@echo Update updates.xml
@call notepad C:\users\alasdair\documents\personal\aljo\webbie\webbrowser\updates.xml
@echo All done. Files are in the web folder ready for upload. 
@pause

