@title WebbIE Package
@echo Building WebbIE Package
@echo.
@echo.
@echo Candle
@"C:\Program Files (x86)\WiX Toolset v3.7\bin\candle.exe" WebbIE4.wxs -ext WixBalExtension -ext WixNetFxExtension 
@echo.
@echo.
@echo Light
@"C:\Program Files (x86)\WiX Toolset v3.7\bin\light.exe" WebbIE4.wixobj -ext WixBalExtension -ext WixNetFxExtension 
@pause