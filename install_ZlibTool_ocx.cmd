@prompt -$G
regsvr32 /u /s ZlibTool.ocx

@echo Uninstall succeed-DownFeed.net. 
@echo Press any key to install or ctrl+c to cancel-DownFeed.net.
@pause >nul

regsvr32 ZlibTool.ocx

@if errorlevel==1 (
   
   @echo Regsvr32 returns Error: %errorlevel% !  Calling rundll32 that may show more details about the problem:
   rundll32 ZlibTool.ocx,DllRegisterServer
   
   @pause >nul
)
