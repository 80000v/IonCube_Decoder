# IonCube_Decoder
IonCube_Decoder--------IonCube decryption tool


ChangeLog

11.02.20010
   Made decoder independent from licensefile

07.02.20010
   Added 'Decode License File' form

15.01.2010
   Recompiled ZlibTool.ocx and make it to use MFC42.dll msvcrt.dll(instead of MFC71.dll msvcr71.dll) that are included within the windows installation - thus are present on every system.
   Removed InitCommonControls() VB6 workaround and integrate the bugfix into ZlibTool.ocx

4.12.2009
   added *.Txt dumper
   improved bytecode log output
   add stack error detecting & correction function 

2.12.2009
fixed ZlibTool.ocx bug to came up on compiled exe

1.12.2009
Started with decompiler part

15.11.2009
fixed bug in HandleServerRestrictions
started with php5 support
