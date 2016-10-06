REM Logon.bat
REM Checks if WSH is installed on client and attempts to install it
REM then executes WSH logon script.
@ECHO OFF

IF "%OS%" == "Windows_NT" goto WIN_NT

IF NOT EXIST %WINDIR%\CSCRIPT.EXE %0\..\WSHBIN\STE50EN.EXE /Q
GOTO ENDSCRIPT
:WIN_NT
IF NOT EXIST %WINDIR%\SYSTEM32\CSCRIPT.EXE %0\..\WSHBIN\STE50EN.EXE /Q
GOTO ENDSCRIPT

:ENDSCRIPT
REM execute WSH script
 cscript login.vbs
