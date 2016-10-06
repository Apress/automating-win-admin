'notepad.vbs
Const WshNormalFocus = 1
Dim objShell

Set objShell = WScript.CreateObject("WScript.shell")
objShell.Run "Notepad.exe", WshNormalFocus ,True
