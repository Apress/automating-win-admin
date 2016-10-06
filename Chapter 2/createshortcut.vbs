'create a shortcut on desktop that starts Notepad with Autoexec.bat 
'as parameter
Dim objShell, strDesktop, objShortCut
'create a shortcut on desktop 
Set objShell = CreateObject("WScript.Shell")
strDesktop = objShell.SpecialFolders("Desktop") 'get the path to desktop
Set objShortCut = objShell.CreateShortCut(strDesktop & "\Show AutoExec.lnk")
objShortCut.IconLocation = "c:\winnt\system32\SHELL32.dll,9"
objShortCut.TargetPath = "notepad.exe" 'script to execute
objShortCut.Arguments = "c:\autoexec.bat" 'argument to pass
objShortCut.HotKey =  "ALT+CTRL+N" 'hotkey to start 
objShortCut.Save ' save and update shortcut
