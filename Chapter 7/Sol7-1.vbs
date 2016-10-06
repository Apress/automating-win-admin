Set objShell = CreateObject("Wscript.Shell")
Wscript.Echo "Your wallpaper Is " & _
    objshell.RegRead("HKCU \Control Panel\Desktop\Wallpaper")
