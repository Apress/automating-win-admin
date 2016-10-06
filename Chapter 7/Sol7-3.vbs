Set objShell = CreateObject("Wscript.Shell")

'delete a registry value
objShell.RegDelete "HKEY_CURRENT_USER\Software\MyApp\Config\coords"
