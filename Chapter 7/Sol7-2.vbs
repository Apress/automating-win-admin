'change the default document directory for Word 97
Set objShell = CreateObject("Wscript.Shell")
objShell.RegWrite _
" HKEY_CURRENT_USER\Software\MyApp\Config\username" _
, "H:\Data\Word"
'create new registry key
objShell.RegWrite _
     "HKEY_CURRENT_USER\Software\MyApp\Config\NewKey\" , ""
