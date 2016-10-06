Const WMICONST = "winmgmts:{impersonationLevel=impersonate}!\\odin\root\default:"

Const HKEY_CURRENT_USER = &H80000001
Dim objRegistry, nRet, strValue

'create an instance of the StdRegProv registry provider
Set objRegistry = GetObject(WMICONST & "StdRegProv")

'read the value Identifier under key HARDWARE\DESCRIPTION\System
'since the local machine root key is the default, the first parameter is omitted
nRet = objRegistry.GetStringValue(, "HARDWARE\DESCRIPTION\System", _
                                  "Identifier", strValue)

If nRet = 0 Then
    Wscript.Echo "The system type is " & strValue
Else
    Wscript.Echo "Error reading the HARDWARE\DESCRIPTION\System key"
End If
