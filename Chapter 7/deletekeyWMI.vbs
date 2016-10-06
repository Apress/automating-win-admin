Const WMICONST = "winmgmts:{impersonationLevel=impersonate}!root\default:"
Const HKEY_CURRENT_USER = &H80000001
Dim strPath, nRet, objRegistry, strKeyPath

'set the key path
strKeyPath = "software\test"

'create an instance of the registry provider
Set objRegistry = GetObject(WMICONST & "StdRegProv")
'delete a key called 'subkey'
nRet = objRegistry.DeleteKey(HKEY_CURRENT_USER, "testkey\subkey")
