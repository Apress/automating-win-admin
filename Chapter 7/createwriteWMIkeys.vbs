Const WMICONST = "winmgmts:{impersonationLevel=impersonate}!root\default:"
Const HKEY_LOCAL_MACHINE = &H80000002

Dim strPath, nRet, objRegistry, strKeyPath

'set the key path to write the values to.
strKeyPath = "software\test"

'create an instance of the registry provider
Set objRegistry = GetObject(WMICONST & "StdRegProv")

'write a binary value
nRet = objRegistry.CreateKey(HKEY_LOCAL_MACHINE, strKeyPath)

'write a binary value
nRet = objRegistry.SetBinaryValue(HKEY_LOCAL_MACHINE, _
                                strKeyPath, "binary", Array(1, 2, 3, 4))
'write expanded registry string
nRet = objRegistry.SetExpandedStringValue(HKEY_LOCAL_MACHINE, _
                                     strKeyPath, "expanded", "%path%")
'write expanded string
nRet = objRegistry.SetStringValue(HKEY_LOCAL_MACHINE, _
                                          strKeyPath, "string", "heh")
'write multistring value
nRet = objRegistry.SetMultiStringValue(HKEY_LOCAL_MACHINE, strKeyPath, _
                              "multistring", Array("1", "2", "3", "4"))
