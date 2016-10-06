'enumwmi.vbs
Const WMICONST = "winmgmts:{impersonationLevel=impersonate}!root\default:"
Const HKEY_LOCAL_MACHINE  = &H80000002
    
Dim strPath, nRet, objRegistry, strKeyPath, aNames
Dim strValue, aValues, nF, nRoot, aRet

strKeyPath = "software\test"
nRoot = HKEY_LOCAL_MACHINE
'create an instance of the registry provider
Set objRegistry = GetObject(WMICONST & "StdRegProv")

'enumerate all values
nRet = objRegistry.EnumValues(, strKeyPath, aNames, aValues)

'if enumeration successful, loop through returned values and display values
If nRet = 0 Then
 For nF = LBound(aValues) To UBound(aValues)
    
  strValue = ""
  Select Case aValues(nF)
   Case 1 'String
    nRet = objRegistry.GetStringValue(nRoot, strKeyPath, _
                                    aNames(nF), strValue)
   Case 2 'Expanded String
    nRet = objRegistry.GetExpandedStringValue(nRoot, strKeyPath, _
                                    aNames(nF), strValue)

   Case 3 'Binary values
    nRet = objRegistry.GetBinaryValue(nRoot, strKeyPath, _
                                    aNames(nF), aRet)
    strValue = Join(aRet)

   Case 4 'DWORD
    nRet = objRegistry.GetDWORDValue(nRoot, strKeyPath, _
                                  , aNames(nF), strValue)
    
   Case 7 'Multi string values
    nRet = objRegistry.GetMultiStringValue(nRoot, strKeyPath, _
                                    aNames(nF), aRet)
    strValue = Join(aRet)
   End Select
    
    Wscript.Echo "Value " & aNames(nF) & " has value " & strValue
 Next
End If

Set objRegistry = Nothing
