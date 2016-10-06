'regsrch.vbs
'searches for specified registry values and optionally replaces them
Const WMICONST = "winmgmts:{impersonationLevel=impersonate}!root\default:"

Dim strRoot, nPos
Dim strFind, strPath, strReplace, bReplace, objRegistry, nRoot

'check the argument count
If WScript.Arguments.Count < 2 Or WScript.Arguments.Count > 3 Then
    ShowUsage
End If

strPath = WScript.Arguments(0)
strFind = WScript.Arguments(1)

nPos = InStr(strPath, "\")

If nPos > 0 Then
    strRoot = Mid(strPath, 1, InStr(strPath, "\") - 1)
    strPath = Mid(strPath, nPos + 1)
Else
    
End If

'get the registry root key
Select Case strRoot
    Case "HKEY_LOCAL_MACHINE", "HKLM"
      nRoot = &H80000002
    Case "HKEY_CURRENT_USER", "HKCU"
      nRoot = &H80000001
    Case "HKEY_CLASSES_ROOT"
      nRoot = &H80000000
    Case "HKEY_USERS"
      nRoot = &H80000003
    Case "HKEY_CURRENT_CONFIG"
      nRoot = &H80000005
    Case "HKEY_DYN_DATA"
      nRoot = &H80000006
    Case Else
     WScript.Echo "Invalid registry root key: " & strRoot
     WScript.Quit
End Select

'create an instance of the registry provider
Set objRegistry = GetObject(WMICONST & "StdRegProv")

'check if replace parameter is passed
 If WScript.Arguments.Count = 3 Then
   strReplace = WScript.Arguments(2)
   bReplace = True
End If

RecurseReg strPath

Sub ShowUsage
WScript.Echo "regsrch search and optional replace registry values." & vbLf & _ 
    "Syntax:" &  vbLf & _
    "regsrch.vbs key findvalue [replacevalue] " &  vbCrLf & _
    "key          key to search. Will search child keys." & vbCrLf & _ 
    "findvalue    value to search for" & vbCrLf & _
    "replacevalue optional. Value to replace"
    WScript.Quit -1
End Sub

Sub RecurseReg(strRegPath)
Dim nRet, aNames, bFound
Dim strValue, aValues, nF, aRet

nRet = objRegistry.EnumKey(nRoot, strRegPath, aNames)

'loop through and enumerate any sub-keys
If nRet = 0 Then
 For nF = LBound(aNames) To UBound(aNames)
    RecurseReg strRegPath & "\" & aNames(nF)

 Next
End If

'enumerate all values
nRet = objRegistry.EnumValues(nRoot, strRegPath, aNames, aValues)

If nRet = 0 Then
 For nF = LBound(aValues) To UBound(aValues)

  strValue = ""

  'check values, only interested in string or expand string values
  Select Case aValues(nF)
   Case 1 'String
    nRet = objRegistry.GetStringValue(nRoot, strRegPath, _
                                    aNames(nF), strValue)

   Case 2 'Expanded String
    nRet = objRegistry.GetExpandedStringValue(nRoot, strRegPath, _
                                    aNames(nF), strValue)
   End Select
    
  If strValue = strFind And (aValues(nF) = 1 Or aValues(nF) = 2) Then
    WScript.Echo "Value '" & strValue & "' found in " _ 
                        & strRegPath & "\" & aNames(nF)
    
   'replace the registry value?
   If bReplace Then
    'check what type and call appropriate method
    Select Case aValues(nF)
     Case 1 'write registry string
        nRet = objRegistry.SetStringValue(nRoot, _
                             strRegPath, aNames(nF), strReplace)

     Case 2 'write expanded registry string
        nRet = objRegistry.SetExpandedStringValue(nRoot, _
                             strRegPath, aNames(nF), strReplace)
     End Select

     'was replace successful?
     If nRet = 0 Then
      WScript.Echo "Value '" & strValue & "' change to '" _
              & strReplace & "' in " & strRegPath & "\" & aNames(nF)
     Else
      WScript.Echo "Value '" & strValue & "' not replaced '" _
              & strReplace & "' in " & strRegPath & "\" & aNames(nF)
     End If
    End If
   End If
    
  Next
End If
End Sub
