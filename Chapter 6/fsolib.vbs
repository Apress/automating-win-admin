'fsolib.vbs
'Description: Contains routines used by FSO scripts

'check if script is being run interactively
'Returns:True if run from command line, otherwise false
Function IsCscript()
  IsCScript = (StrComp(Right(Wscript.Fullname,11),"cscript.exe",1) = 0)
End Function

'display an error message and exist script
'Parameters:
'strMsg        Message to display
'strUseWscript Use Wscript.Echo to display message. 
'   By default StdErr is used, but this cannot be used in 
'   interactive (wscript) mode unless redirected to somewhere else.
Sub ExitScript(strMsg, bUseWscript)
 If bUseWscript Then
  Wscript.Echo strMsg 
 Else
  'get the standard error stream
   Wscript.StdErr.WriteLine strMsg
 End If 
 Wscript.Quit -1
End Sub

'returns contents of specified file. If file doesn't exist
'terminates script and displays error message
'Parameters:
'strFile Path to file to return
'Returns
'contents of specified file
Function GetFile(strFile)
    On Error Resume Next
    Dim objFSO, objFile
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFile = objFSO.OpenTextFile(strFile)
    If Err Then ExitScript _ 
        "Error " & Err.Description & " opening file " & _
        strFile, False
    GetFile = objFile.ReadAll
    objFile.Close
End Function

'terminates script with message if script not run using cscript.ext
'Parameters:None
Sub CheckCScript()
    If Not IsCscript Then ExitScript _
        "This script must be run from command line using cscript.exe", True
End Sub

'checks if specified number of arguments have been passed and exits script
'displaying usage information if not
'Parameters:
'nCount  Number of arguments expected
Sub CheckArguments(nCount)
    If WScript.Arguments.Count <> nCount Then
        WScript.Arguments.ShowUsage
        WScript.Quit
    End If
End Sub
