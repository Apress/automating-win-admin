'fsolib.vbs
'Description: Contains routines used by FSO scripts

'check if script is being run interactively
'Returns:True if run from command line, otherwise false
Function IsCscript()

 If strcomp(Right(Wscript.Fullname,11),"cscript.exe",1)=0 Then 
   IsCscript = True
 Else
   IsCscript = False
 End If
End Function

'display an error message and exist script
'Parameters:
'strMsg        Message to display
'strUseWscript Use Wscript.Echo to display message. 
'   By default StdErr is used, but this cannot be used in 
'   interactive (wscript) mode.
Sub ExitScript(strMsg, strUseWscript)
 Dim objErr
 If strUseWscript Then
  Wscript.Echo strMsg 
 Else
  'get the standard error stream
  Set objErr = Wscript.StdErr
  objErr.Write strMsg
  objErr.Close
  Set objErr = Nothing
 End If 
 Wscript.Quit -1
End Sub