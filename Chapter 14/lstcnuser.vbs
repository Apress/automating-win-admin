'lstcnusers.vbs
'list connected users to specified server
Dim objFileService, objSession, strComputer 

'check argument count
If Not Wscript.Arguments.Count = 1 Then
 ShowUsage
 Wscript.Quit
End If

'get the file service object
strComputer = Wscript.Arguments(0) 
'get a reference to the LanmanServer service
Set objFileService = GetObject("WinNT://" & strComputer & "/LanmanServer")

'loop through each session and display any connected users
For Each objSession In objFileService.sessions
 'check if the session user ID is not empty
 If Not objSession.user = "" Then 
     Wscript.StdOut.WriteLine objSession.user 
 End If

Next

Sub ShowUsage
    WScript.Echo "lstcnusers.vbs lists connected users " & vbCrLf & _ 
    "Syntax:"  &  vbCrLf & _
    "lstcnusers.vbs computer" & vbCrLf & _
    "computer computer to enumerate connected users" 
End Sub
