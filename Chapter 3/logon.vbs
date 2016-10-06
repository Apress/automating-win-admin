'logon.vbs
Const DOMAIN = "Acme"

Set objShell = CreateObject("WScript.Shell")
Set objNetwork = CreateObject("WScript.Network") 
Set objFSO = CreateObject("Scripting.FileSystemObject")

'create a ENTWSH.HTMLGen component 
Set objHTMLGen = GetObject("script:\\odin\wsc$\htmlgen.wsc")

On Error Resume Next 
objHTMLGen.StartDOC "ACME Ltd. Logon Script", True

'get logged on user name to display in greeting. Loop to ensure 
' user ID is returned correctly on Win 9x/ME computers
Do While strUser =""
 strUser = objNetwork.UserName
 WScript.Sleep 100
Loop

'get ADSI User object from domain
Set objUser = GetObject("WinNT://" & DOMAIN & "/" & strUser & ",User")
'if no error occurred getting ADSI user object, set user ID to
'user full name
If Not Err Then strUser = objUser.FullName

Set objIE = objHTMLGen.Object
objIE.MenuBar = 0
objIE.Toolbar = 0

strMsg = "Good "

If Hour(Time)>=17 Then 
  strMsg = strMsg & "Evening, "
ElseIf Hour(Time)>=12 Then 
  strMsg = strMsg & "Afternoon, "
Else
  strMsg = strMsg & "Morning, "
End If

Err.Clear
'if OS Windows 9x/ME map home directory
If Not objShell.Environment("OS") = "Windows_NT" Then  
 objNetwork.RemoveNetworkDrive "H:", True
 Err.Clear
 objNetwork.MapNetworkDrive "H:", "\\Odin\" & objNetwork.UserName _
                            & "$", True
End If 

objHTMLGen.WriteLine "<center><b>" & strMsg & strUser & _
                     "</center></b><br>"
Set objTS = objFSO.OpenTextFile("\\odin\d$\data\messages\hello.htm")
objHTMLGen.WriteLine objTS.ReadAll
objHTMLGen.WriteLine "<center><input type=""submit"" value=""" & _ 
                  "Close Window"" onClick=""window.close()""></center>"

objHTMLGen.EndDOC
