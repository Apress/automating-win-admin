'chkmember.vbs
Const Domain = "ACME" 
Dim strUser, objGroup, strGroup, objUser

Set objNetwork = CreateObject("WScript.Network")

'get name of logged on user - loop until user id is retrieved, required for Win9x
Do While strUser = ""
  strUser = objNetwork.UserName
  WScript.Sleep 100
Loop

If CheckGroup(Domain, "Accounting", strUser) Then
   'connect to accounting share
   objWshNetwork.MapNetworkDrive "P:", "\\THOR\Accounting", True
End If

Function CheckGroup(strDomain, strGroup, strUser)
Dim objGroup
'get ADSI User object from domain
Set objGroup = GetObject("WinNT://" & strDomain & "/" _
               & strGroup & ",Group")

If objGroup.IsMember("WinNT://" & strDomain & "/" & strUser) Then
  CheckGroup = True
Else
  CheckGroup = False
End If

End Function
