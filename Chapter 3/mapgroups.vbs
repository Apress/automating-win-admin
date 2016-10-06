Dim strGroups, objGroup, objUser, strUser, objNetwork
Const Domain = "acme"

Set objNetwork = CreateObject("WScript.Network")
'get logged on user name
'get name of logged on user - loop until user id is retrieved, required for Win9x
Do While strUser = ""
 strUser = objNetwork.UserName
 WScript.Sleep 100
Loop

'get ADSI User object from domain
Set objUser = GetObject("WinNT://" & Domain & "/" & strUser & ",User")

strGroups = ";"
'enumerate all user groups
For Each objGroup In objUser.Groups
  strGroups = strGroups & objGroup.Name & ";"
Next

'check if member of Accounting or Finance group 
MapGroup "Accounting", "\\THOR\Accounting", "P:"
MapGroup "Finance", "\\THOR\Finance", "O:"

Sub MapGroup(strName, strShare, strDrive)
'check if member of specified group
   If InStr(strGroups, ";" & strName & ";", vbTextCompare) > 0 Then
    objNetwork.MapNetworkDrive strDrive,  strShare, True
   End If
End Sub
