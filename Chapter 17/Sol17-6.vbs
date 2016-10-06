'connect to WMI namespace on local machine
Set objServices = GetObject("winmgmts:{impersonationLevel=impersonate}")
'get a reference to data file
Set objFile = objServices.Get("CIM_DataFile.Name='d:\data\report.doc'")
If objFile.TakeOwnership = 0 Then
 Wscript.Echo "File ownership successfully changed"
Else
 Wscript.Echo "File ownership transfer operation"
End If
