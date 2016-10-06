'servicechng.vbs
'changes the password for all services with specific user ID
Dim objWMIObjects, objWMIObject, objService, nResult

'get a reference to local WMI service
Set objService = _
    GetObject("winmgmts:{impersonationLevel=impersonate}")

'get all services that use the Administrator account 
Set objWMIObjects = objService.ExecQuery( _
                                    "Select * From Win32_Service " & _ 
                                    Where StartName='Acme\\Administrator'")

'loop through each service and change the password
For Each objWMIObject In objWMIObjects

    nResult = objWMIObject.Change(, , , , , , , "newpassword")

    If nResult = 0 Then
         Wscript.Echo "Password successfully set for " & _ 
                       objWMIObject.DisplayName
    Else
        Wscript.Echo "Error setting password set for service " _ 
                       &  objWMIObject.DisplayName 
    End If

Next
