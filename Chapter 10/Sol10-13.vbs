Dim nStatus, objService, objWMIObject, objWMIObjects

'create process on remote machine
Set objService = _
 GetObject("winmgmts:{impersonationLevel=impersonate,(Shutdown)}!\\thor")

'get the active OS
Set objWMIObjects = objService.ExecQuery _
          ("Select * From Win32_OperatingSystem Where Primary = True")

For Each objOS In objWMIObjects
    Set objWMIObject = objOS
Next
nStatus = objWMIObject.Reboot
