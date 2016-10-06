'reboot.vbs
'reboots local or remote computer
 Dim objOS, nStatus, objService, objWMIObject, objWMIObjects
 Dim strMachine
 On Error Resume Next
 strMachine =  ""
 'check if argument passed, if so assume remote computer 
 'to reboot
 If Wscript.Arguments.Count = 1 Then 
     strMachine =  "!\\" & Wscript.Arguments(0)
 End If

 'create instance of WMI service with Shutdown privilege
 Set objService = GetObject( _
                "winmgmts:{impersonationLevel=impersonate,(Shutdown)}" _
                 & strMachine )

 If Err Then
   Wscript.Echo "Error getting reference to computer " & strMachine
 End If

 'get the primary O/S to reboot
 Set objWMIObjects = objService.ExecQuery( _ 
             "Select * From Win32_OperatingSystem Where Primary = True")

 'get the instance of the O/S to reboot
 For Each objOS In objWMIObjects
    Set objWMIInstance = objOS
 Next

 nStatus = objWMIInstance.Reboot() 

 If nStatus = 0 Then
   Wscript.Echo "Computer successfully reboot"
 Else
   Wscript.Echo "Error rebooting computer: "  & nStatus
 End If
 