'list event logs 
Dim objWMIObjects, objWMIObject

Set objService = GetObject("winmgmts:{impersonationLevel=impersonate}")
Set objWMIObjects = objService.InstancesOf("Win32_NTEventLogFile")

'loop through and list all available event logs
For Each objWMIObject In objWMIObjects
    WScript.Echo objWMIObject.LogFileName
Next
