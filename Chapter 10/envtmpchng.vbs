'envtmpchng.vbs
'changes the location of temporary file environment 
'variables Tmp and Temp
Dim objService
Dim objWMIObject, objWMIObjects

'get a reference to a WMI service
Set objService = _
    GetObject("winmgmts:{impersonationLevel=impersonate}")

Set objWMIObjects = objService.InstancesOf("Win32_Environment")

On Error Resume Next

'loop through each environment variable for temp
For Each objWMIObject In objWMIObjects
'
 If objWMIObject.Name = "TEMP" Or objWMIObject.Name = "TMP" Then
    Debug.Print objWMIObject.Path_, objWMIObject.VariableValue
    objWMIObject.VariableValue = "d:\temp"
    
    'update the settings
    Call objWMIObject.Put_
 End If

Next
