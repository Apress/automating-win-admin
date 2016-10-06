'osrecover.vbs
Dim objServices
Dim objWMIObject, objWMIObjects

'create an instance of a Services object for the local machine
Set objServices = _
    GetObject("winmgmts:{impersonationLevel=impersonate}!root\cimv2")

'create an instance of the Win32_OSRecoveryOption class
Set objWMIObjects = objServices.InstancesOf _
                        ("Win32_OSRecoveryConfiguration")

'loop through each object (there will be only one)
For Each objWMIObject In objWMIObjects
    'set the DebugFilePath property
    objWMIObject.DebugFilePath = "d:\MEMORY.DMP"

    'update the settings
    Call objWMIObject.Put_
Next
