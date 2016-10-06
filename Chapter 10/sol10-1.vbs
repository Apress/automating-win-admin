Dim objService, objWMIObject, objWMIObjects

'create an instance of a Services object for namespace root\cimv2
Set objServices = _
    GetObject("winmgmts:{impersonationLevel=impersonate}!root\cimv2")

'create a collection of Win32_ComputerSystem class objects
Set objWMIObjects = objServices.InstancesOf("Win32_ComputerSystem")

'enumerate collection.. there will only be one object, the local computer
For Each objWMIObject In objWMIObjects
    'display some information from the class
    WScript.Echo "Computer description:" & objWMIObject.Description
    WScript.Echo "Physical memory:" & objWMIObject.TotalPhysicalMemory
    WScript.Echo "Manufacturer:" & objWMIObject.Manufacturer
Next
