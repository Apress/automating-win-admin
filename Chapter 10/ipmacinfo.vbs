'IPMacInfo.vbs
'list IP and MAC information for local computer
Dim objServices, objWMIObjects, objWMIObject, nF

'create Services object for default namespace on local computer
Set objServices = GetObject("winmgmts:{impersonationLevel=impersonate}")

'get all instances of IP enabled devices for 
'Win32_NetworkAdapterConfiguration class
Set objWMIObjects = _
                      objServices.ExecQuery( _
                      "Select * From Win32_NetworkAdapterConfiguration" _
                       & " Where IPEnabled = True")

'enumerate each Win32_NetworkAdapterConfiguration instance
For Each objWMIObject In objWMIObjects
   WScript.Echo objWMIObject.Caption & " has the MAC address " _
                                 & objWMIObject.MACAddress
    'make sure array is not empty
    If Not IsNull(objWMIObject.IPAddress) Then
     'list all associated IP addresses with adapter
     For nF = 0 To UBound(objWMIObject.IPAddress)
        WScript.Echo "   " & objWMIObject.IPAddress(nF)
     Next
    End If
Next
