'logon.vbs
Dim objNetwork, strName
Set objNetwork = CreateObject("Wscript.Network")

On Error Resume Next
  'get user name
While strName = ""
    strName = objNetwork.UserName
    Wscript.Sleep 100
WEnd
'map user to home drive
objNetwork.MapNetworkDrive "H:", 
              "\\THOR\" & strName & "$", True
'connect to public area
objNetwork.MapNetworkDrive "P:", "\\THOR\PublicArea", True
'connect to apps share
objNetwork.MapNetworkDrive "W:", "\\THOR\Apps", True
