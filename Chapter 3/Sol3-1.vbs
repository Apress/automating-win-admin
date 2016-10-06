Dim objNetwork, strUser
Set objNetwork = CreateObject("WScript.Network")
strUser =""

'get logged on user name 
' user ID is returned correctly on Win 9x/ME computers
  Do While strUser =""
 strUser = objNetwork.UserName

Loop

'map user to home drive – assumes home drive share is combination of
'user-id and $ sign (hidden share)
objNetwork.MapNetworkDrive "H:", _
       "\\THOR\" & strUser & "$" , True
'connect to public area
objNetwork.MapNetworkDrive "P:", _
             "\\THOR\PublicArea", True
