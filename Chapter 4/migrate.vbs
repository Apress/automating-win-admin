'migrate.vbs
Dim objWshNetwork, nF
Dim sShare, sDrive, objEnumNetwork

Set objWshNetwork = CreateObject("WScript.Network")

Set objEnumNetwork = objWshNetwork.EnumNetworkDrives

For nF = 0 To objEnumNetwork.Count - 1 Step 2

    'get the drive letter and share name for the current share
    sShare = objEnumNetwork(nF)
    sDrive = objEnumNetwork(nF + 1)
    
    'check if the current share is connected to ODIN.
    If StrComp(Left(sShare, 7), "\\ODIN\", vbTextCompare) = 0 Then
        'remove the existing share
        objWshNetwork.RemoveNetworkDrive sDrive, True, True
        
        'remap the drive to the new server
        objWshNetwork.MapNetworkDrive sDrive, _
                                                   "\\THOR\" & Mid(sShare, 8), True
       
    End If
Next
