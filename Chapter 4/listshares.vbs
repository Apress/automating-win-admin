'listshares.vbs
'lists connected network shares
Dim objNetwork, objShares, nf
Set objNetwork = CreateObject("Wscript.Network")
Set objShares = objNetwork.EnumNetworkDrives
'output the first connected share and associated drive letter
Wscript.Echo "Drive " & objShares(0) & " is connected to " & objShares(1)

'loop through all connected shares and output details
For nf = 0 To objShares.Count - 1 Step 2
  Wscript.Echo objShares(nf) & " is connected to " & objShares(nf + 1)
Next
