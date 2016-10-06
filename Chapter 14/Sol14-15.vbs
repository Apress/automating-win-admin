Set objNetwork = Wscript.Create("Wscript.Network")

Set objGroup= GetObject("WinNT://Acme/Accounting Users,group")
'
If (objGroup.IsMember("WinNT://ACME/" & objNetwork.UserName)) Then 
   'connect to printer
End If
