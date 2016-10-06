Set objGroup = GetObject("WinNT://Acme/Domain Users,group")
'display name of each object in group
For Each objUser In objGroup.Members
    Wscript.Echo objUser.Name
Next
