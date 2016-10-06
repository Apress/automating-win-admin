Dim objGroup

'get the group to add objects to..
Set objGroup = GetObject("WinNT://Acme/Acctusers")

objGroup.add "WinNT://Acme/freds,user" 'add a user
objGroup.add "WinNT://Acme/joeb" 'add a another user
'add a group – can only add other groups to Local groups, not Global 
objGroup.add "WinNT://Acme/finance,group"
