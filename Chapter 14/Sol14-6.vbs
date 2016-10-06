'get a reference to user 
Set objUser = GetObject("WinNT://Acme/freds,user")

'set fullname and description properties
objUser.FullName = "Fred Smith"
objUser.Put "Description", "Accounting Manager"
objUser.SetInfo
