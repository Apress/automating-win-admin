Dim objDomain
'get a reference to the Acme domain
Set objDomain = GetObject("WinNT://Acme")
'delete a user
objDomain.Delete "user", "Freds"
