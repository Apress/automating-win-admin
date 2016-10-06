Dim objDomain

'get a reference to the Acme domain
Set objDomain = GetObject("WinNT://ACME")
objDomain. MinPasswordLength = 6

objDomain.SetInfo
