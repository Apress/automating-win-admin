'bind to a domain
Set objDomain = GetObject("WinNT://ACME")
'create a new user – Fred Smith
Set objUser = objDomain.Create("User", "FredS")
objUser.SetPassword("iu12yt09")
objUser.SetInfo
