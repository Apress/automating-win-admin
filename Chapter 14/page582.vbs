'set unlock interval to 40 minutes (2400 seconds)
Set objDom = GetObject("WinNT://Acme")
objDom.AutoUnlockInterval = 2400
objDom.SetInfo

'set unlock interval to 40 minutes (2400000000000 nano-seconds)
Set objDom = GetObject("LDAP://DC=Acme,DC=COM")
objDom.lockoutDuration.LowPart = 1769803776
objDom.lockoutDuration.HighPart = -6
objDom.SetInfo
