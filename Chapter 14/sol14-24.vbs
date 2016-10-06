'set WTS properties for user Fred Smith
Set objUser = GetObject("LDAP://CN=fred smith,CN=Users,DC=Acme,DC=com")

'allow user to to logon
objUser.AllowLogon = 1
'set profile and home directories
objUser.TerminalServicesProfilePath = "\\odin\profiles\freds"
objUser.TerminalServicesHomeDrive = "h:"
objUser.TerminalServicesHomeDirectory = "\\odin\freds$"
'set maximum session time to 10 hours
objUser.MaxConnectionTime = 600
objUser.SetInfo 
