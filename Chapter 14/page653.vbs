Set objWTS = CreateObject("WTSSupport.WTSLib")
'set user and domain controler
objWTS.UserName = "freds" 
objWTS.ServerName = "\\Odin" 

'allow user to to logon
objWTS.AllowWTSLogon = True

'set profile and home directories
objWTS.ProfilePath = "\\odin\profiles\freds"
objWTS.HomeDrive = "h:"
objWTS.HomeDirectory = "\\odin\freds$"

'set maximum session time to 2 hours
objWTS.ConnectionTimeout = 120
