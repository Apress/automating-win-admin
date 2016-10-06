Dim objSession ' Session object
' create a MAPI session, then log on
Set objSession = CreateObject("MAPI.Session")
'attempt to logon. Since parameters are omitted, you will be prompted 
'for a valid mail profile
objSession.Logon 
