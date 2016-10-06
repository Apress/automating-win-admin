' create a session then log on
Set objSession = CreateObject("MAPI.Session")
' logon using the Anonymous account on server Odin
objSession.Logon , , , , , , "Odin" & vbLF & "Anonymous"
