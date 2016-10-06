Dim objSession 
Dim sMsg
' create a session then log on, supplying username and password
Set objSession = CreateObject("MAPI.Session")
' change the parameters to valid values for your configuration
objSession.Logon 
'retrieve some Session properties available after a valid logon
sMsg = "OS: " & objSession.OperatingSystem & vbCr ' use vbCr
sMsg =  sMsg & "User Name:" & objSession.CurrentUser & vbCr
sMsg = sMsg & "MAPI Version:" & objSession.Version & vbCr 
sMsg = sMsg  & "Profile Name:" &  objSession.Name
Wscript.Echo sMsg
objSession.Logoff
