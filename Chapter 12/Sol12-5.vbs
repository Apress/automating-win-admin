Dim objMessage 
Dim objSession Set objSession = CreateObject("MAPI.Session")
objSession.Logon 
'create a new message, setting the subject to Hello There
Set objMessage = objSession.Outbox.Messages.Add("Hello There")
objMessage.Text = "this is the body of the message"
' perform operations and log off…
