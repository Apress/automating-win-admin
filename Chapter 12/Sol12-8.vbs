Const MsgTo = 1
Dim objSession, objRecipient 

Set objSession = CreateObject("MAPI.Session")
objSession.Logon  "Valid Profile"

Set objMessage = objSession.Outbox.Messages.Add
objMessage.Subject = "Test message" 
objMessage.Text= "This is the body of the message" 
'add Fred Smith as a recipient - resolve E-maile-mail address from Address book
Set objRecipient = objMessage.Recipients.Add("Fred Smith",, MsgTo)
objRecipient.Resolve
objMessage.Send
objSession.Logoff
