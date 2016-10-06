Const MsgTo = 1
Const MsgCC = 2
Const MsgBCC = 3

Dim objSession, objMessage, objRecipient 
Set objSession = CreateObject("MAPI.Session")
objSession.Logon  "Valid Profile"

Set objMessage = objSession.Outbox.Messages.Add

objMessage.Subject = "Test message" 
objMessage.Text= "This is the body of the message" 
Set objRecipient = objMessage.Recipients.Add("Fred Smith",, MsgTo)
objRecipient.Resolve
'carbon copy Joe Blow 
Set objRecipient =objMessage.Recipients.Add("Joe B","SMTP:joeb@abc.com", MsgCC)
objRecipient.Resolve objMessage.Send
objSession.Logoff
