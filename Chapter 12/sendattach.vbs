'sendattach.vbs
Const MsgFileData =1
Const MsgFileLink = 2
Const MsgOLE = 3 
Dim objSession, objMessage, objRecipient '
Set objSession = CreateObject("MAPI.Session")
objSession.Logon "Valid profile" 
Set objMessage = objSession.Outbox.Messages.Add("Attachment Test")

objMessage.Text= "This is the body of the message" 
Set objRecipient = _
                    objMessage.Recipients.Add("Joe Blow ","SMTP:administrator@acme.com")
objRecipient.Resolve
Set objAttachment = objMessage.Attachments.Add("Attached File", , _
                    MsgFileData,"c:\data\Weekly.xls")

Set objAttachment = objMessage.Attachments.Add("Linked File", , _
                    MsgFileLink, "\\odin\xldata\Weekly.xls")

Set objAttachment = objMessage.Attachments.Add("Embedded File", , _ 
                    MsgOLE, "c:\data\Weekly.xls")
objMessage.Update
objMessage.Send 
objSession.Logoff
